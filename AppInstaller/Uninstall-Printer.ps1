<#
.SYNOPSIS
    Uninstalls a network printer (queue + port + driver + Driver Store) deployed via Intune Win32 app.

.DESCRIPTION
    Removes the printer queue, TCP/IP port, printer driver, and optionally the OEM INF package
    from the Windows Driver Store.

    Safe-guards applied before removal:
      - The printer queue's port name is captured before deletion so the port can be cleaned up.
      - The port is only removed if no other printer queues reference it.
      - The driver is only removed if no other printer queues use it.
      - Any remaining printer queues using the driver are deleted before driver removal.
      - The OEM INF entry in the Driver Store is located and deleted with pnputil.

    Designed to run as SYSTEM under the Intune Management Extension (IME).

.PARAMETER PrinterName
    The display name of the printer queue to remove (e.g. "Brother MFC-J6540DW").

.PARAMETER DriverName
    The exact driver name as it appears in the Print Spooler / Driver Store
    (e.g. "Brother MFC-J6540DW Printer"). Must match the value used during install.

.PARAMETER KeepDriver
    When specified, the printer driver is kept in the Print Spooler after removing
    the printer queue. By default the driver is removed.

.PARAMETER KeepPort
    When specified, the TCP/IP port is kept after removing the printer queue.
    By default the port is removed.

.PARAMETER KeepDriverStore
    When specified, the OEM INF package is kept in the Windows Driver Store.
    By default it is removed with pnputil /delete-driver. Use this switch when
    multiple printer models share the same INF.

.PARAMETER SpoolerTimeoutSeconds
    Maximum seconds to wait for the Print Spooler service before aborting. Default: 60.

.EXAMPLE
    powershell.exe -ExecutionPolicy Bypass -File Uninstall-Printer.ps1 `
        -PrinterName "Brother MFC-J6540DW" `
        -DriverName "Brother MFC-J6540DW Printer"

.EXAMPLE
    # Keep the driver and port (remove only the printer queue)
    powershell.exe -ExecutionPolicy Bypass -File Uninstall-Printer.ps1 `
        -PrinterName "Brother MFC-J6540DW" `
        -DriverName "Brother MFC-J6540DW Printer" `
        -KeepDriver -KeepPort

.NOTES
    Exit codes:
        0   Success (or components were already absent - idempotent)
        1   Failure (see log for details)

    Log: %ProgramData%\Microsoft\IntuneManagementExtension\Logs\Uninstall-Printer.log

    Version history:
        1.4.2  2026-03-27  Demote pnputil block-match fallback log from WARN to INFO: pnputil
                           /enum-drivers never includes the Spooler driver friendly name so the
                           FileRepository scan is the expected code path for printer drivers.
        1.4.1  2026-03-27  Fix [uint32] HResult cast overflow in Write-ExceptionDetail and driver
                           removal catch block: [uint32] throws OverflowException on negative HResult
                           values. Replaced with [int32] (same fix as Install-Printer.ps1 v1.6.2/1.6.3).
        1.4.0  2026-03-25  Extended diagnostic logging: current-step tracker, rich exception detail
                           (type/inner/HResult/position), PrintService/Admin event log on failures,
                           registry key confirmation after each removal step.
        1.3.0  2026-03-25  Removed 64-bit re-launch block entirely; print management cmdlets work
                           from 32-bit PowerShell on Windows 10/11. pnputil resolved via sysnative
                           path only.
        1.2.0  2026-03-25  Expanded diagnostic logging; added per-step timing, printer detail
                           dump before removal, and script versioning.
        1.1.0  2026-03-25  Fixed 64-bit re-launch argument quoting; moved logging before re-launch.
        1.0.0              Initial release.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string] $PrinterName,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string] $DriverName,

    [Parameter(Mandatory = $false)]
    [switch] $KeepDriver,

    [Parameter(Mandatory = $false)]
    [switch] $KeepPort,

    [Parameter(Mandatory = $false)]
    [switch] $KeepDriverStore,

    [Parameter(Mandatory = $false)]
    [ValidateRange(10, 300)]
    [int] $SpoolerTimeoutSeconds = 60
)

$ScriptVersion = '1.4.2'

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Convert switch params to boolean removal flags
$RemoveDriver          = -not $KeepDriver
$RemovePort            = -not $KeepPort
$RemoveFromDriverStore = -not $KeepDriverStore

#region -- Logging ----
$LogDir  = Join-Path $env:ProgramData 'Microsoft\IntuneManagementExtension\Logs'
$LogFile = Join-Path $LogDir 'Uninstall-Printer.log'

if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
}

# Append separator instead of clearing - preserves history across retries
if (Test-Path $LogFile) {
    Add-Content -Path $LogFile -Value ("`n" + ('=' * 80) + "`n[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] NEW RUN`n" + ('=' * 80)) -Encoding UTF8
}

function Write-Log {
    param(
        [Parameter(Mandatory = $true)][string] $Message,
        [ValidateSet('INFO','WARN','ERROR')][string] $Level = 'INFO'
    )
    $line = '[{0}] [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    switch ($Level) {
        'ERROR' { Write-Error   $line -ErrorAction Continue }
        'WARN'  { Write-Warning $line }
        default { Write-Host    $line }
    }
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
}
#endregion

#region -- Helpers ----
function Resolve-PnpUtil {
    $candidates = @(
        (Join-Path $env:windir 'sysnative\pnputil.exe'),
        (Join-Path $env:windir 'System32\pnputil.exe')
    )
    $found = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $found) { throw 'pnputil.exe not found on this system.' }
    return $found
}

function Wait-ForSpooler {
    param([int] $TimeoutSeconds)
    Write-Log 'Ensuring Print Spooler service is running...'
    $svc = Get-Service -Name Spooler -ErrorAction SilentlyContinue
    if (-not $svc) { throw 'Print Spooler service not found.' }

    if ($svc.Status -ne 'Running') {
        Write-Log "Spooler is '$($svc.Status)' - attempting to start..." 'WARN'
        Start-Service -Name Spooler -ErrorAction SilentlyContinue
    }

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        $svc.Refresh()
        if ($svc.Status -eq 'Running') { Write-Log 'Spooler is Running.'; return }
        Start-Sleep -Seconds 2
    }
    throw "Spooler did not reach Running state within $TimeoutSeconds seconds."
}

function Write-PrintServiceEvents {
    param([int] $LastSeconds = 60)
    $since = (Get-Date).AddSeconds(-$LastSeconds)
    $logs  = @('Microsoft-Windows-PrintService/Admin', 'Microsoft-Windows-PrintService/Operational')
    foreach ($logName in $logs) {
        try {
            $events = Get-WinEvent -FilterHashtable @{
                LogName   = $logName
                Level     = @(1, 2, 3)
                StartTime = $since
            } -MaxEvents 10 -ErrorAction Stop
            if ($events) {
                Write-Log "--- $logName (last ${LastSeconds}s, errors/warnings) ---"
                foreach ($ev in $events) {
                    $msg = $ev.Message -replace '\r?\n', ' '
                    if ($msg.Length -gt 300) { $msg = $msg.Substring(0, 300) + '...' }
                    Write-Log "  [$($ev.TimeCreated.ToString('HH:mm:ss'))] ID=$($ev.Id) $($ev.LevelDisplayName): $msg"
                }
            }
        } catch {
            Write-Log "  ($logName not available: $($_.Exception.Message))"
        }
    }
}

function Write-ExceptionDetail {
    param([System.Management.Automation.ErrorRecord] $ErrorRecord)
    $ex = $ErrorRecord.Exception
    Write-Log "Exception type    : $($ex.GetType().FullName)" 'ERROR'
    Write-Log "Exception message : $($ex.Message)" 'ERROR'
    $depth = 1
    $inner = $ex.InnerException
    while ($null -ne $inner -and $depth -le 3) {
        Write-Log "  Inner[$depth] type    : $($inner.GetType().FullName)" 'ERROR'
        Write-Log "  Inner[$depth] message : $($inner.Message)" 'ERROR'
        $inner = $inner.InnerException
        $depth++
    }
    if ($ex.HResult -ne 0) {
        $hresult = '0x{0:X8}' -f [int32]$ex.HResult
        Write-Log "HResult           : $hresult  (decimal $($ex.HResult))" 'ERROR'
    }
    Write-Log "Script stack      : $($ErrorRecord.ScriptStackTrace)" 'ERROR'
    if ($ErrorRecord.InvocationInfo -and $ErrorRecord.InvocationInfo.PositionMessage) {
        Write-Log "Position          : $($ErrorRecord.InvocationInfo.PositionMessage.Trim())" 'ERROR'
    }
}

function Find-OemInfsForDriver {
    param([string] $PnpUtilPath, [string] $Name)

    Write-Log "Searching Driver Store for OEM INF(s) matching driver '$Name'..."
    $enumOutput = & $PnpUtilPath /enum-drivers 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-Log "pnputil /enum-drivers returned exit code $LASTEXITCODE" 'WARN'
        return @()
    }

    $blocks  = ($enumOutput -join "`n") -split '(?m)^\s*$' | Where-Object { $_.Trim() }
    $results = [System.Collections.Generic.List[string]]::new()

    foreach ($block in $blocks) {
        $oemLine = ($block -split "`n") | Where-Object { $_ -match 'Published Name' } |
                   Select-Object -First 1
        if (-not $oemLine) { continue }

        if ($block -match [regex]::Escape($Name)) {
            if ($oemLine -match ':\s*(oem\d+\.inf)') {
                $oemInf = $Matches[1].Trim()
                Write-Log "Found OEM INF '$oemInf' for driver '$Name'."
                $results.Add($oemInf)
            }
        }
    }

    if ($results.Count -eq 0) {
        Write-Log 'pnputil block match found nothing (expected for printer drivers) - trying FileRepository INF scan...'
        $repoPath = Join-Path $env:windir 'System32\DriverStore\FileRepository'
        if (Test-Path $repoPath) {
            $matchingInfs = Get-ChildItem -Path $repoPath -Recurse -Filter '*.inf' -ErrorAction SilentlyContinue |
                    Select-String -Pattern $Name -SimpleMatch -List -ErrorAction SilentlyContinue |
                    Select-Object -ExpandProperty Path
            foreach ($infFile in $matchingInfs) {
                $originalName = [System.IO.Path]::GetFileName($infFile)
                foreach ($block in $blocks) {
                    if ($block -match "Original Name:\s*$([regex]::Escape($originalName))") {
                        $oemLine = ($block -split "`n") | Where-Object { $_ -match 'Published Name' } |
                                   Select-Object -First 1
                        if ($oemLine -match ':\s*(oem\d+\.inf)') {
                            $oemInf = $Matches[1].Trim()
                            if (-not $results.Contains($oemInf)) {
                                Write-Log "Found OEM INF '$oemInf' via FileRepository fallback."
                                $results.Add($oemInf)
                            }
                        }
                    }
                }
            }
        }
    }

    if ($results.Count -eq 0) {
        Write-Log "No OEM INF found in Driver Store for driver '$Name'." 'WARN'
    }
    return $results.ToArray()
}
#endregion

#region -- Main ----
$stopwatch   = $null
$CurrentStep = 'startup'

try {
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    Write-Log '==== Uninstall-Printer START ===='
    Write-Log "PrinterName          : $PrinterName"
    Write-Log "DriverName           : $DriverName"
    Write-Log "RemoveDriver         : $RemoveDriver"
    Write-Log "RemovePort           : $RemovePort"
    Write-Log "RemoveFromDriverStore: $RemoveFromDriverStore"
    Write-Log "Script version       : $ScriptVersion"
    Write-Log "Running as           : $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"
    Write-Log "Computer name        : $env:COMPUTERNAME"
    Write-Log "OS arch              : $env:PROCESSOR_ARCHITECTURE"
    Write-Log "OS version           : $([System.Environment]::OSVersion.VersionString)"
    Write-Log "PS version           : $($PSVersionTable.PSVersion)"
    Write-Log "PS 64-bit            : $([Environment]::Is64BitProcess)"
    Write-Log "Process ID           : $([System.Diagnostics.Process]::GetCurrentProcess().Id)"

    $pnputil = Resolve-PnpUtil
    Write-Log "Using pnputil        : $pnputil"

    # -- Step 1: Ensure Spooler is running ----
    $CurrentStep = 'Step 1: Ensure Spooler is running'
    Wait-ForSpooler -TimeoutSeconds $SpoolerTimeoutSeconds

    # -- Step 2: Locate printer queue and capture its port name ----
    $CurrentStep = 'Step 2: Locate printer queue'
    $capturedPortName = $null
    $printer = Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue

    if (-not $printer) {
        Write-Log "Printer queue '$PrinterName' not found - nothing to remove."
    } else {
        $capturedPortName = $printer.PortName
        Write-Log "Printer queue '$PrinterName' found (port='$capturedPortName', driver='$($printer.DriverName)')."
        Write-Log "Printer details: Status='$($printer.PrinterStatus)'  Type='$($printer.Type)'  Shared=$($printer.Shared)  Published=$($printer.Published)"

        $portDetail = Get-PrinterPort -Name $capturedPortName -ErrorAction SilentlyContinue
        if ($portDetail) {
            Write-Log "Port details: Address='$($portDetail.PrinterHostAddress)'  Port=$($portDetail.PortNumber)  Protocol=$($portDetail.Protocol)"
        }
        $driverDetail = Get-PrinterDriver -Name $printer.DriverName -ErrorAction SilentlyContinue
        if ($driverDetail) {
            Write-Log "Driver details: Manufacturer='$($driverDetail.Manufacturer)'  Version='$($driverDetail.DriverVersion)'  Environment='$($driverDetail.PrinterEnvironment)'"
        }

        # -- Step 3: Remove printer queue ----
        $CurrentStep = 'Step 3: Remove printer queue'
        Write-Log "Removing printer queue '$PrinterName'..."
        Remove-Printer -Name $PrinterName -ErrorAction Stop
        Write-Log "Printer queue '$PrinterName' removed."
        $prtRegKey = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Printers\$PrinterName"
        Write-Log "Printer registry key: $(if (Test-Path $prtRegKey) { 'STILL EXISTS (unexpected)' } else { 'GONE (confirmed)' })"
        Write-Log "Step 3 elapsed: $($stopwatch.Elapsed.TotalSeconds.ToString('F1'))s"
    }

    # -- Step 4: Remove printer port ----
    $CurrentStep = 'Step 4: Remove printer port'
    if ($RemovePort -and $capturedPortName) {
        $otherPrinters = Get-Printer -ErrorAction SilentlyContinue |
                    Where-Object { $_.PortName -eq $capturedPortName }
        if ($otherPrinters) {
            $names = ($otherPrinters | Select-Object -ExpandProperty Name) -join ', '
            Write-Log "Port '$capturedPortName' is still used by: $names - skipping removal." 'WARN'
        } else {
            $portExists = $null -ne (Get-PrinterPort -Name $capturedPortName -ErrorAction SilentlyContinue)
            if ($portExists) {
                Write-Log "Removing port '$capturedPortName'..."
                Remove-PrinterPort -Name $capturedPortName -Confirm:$false -ErrorAction Stop
                Write-Log "Port '$capturedPortName' removed."
            } else {
                Write-Log "Port '$capturedPortName' not found - already removed."
            }
        }
    } elseif (-not $RemovePort) {
        Write-Log 'RemovePort is false - skipping port removal.'
    } elseif (-not $capturedPortName) {
        Write-Log 'No port name captured (printer was not found) - skipping port removal.'
    }
    Write-Log "Step 4 elapsed: $($stopwatch.Elapsed.TotalSeconds.ToString('F1'))s"

    # -- Step 5: Remove driver from Print Spooler ----
    $CurrentStep = 'Step 5: Remove driver from Spooler'
    if ($RemoveDriver) {
        $driverPresent = $null -ne (Get-PrinterDriver -Name $DriverName -ErrorAction SilentlyContinue)

        if (-not $driverPresent) {
            Write-Log "Driver '$DriverName' is not registered with the Print Spooler - nothing to remove."
        } else {
            $otherQueues = Get-Printer -ErrorAction SilentlyContinue |
                    Where-Object { $_.DriverName -eq $DriverName }
            if ($otherQueues) {
                $names = ($otherQueues | Select-Object -ExpandProperty Name) -join ', '
                Write-Log "Driver '$DriverName' is still used by: $names - removing those queues first..." 'WARN'
                foreach ($q in $otherQueues) {
                    Write-Log "Removing printer queue '$($q.Name)'..." 'WARN'
                    Remove-Printer -Name $q.Name -ErrorAction SilentlyContinue
                }
            }

            Write-Log "Removing driver '$DriverName' from Print Spooler..."
            try {
                Remove-PrinterDriver -Name $DriverName -ErrorAction Stop
                Write-Log 'Remove-PrinterDriver succeeded.'
            } catch {
                Write-Log "Remove-PrinterDriver error: $($_.Exception.Message)" 'WARN'
                Write-Log "  Exception type : $($_.Exception.GetType().FullName)" 'WARN'
                if ($_.Exception.HResult -ne 0) {
                    Write-Log "  HResult        : $('0x{0:X8}' -f [int32]$_.Exception.HResult)  (decimal $($_.Exception.HResult))" 'WARN'
                }
                Write-Log '  PrintService event log (last 30s):' 'WARN'
                Write-PrintServiceEvents -LastSeconds 30
                Write-Log 'Attempting printui fallback...'
                $puArgs = "printui.dll,PrintUIEntry /dd /m `"$DriverName`" /v 3"
                Start-Process -FilePath 'rundll32.exe' -ArgumentList $puArgs `
                    -Wait -NoNewWindow -ErrorAction SilentlyContinue
            }

            if ($null -ne (Get-PrinterDriver -Name $DriverName -ErrorAction SilentlyContinue)) {
                Write-Log "Driver '$DriverName' still present in Spooler after removal attempt." 'WARN'
            } else {
                $driverArch  = if ([Environment]::Is64BitOperatingSystem) { 'Windows x64' } else { 'Windows NT x86' }
                $drvRegKey   = "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Environments\$driverArch\Drivers\Version-3\$DriverName"
                Write-Log "Driver '$DriverName' removed from Print Spooler."
                Write-Log "Driver registry key: $(if (Test-Path $drvRegKey) { 'STILL EXISTS (unexpected)' } else { 'GONE (confirmed)' })"
            }
        }

        # -- Step 6: Remove OEM INF(s) from the Driver Store ----
        $CurrentStep = 'Step 6: Remove OEM INF from Driver Store'
        if ($RemoveFromDriverStore) {
            $oemInfs = @(Find-OemInfsForDriver -PnpUtilPath $pnputil -Name $DriverName)

            if ($oemInfs.Count -eq 0) {
                Write-Log 'Skipping Driver Store removal - no OEM INF found.' 'WARN'
            } else {
                # Shared INF safety check: get all other registered drivers to see if they share this INF
                $allOtherDrivers = Get-PrinterDriver | Where-Object { $_.Name -ne $DriverName }
                
                foreach ($oemInf in $oemInfs) {
                    $sharingDrivers = $allOtherDrivers | Where-Object { 
                        $_.InfPath -like "*\$oemInf" -or $_.InfPath -eq $oemInf 
                    }

                    if ($sharingDrivers) {
                        $driverList = ($sharingDrivers | Select-Object -ExpandProperty Name) -join ', '
                        Write-Log "OEM package '$oemInf' is still in use by other drivers: [$driverList]. Skipping Driver Store removal for safety." 'WARN'
                        continue
                    }

                    Write-Log "Removing OEM package '$oemInf' from Driver Store..."
                    $delArgs   = @('/delete-driver', $oemInf, '/uninstall', '/force')
                    $delOutput = & $pnputil @delArgs 2>&1
                    $delExit   = $LASTEXITCODE

                    if ($delOutput) { Write-Log ($delOutput -join [Environment]::NewLine) }
                    Write-Log "pnputil /delete-driver exit code: $delExit"

                    if ($delExit -notin 0, 2) {
                        Write-Log "pnputil /delete-driver returned unexpected exit code $delExit." 'WARN'
                    } else {
                        Write-Log "OEM package '$oemInf' removed from Driver Store."
                    }
                }
            }
        } else {
            Write-Log 'RemoveFromDriverStore is false - skipping Driver Store cleanup.'
        }
    } else {
        Write-Log 'RemoveDriver is false - skipping driver and Driver Store removal.'
    }

    $stopwatch.Stop()
    Write-Log "Elapsed time         : $($stopwatch.Elapsed.TotalSeconds.ToString('F1')) seconds"
    Write-Log '==== Uninstall-Printer END (SUCCESS) ===='
    exit 0

} catch {
    Write-Log "Failed at         : $CurrentStep" 'ERROR'
    Write-ExceptionDetail -ErrorRecord $_
    Write-Log '--- PrintService event log (last 60s) ---' 'ERROR'
    Write-PrintServiceEvents -LastSeconds 60
    if ($stopwatch) { Write-Log "Elapsed time: $($stopwatch.Elapsed.TotalSeconds.ToString('F1')) seconds" }
    Write-Log '==== Uninstall-Printer END (FAILURE) ===='
    exit 1
}
#endregion
