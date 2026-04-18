<#
.SYNOPSIS
    Installs a network printer (driver + TCP/IP port + queue + configuration) via Intune Win32 app.

.DESCRIPTION
    Stages the printer driver into the Windows Driver Store using pnputil, restarts the Print
    Spooler to process the staged driver (critical for complex drivers with CoInstallers such as
    Brother), registers the driver with the Print Spooler with a retry loop, creates a Standard
    TCP/IP port, creates a printer queue, and optionally configures printer settings.

    Designed to run as SYSTEM under the Intune Management Extension (IME).

.PARAMETER DriverName
    The exact driver name as declared in the INF file (e.g. "Brother MFC-J6540DW Printer").
    Used to register and verify the driver with the Print Spooler after staging.

.PARAMETER InfPath
    Path to the driver INF file. Accepts an absolute path or a path relative to the
    script's own directory ($PSScriptRoot). If omitted, the script searches for any
    .inf file under its own directory.

.PARAMETER DriverSourceFolder
    The folder that contains the INF and all accompanying driver files (SYS, DLL, CAT, etc.).
    Defaults to the folder that contains InfPath.

.PARAMETER SpoolerTimeoutSeconds
    Maximum seconds to wait for the Print Spooler service. Default: 60.

.PARAMETER PrinterName
    Display name for the printer queue (e.g. "Office Printer 3rd Floor").

.PARAMETER PrinterIPAddress
    IP address of the network printer (e.g. "192.168.1.100").

.PARAMETER PortName
    Custom TCP/IP port name. Defaults to "IP_<PrinterIPAddress>".

.PARAMETER PortNumber
    TCP port number for the printer port. Default: 9100.

.PARAMETER PaperSize
    Default paper size for the printer. Applied via PrintTicket XML modification.
    Supported values: A4, A5, A3, Letter, Legal, Executive, Tabloid, Statement, B4, B5,
    Folio, Quarto, Note.

.PARAMETER DuplexMode
    Duplex printing mode: OneSided, TwoSidedLongEdge, or TwoSidedShortEdge.

.PARAMETER ColorMode
    Color printing mode: Color or Monochrome.

.PARAMETER Collate
    Collation setting: Enabled or Disabled.

.PARAMETER PrintTicketXml
    Raw PrintTicket XML string for advanced driver-specific settings not covered by the
    standard parameters. Applied LAST, after all standard settings.

.PARAMETER Location
    Physical location text for the printer.

.PARAMETER Comment
    Description text for the printer.

.PARAMETER Shared
    When specified, shares the printer on the network.

.PARAMETER ShareName
    Network share name for the printer (used with -Shared).

.EXAMPLE
    powershell.exe -ExecutionPolicy Bypass -File Install-Printer.ps1 `
        -DriverName "Brother MFC-J6540DW Printer" `
        -InfPath ".\Driver\BRPRI20A.INF" `
        -PrinterName "Brother MFC-J6540DW" `
        -PrinterIPAddress "192.168.1.100"

.EXAMPLE
    powershell.exe -ExecutionPolicy Bypass -File Install-Printer.ps1 `
        -DriverName "Brother MFC-J6540DW Printer" `
        -InfPath ".\Driver\BRPRI20A.INF" `
        -PrinterName "Brother MFC-J6540DW" `
        -PrinterIPAddress "192.168.1.100" `
        -PaperSize A4 -DuplexMode TwoSidedLongEdge -ColorMode Color

.NOTES
    Exit codes:
        0   Success
        1   Failure (see log for details)

    Log: %ProgramData%\Microsoft\IntuneManagementExtension\Logs\Install-Printer.log

    Version history:
        1.7.1  2026-04-15  Fix Add-PrinterPort -SNMPEnabled ParameterBindingException on Windows 11
                           24H2+: the parameter is no longer supported by Add-PrinterPort in any PS
                           bitness. Remove -SNMPEnabled from portArgs entirely; apply SNMPEnabled
                           unconditionally via the Standard TCP/IP Port monitor registry key after
                           port creation (already done for 32-bit PS since v1.6.8/v1.6.9).
        1.7.0  2026-03-28  Add -SNMPCommunity parameter: writes community string to the Standard TCP/IP
                           Port monitor registry key after port creation (Add-PrinterPort does not expose
                           this parameter in any PS bitness; default 'public' was previously hardcoded by
                           Windows). Add WPP proactive startup check: reads EnableWindowsProtectedPrint
                           at script start and logs a WARN with the exact Intune policy name before any
                           driver work begins -- surfaces the blocker immediately instead of after 5
                           failed Add-PrinterDriver attempts. Add ARM64 WHQL signing warning: logs a WARN
                           at startup when running on an ARM64 device because WHQL signing is required
                           (EV Code Signing alone is insufficient and causes error 0xE0000242 on ARM64).
        1.6.10 2026-03-27  Fix port details log showing SNMP=False on 32-bit PS: Get-PrinterPort does
                           not return SNMPEnabled correctly in 32-bit PS. Read the value directly from
                           the registry for the verification log line when running as a 32-bit process.
        1.6.9  2026-03-27  Eliminate Add-PrinterPort SNMP warning: check Is64BitProcess upfront and
                           omit -SNMPEnabled from portArgs on 32-bit PS instead of catching a
                           ParameterBindingException after the fact. Removes the try/catch fallback
                           and the WARN log line entirely on 32-bit PS (IME/manual 32-bit context).
        1.6.8  2026-03-27  Fix SNMP registry write: Set-PrinterPort does not exist in 32-bit PS.
                           Write SNMPEnabled directly to the Standard TCP/IP Port monitor registry
                           key after port creation when Add-PrinterPort fallback was triggered.
        1.6.7  2026-03-27  Fix SNMP not applied when Add-PrinterPort falls back to no-SNMPEnabled path:
                           Add a Set-PrinterPort follow-up call to apply SNMPEnabled after fallback.
                           Remove TCP reachability check (added unnecessary noise and delay for
                           printers that are legitimately offline during pre-staging deployments).
        1.6.6  2026-03-27  Fix Add-PrinterPort -SNMPEnabled ParameterBindingException in 32-bit PS:
                           Same 32-bit module limitation as Add-PrinterDriver. Catch the binding
                           error and retry without -SNMPEnabled (TCP/IP ports default to SNMP
                           enabled, so the default $true behaviour is preserved).
        1.6.5  2026-03-27  Fix Add-PrinterDriver blocked by hardened RPC/policy (HRESULT 0x80070032):
                           Drop -InfPath from the Add-PrinterDriver call when the driver is
                           already staged in the driver store (pnputil published path confirmed).
                           Using -Name only invokes the driver-store lookup code path, which is
                           not restricted by the RPC endpoint / security baseline policies that
                           block the INF-path install path. Add policy registry diagnostics on
                           total failure so admins can see which keys are in effect.
        1.6.4  2026-03-27  Fix null-dereference in Invoke-AddPrinterDriver 64-bit delegate:
                           Get-Content -Raw on an empty stderr temp file returns $null; calling
                           .Trim() on $null threw "You cannot call a method on a null-valued
                           expression" before $proc.ExitCode was ever checked, causing all 5
                           retry attempts to fail with a misleading error.
        1.6.3  2026-03-27  Fix HResult hex format in both catch locations: [int64]-band-0xFFFFFFFF does not
                           work in PS 5.1 because 0xFFFFFFFF is parsed as Int32(-1), which promotes to
                           all-bits-set Int64, making the mask a no-op. Replaced with [int32] cast -- HResult
                           is already Int32 in .NET and {0:X8} renders it as correct 8-digit unsigned hex.
        1.6.2  2026-03-27  Fix [uint32] HResult cast overflow in global error handler (Write-ExceptionDetail).
                           Same bug pattern as v1.6.0 fix but in a different catch location that
                           fires on any outer exception (e.g. pnputil Step 6 failure).
        1.6.1  2026-03-27  Fix 64-bit delegate temp script missing $ErrorActionPreference = 'Stop':
                           Add-PrinterDriver non-terminating errors were silently swallowed, child
                           process exited 0, caller concluded success but driver was never registered.
                           Added stderr capture to expose the actual error message on failure.
        1.6.0  2026-03-26  Fix Add-PrinterDriver failure in 32-bit PS (IME context): delegate to
                           sysnative 64-bit PowerShell via a temp script file when Is64BitProcess
                           is False. Avoids PS5.1 ArgumentList quoting issues with driver names
                           containing spaces. Fix HResult [uint32] cast overflow in retry-loop
                           exception handler that caused the retry loop to abort after attempt 1.
        1.5.0  2026-03-26  Smarter Spooler initialization wait: reads spoolsv process start time
                           after restart and only sleeps the remaining time needed to reach a
                           5-second uptime minimum. Avoids over-waiting when Restart-Spooler
                           itself already consumed several seconds.
        1.4.0  2026-03-25  Extended diagnostic logging: current-step tracker, disk space check,
                           PrintService/Admin event log queries on failures, rich exception detail
                           (type/inner/HResult/position), TCP port reachability check, print
                           configuration readback, registry key confirmation after each major step.
        1.3.0  2026-03-25  Removed 64-bit re-launch block entirely; print management cmdlets work
                           from 32-bit PowerShell on Windows 10/11. pnputil resolved via sysnative
                           path only. Eliminates Start-Process argument quoting complexity.
        1.2.0  2026-03-25  Expanded diagnostic logging; added per-step timing, driver/port/printer
                           detail dumps, driver folder inventory, and script versioning.
        1.1.0  2026-03-25  Fixed 64-bit re-launch argument quoting (Start-Process does not auto-quote
                           array elements with spaces in PowerShell 5.1); moved logging setup before
                           re-launch block so the 32-bit wrapper process writes to the log.
        1.0.0              Initial release.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string] $DriverName,

    [Parameter(Mandatory = $false)]
    [string] $InfPath = '',

    [Parameter(Mandatory = $false)]
    [string] $DriverSourceFolder = '',

    [Parameter(Mandatory = $false)]
    [ValidateRange(10, 300)]
    [int] $SpoolerTimeoutSeconds = 60,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string] $PrinterName,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string] $PrinterIPAddress,

    [Parameter(Mandatory = $false)]
    [string] $PortName = '',

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 65535)]
    [int] $PortNumber = 9100,

    [Parameter(Mandatory = $false)]
    [ValidateSet('A4','A5','A3','Letter','Legal','Executive','Tabloid','Statement','B4','B5','Folio','Quarto','Note')]
    [string] $PaperSize,

    [Parameter(Mandatory = $false)]
    [ValidateSet('OneSided','TwoSidedLongEdge','TwoSidedShortEdge')]
    [string] $DuplexMode,

    [Parameter(Mandatory = $false)]
    [ValidateSet('Color','Monochrome')]
    [string] $ColorMode,

    [Parameter(Mandatory = $false)]
    [ValidateSet('Enabled','Disabled')]
    [string] $Collate,

    [Parameter(Mandatory = $false)]
    [string] $PrintTicketXml,

    [Parameter(Mandatory = $false)]
    [string] $Location,

    [Parameter(Mandatory = $false)]
    [string] $Comment,

    [Parameter(Mandatory = $false)]
    [switch] $Shared,

    [Parameter(Mandatory = $false)]
    [string] $ShareName,

    [Parameter(Mandatory = $false)]
    [string] $LprQueueName,

    [Parameter(Mandatory = $false)]
    [ValidateNotNull()]
    [bool] $SNMPEnabled = $true,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string] $SNMPCommunity = 'public'
)

$ScriptVersion = '1.7.1'

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region -- Logging ----
$LogDir  = Join-Path $env:ProgramData 'Microsoft\IntuneManagementExtension\Logs'
$LogFile = Join-Path $LogDir 'Install-Printer.log'

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

function Invoke-AddPrinterDriver {
    # Wrapper for Add-PrinterDriver that works correctly from 32-bit PowerShell.
    # IME runs scripts as 32-bit; Add-PrinterDriver CIM calls to the 64-bit Spooler fail from
    # a 32-bit process. When 32-bit-on-64-bit is detected, the command is written to a temp
    # script and executed via sysnative (native 64-bit) PowerShell to avoid the CIM mismatch.
    # A temp file is used to sidestep PS5.1 Start-Process ArgumentList quoting issues with
    # driver names that contain spaces.
    param(
        [Parameter(Mandatory = $true)][string] $Name,
        [Parameter(Mandatory = $false)][string] $InfPath
    )
    if (-not [Environment]::Is64BitProcess -and [Environment]::Is64BitOperatingSystem) {
        Write-Log 'Running as 32-bit process on 64-bit OS -- delegating Add-PrinterDriver to sysnative PowerShell...'
        $ps64    = "$env:windir\sysnative\WindowsPowerShell\v1.0\powershell.exe"
        $tmpFile = [IO.Path]::GetTempFileName() + '.ps1'
        $safeName = $Name -replace "'", "''"
        $cmd = if ($InfPath) {
            $safeInf = $InfPath -replace "'", "''"
            "Add-PrinterDriver -Name '$safeName' -InfPath '$safeInf'"
        } else {
            "Add-PrinterDriver -Name '$safeName'"
        }
        # $ErrorActionPreference = 'Stop' ensures Add-PrinterDriver errors are terminating
        # so the child process exits non-zero on failure -- without this the process exits 0
        # even when Add-PrinterDriver silently fails (non-terminating error), making the
        # caller think the driver was registered when it was not.
        $script = "`$ErrorActionPreference = 'Stop'" + [Environment]::NewLine + $cmd
        Set-Content -Path $tmpFile -Value $script -Encoding UTF8
        $errFile = [IO.Path]::GetTempFileName()
        try {
            $proc = Start-Process -FilePath $ps64 `
                        -ArgumentList "-NonInteractive -ExecutionPolicy Bypass -File `"$tmpFile`"" `
                        -Wait -PassThru -WindowStyle Hidden `
                        -RedirectStandardError $errFile
            # Get-Content -Raw returns $null on an empty file; guard before calling .Trim()
            $rawErr  = if (Test-Path $errFile) { Get-Content $errFile -Raw -ErrorAction SilentlyContinue } else { $null }
            $errText = if ($rawErr) { $rawErr.Trim() } else { '' }
            if ($errText) { Write-Log "Add-PrinterDriver (64-bit delegate) stderr: $errText" 'WARN' }
            Write-Log "Add-PrinterDriver (64-bit delegate) exit code: $($proc.ExitCode)"
            if ($proc.ExitCode -ne 0) {
                throw "Add-PrinterDriver (64-bit delegate) exited with code $($proc.ExitCode). $errText"
            }
        } finally {
            Remove-Item $tmpFile  -Force -ErrorAction SilentlyContinue
            Remove-Item $errFile  -Force -ErrorAction SilentlyContinue
        }
    } else {
        if ($InfPath) {
            Add-PrinterDriver -Name $Name -InfPath $InfPath -ErrorAction Stop
        } else {
            Add-PrinterDriver -Name $Name -ErrorAction Stop
        }
    }
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

function Restart-Spooler {
    param([int] $TimeoutSeconds = 30)
    Write-Log 'Restarting Print Spooler to process staged driver...'
    try {
        Stop-Service  -Name Spooler -Force -ErrorAction SilentlyContinue
        Start-Sleep   -Seconds 2
        Start-Service -Name Spooler -ErrorAction Stop
        $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
        $svc = Get-Service -Name Spooler
        while ((Get-Date) -lt $deadline) {
            $svc.Refresh()
            if ($svc.Status -eq 'Running') {
                Write-Log 'Print Spooler restarted successfully.'
                return
            }
            Start-Sleep -Seconds 2
        }
        throw "Spooler did not reach Running state within $TimeoutSeconds seconds after restart."
    } catch {
        Write-Log "Spooler restart warning: $($_.Exception.Message)" 'WARN'
    }
}

function Test-DriverInSpooler {
    param([string] $Name)
    return ($null -ne (Get-PrinterDriver -Name $Name -ErrorAction SilentlyContinue))
}

function Test-PrinterPortExists {
    param([string] $Name)
    return ($null -ne (Get-PrinterPort -Name $Name -ErrorAction SilentlyContinue))
}

function Test-PrinterExists {
    param([string] $Name)
    return ($null -ne (Get-Printer -Name $Name -ErrorAction SilentlyContinue))
}

# Queries the Microsoft-Windows-PrintService/Admin and /Operational event logs for recent
# errors and warnings. Called after Add-PrinterDriver failures and in the fatal catch block.
function Write-PrintServiceEvents {
    param([int] $LastSeconds = 60)
    $since = (Get-Date).AddSeconds(-$LastSeconds)
    $logs  = @('Microsoft-Windows-PrintService/Admin', 'Microsoft-Windows-PrintService/Operational')
    foreach ($logName in $logs) {
        try {
            $events = Get-WinEvent -FilterHashtable @{
                LogName   = $logName
                Level     = @(1, 2, 3)   # Critical, Error, Warning
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
            # Log not present or access denied - not a failure
            Write-Log "  ($logName not available: $($_.Exception.Message))"
        }
    }
}

# Logs structured exception detail: type, message, inner exceptions, HResult, invocation position.
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
        # [int32] keeps HResult as its native 32-bit type; {0:X8} then renders the
        # unsigned two's-complement hex correctly.  Do NOT use [uint32] (throws
        # OverflowException on negative values) or [int64]-band-0xFFFFFFFF (0xFFFFFFFF
        # is parsed as Int32(-1) in PS 5.1, promoting to all-bits-set Int64 -- no mask).
        Write-Log "HResult           : $('0x{0:X8}' -f [int32]$ex.HResult)  (decimal $($ex.HResult))" 'ERROR'
    }
    Write-Log "Script stack      : $($ErrorRecord.ScriptStackTrace)" 'ERROR'
    if ($ErrorRecord.InvocationInfo -and $ErrorRecord.InvocationInfo.PositionMessage) {
        Write-Log "Position          : $($ErrorRecord.InvocationInfo.PositionMessage.Trim())" 'ERROR'
    }
}
#endregion

#region -- Paper-size keyword map ----
$PaperSizeKeywords = @{
    'A4'        = 'psk:ISOA4'
    'A5'        = 'psk:ISOA5'
    'A3'        = 'psk:ISOA3'
    'Letter'    = 'psk:NorthAmericaLetter'
    'Legal'     = 'psk:NorthAmericaLegal'
    'Executive' = 'psk:NorthAmericaExecutive'
    'Tabloid'   = 'psk:NorthAmericaTabloid'
    'Statement' = 'psk:NorthAmericaStatement'
    'B4'        = 'psk:JISB4'
    'B5'        = 'psk:JISB5'
    'Folio'     = 'psk:OtherMetricFolio'
    'Quarto'    = 'psk:NorthAmericaQuarto'
    'Note'      = 'psk:NorthAmericaNote'
}
#endregion

#region -- Main ----
$CreatedPrinter = $false
$CreatedPort    = $false
$stopwatch      = $null
$CurrentStep    = 'startup'   # updated before each step so the catch block can name the failure

try {
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    Write-Log '==== Install-Printer START ===='
    Write-Log "DriverName          : $DriverName"
    Write-Log "InfPath (raw)       : $InfPath"
    Write-Log "DriverSourceFolder  : $DriverSourceFolder"
    Write-Log "PrinterName         : $PrinterName"
    Write-Log "PrinterIPAddress    : $PrinterIPAddress"
    Write-Log "PortName (raw)      : $PortName"
    Write-Log "PortNumber          : $PortNumber"
    Write-Log "PaperSize           : $PaperSize"
    Write-Log "DuplexMode          : $DuplexMode"
    Write-Log "ColorMode           : $ColorMode"
    Write-Log "Collate             : $Collate"
    Write-Log "PrintTicketXml      : $(if ($PrintTicketXml) { '<provided>' } else { '<not provided>' })"
    Write-Log "Location            : $Location"
    Write-Log "Comment             : $Comment"
    Write-Log "Shared              : $Shared"
    Write-Log "ShareName           : $ShareName"
    Write-Log "SNMPEnabled         : $SNMPEnabled"
    Write-Log "SNMPCommunity       : $SNMPCommunity"
    Write-Log "Script version      : $ScriptVersion"
    Write-Log "Running as          : $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"
    Write-Log "Computer name       : $env:COMPUTERNAME"
    Write-Log "OS arch             : $env:PROCESSOR_ARCHITECTURE"
    if ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64') {
        Write-Log 'ARM64 device: driver must be WHQL-signed. EV Code Signing alone is insufficient on ARM64 and causes error 0xE0000242.' 'WARN'
    }
    Write-Log "OS version          : $([System.Environment]::OSVersion.VersionString)"
    Write-Log "PS version          : $($PSVersionTable.PSVersion)"
    Write-Log "PS 64-bit           : $([Environment]::Is64BitProcess)"
    Write-Log "Process ID          : $([System.Diagnostics.Process]::GetCurrentProcess().Id)"
    Write-Log "Script root         : $PSScriptRoot"

    # Disk space on the system drive (low disk causes silent pnputil failures)
    try {
        $sysDrive = Split-Path $env:windir -Qualifier
        $disk = Get-PSDrive -Name ($sysDrive.TrimEnd(':')) -ErrorAction Stop
        $freeGB = [math]::Round($disk.Free / 1GB, 2)
        $usedGB = [math]::Round($disk.Used / 1GB, 2)
        $level  = if ($freeGB -lt 1) { 'WARN' } else { 'INFO' }
        Write-Log "Disk space ($sysDrive) : ${freeGB} GB free / ${usedGB} GB used" $level
    } catch {
        Write-Log "Disk space check skipped: $($_.Exception.Message)"
    }

    # -- Windows Protected Print Mode check (Windows 11 24H2+) ----
    # Check early so the blocker is surfaced before any driver work begins.
    $wppKey = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers\WPP'
    $wppProp = Get-ItemProperty -Path $wppKey -Name 'EnableWindowsProtectedPrint' -ErrorAction SilentlyContinue
    if ($wppProp -and $wppProp.EnableWindowsProtectedPrint -eq 1) {
        Write-Log 'POLICY BLOCKER DETECTED: Windows Protected Print Mode is ENABLED (EnableWindowsProtectedPrint = 1).' 'WARN'
        Write-Log 'All third-party printer drivers are blocked on Windows 11 24H2+. Add-PrinterDriver will fail.' 'WARN'
        Write-Log 'Fix: set "Configure Windows protected print" to Disabled in your Intune printing policy.' 'WARN'
    } else {
        Write-Log 'Windows Protected Print Mode: not enforced (EnableWindowsProtectedPrint absent or 0).'
    }

    # -- Step 1: Resolve INF path ----
    $CurrentStep = 'Step 1: Resolve INF path'
    if ([string]::IsNullOrWhiteSpace($InfPath)) {
        $discovered = Get-ChildItem -Path $PSScriptRoot -Filter '*.inf' -Recurse -ErrorAction SilentlyContinue |
                    Select-Object -First 1
        if ($discovered) {
            $InfPath = $discovered.FullName
            Write-Log "InfPath not supplied - auto-discovered: $InfPath" 'WARN'
        } else {
            throw 'InfPath was not supplied and no .inf file was found under the script directory.'
        }
    } elseif (-not [System.IO.Path]::IsPathRooted($InfPath)) {
        $InfPath = Join-Path $PSScriptRoot $InfPath
    }

    $InfPath = [System.IO.Path]::GetFullPath($InfPath)
    Write-Log "Resolved InfPath    : $InfPath"

    if (-not (Test-Path -LiteralPath $InfPath -PathType Leaf)) {
        throw "INF file not found: $InfPath"
    }

    # -- Step 2: Resolve driver source folder ----
    $CurrentStep = 'Step 2: Resolve driver source folder'
    if ([string]::IsNullOrWhiteSpace($DriverSourceFolder)) {
        $DriverSourceFolder = Split-Path -Parent $InfPath
    } elseif (-not [System.IO.Path]::IsPathRooted($DriverSourceFolder)) {
        $DriverSourceFolder = Join-Path $PSScriptRoot $DriverSourceFolder
    }
    $DriverSourceFolder = [System.IO.Path]::GetFullPath($DriverSourceFolder)
    Write-Log "Resolved SourceFolder: $DriverSourceFolder"

    if (-not (Test-Path -LiteralPath $DriverSourceFolder -PathType Container)) {
        throw "Driver source folder not found: $DriverSourceFolder"
    }

    $driverFiles = Get-ChildItem -Path $DriverSourceFolder -Recurse -File -ErrorAction SilentlyContinue
    Write-Log "Driver folder contents ($($driverFiles.Count) file(s)):"
    foreach ($f in $driverFiles) {
        Write-Log "  $($f.Name)  ($([math]::Round($f.Length / 1KB, 1)) KB)  [$($f.Extension.ToUpper().TrimStart('.'))]"
    }

    # -- Step 3: Locate pnputil ----
    $CurrentStep = 'Step 3: Locate pnputil'
    $pnputil = Resolve-PnpUtil
    Write-Log "Using pnputil       : $pnputil"

    # -- Step 4: Ensure Spooler is running ----
    $CurrentStep = 'Step 4: Ensure Spooler is running'
    Wait-ForSpooler -TimeoutSeconds $SpoolerTimeoutSeconds

    # -- Step 5: Validate IP address, derive PortName ----
    $CurrentStep = 'Step 5: Validate IP address'
    $ipAddr = $null
    if (-not [System.Net.IPAddress]::TryParse($PrinterIPAddress, [ref]$ipAddr)) {
        throw "Invalid IP address: '$PrinterIPAddress'"
    }
    Write-Log "Validated IP address: $PrinterIPAddress"

    if ([string]::IsNullOrWhiteSpace($PortName)) {
        $PortName = "IP_$PrinterIPAddress"
    }
    Write-Log "Resolved PortName   : $PortName"

    # -- Step 6: Stage driver into the Driver Store ----
    $CurrentStep = 'Step 6: Stage driver (pnputil)'
    Write-Log "Staging driver with pnputil: $InfPath"
    $pnpArgs   = @('/add-driver', $InfPath, '/install')
    $pnpOutput = & $pnputil @pnpArgs 2>&1
    $pnpExit   = $LASTEXITCODE

    if ($pnpOutput) { Write-Log ($pnpOutput -join [Environment]::NewLine) }
    Write-Log "pnputil exit code   : $pnpExit"

    # 0 = success, 259 = driver already staged / up-to-date
    if ($pnpExit -notin 0, 259) {
        throw "pnputil failed with exit code $pnpExit staging '$InfPath'. See log for output."
    }

    # Capture the published INF name (e.g. oem31.inf) for Add-PrinterDriver -InfPath
    $publishedInfPath = $null
    $joinedOutput = $pnpOutput -join "`n"
    if ($joinedOutput -match 'Published Name:\s+(\S+)') {
        $candidate = Join-Path $env:windir "INF\$($Matches[1])"
        if (Test-Path $candidate) {
            $publishedInfPath = $candidate
            Write-Log "Published INF path  : $publishedInfPath"
        }
    }

    Write-Log "Step 6 elapsed: $($stopwatch.Elapsed.TotalSeconds.ToString('F1'))s"

    # -- Step 7: Restart Spooler + delay (CRITICAL for CoInstaller drivers e.g. Brother) ----
    $CurrentStep = 'Step 7: Restart Spooler'
    Restart-Spooler

    # Give the Spooler at least 5 seconds of uptime before calling Add-PrinterDriver.
    # Restart-Spooler may itself have taken several seconds; only sleep the remainder.
    # This avoids over-waiting while still covering Spooler initialization time.
    $nativeArch = if ($env:PROCESSOR_ARCHITEW6432) { $env:PROCESSOR_ARCHITEW6432 } else { $env:PROCESSOR_ARCHITECTURE }
    $driverArch = switch ($nativeArch) {
        'AMD64' { 'Windows x64' }
        'ARM64' { 'Windows ARM64' }
        'x86'   { 'Windows NT x86' }
        default { 'Windows x64' }
    }
    $drvRegKeys = @(
        "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Environments\$driverArch\Drivers\Version-3\$DriverName",
        "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Environments\$driverArch\Drivers\Version-4\$DriverName"
    )
    $drvRegKeyFound = $drvRegKeys | Where-Object { Test-Path $_ } | Select-Object -First 1
    
    $minUptimeSeconds = 5
    $spoolerProc = Get-Process -Name spoolsv -ErrorAction SilentlyContinue
    if ($spoolerProc) {
        $uptimeSec = (New-TimeSpan -Start $spoolerProc.StartTime -End (Get-Date)).TotalSeconds
        $remaining = $minUptimeSeconds - $uptimeSec
        if ($remaining -gt 0) {
            Write-Log "Spooler uptime $($uptimeSec.ToString('F1'))s -- waiting $($remaining.ToString('F1'))s more for initialization..."
            Start-Sleep -Seconds ([math]::Ceiling($remaining))
        } else {
            Write-Log "Spooler uptime $($uptimeSec.ToString('F1'))s -- initialization window already elapsed, proceeding."
        }
    } else {
        Write-Log "spoolsv process not found -- falling back to 5s delay." 'WARN'
        Start-Sleep -Seconds 5
    }
    if ($spoolerProc) {
        Write-Log "Spooler process: PID=$($spoolerProc.Id)  Started=$($spoolerProc.StartTime.ToString('yyyy-MM-dd HH:mm:ss'))  Memory=$([math]::Round($spoolerProc.WorkingSet64 / 1MB, 1))MB"
    }
    Write-Log "Step 7 elapsed: $($stopwatch.Elapsed.TotalSeconds.ToString('F1'))s"

    # -- Step 8: Register driver with Print Spooler (retry loop) ----
    $CurrentStep = 'Step 8: Register driver (Add-PrinterDriver)'
    if (-not (Test-DriverInSpooler -Name $DriverName)) {
        Write-Log "Registering driver '$DriverName' with Print Spooler..."
        $driverRegistered = $false
        $maxAttempts      = 5

        for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
            Write-Log "Add-PrinterDriver attempt $attempt of $maxAttempts..."
            try {
                # Omit -InfPath when the driver is confirmed staged in the driver store.
                # The -Name-only call uses the driver-store lookup code path, which avoids
                # the INF-path install restriction that hardened RPC/security baseline
                # policies (e.g. Event 817 / 0x80070032) enforce on the file-based path.
                # Fall back to source InfPath only if pnputil did not return a published name.
                if ($publishedInfPath) {
                    Invoke-AddPrinterDriver -Name $DriverName
                } else {
                    Invoke-AddPrinterDriver -Name $DriverName -InfPath $InfPath
                }
                Write-Log "Add-PrinterDriver succeeded on attempt $attempt."
                $driverRegistered = $true
                $driverDetail = Get-PrinterDriver -Name $DriverName -ErrorAction SilentlyContinue
                if ($driverDetail) {
                    Write-Log "Driver registered: Manufacturer='$($driverDetail.Manufacturer)'  Version='$($driverDetail.DriverVersion)'  Environment='$($driverDetail.PrinterEnvironment)'"
                    Write-Log "Driver INF (Spooler): $($driverDetail.InfPath)"
                }
                # Confirm registry key written by the Spooler (variables defined in Step 7)
                Write-Log "Driver registry key : $(if ($drvRegKeyFound -or (Test-Path $drvRegKeys[0]) -or (Test-Path $drvRegKeys[1])) { 'EXISTS' } else { 'NOT FOUND (unexpected)' })"
                Write-Log "Step 8 elapsed: $($stopwatch.Elapsed.TotalSeconds.ToString('F1'))s"
                break
            } catch {
                Write-Log "Attempt $attempt failed: $($_.Exception.Message)" 'WARN'
                Write-Log "  Exception type : $($_.Exception.GetType().FullName)" 'WARN'
                if ($_.Exception.HResult -ne 0) {
                    Write-Log "  HResult        : $('0x{0:X8}' -f [int32]$_.Exception.HResult)  (decimal $($_.Exception.HResult))" 'WARN'
                }
                if ($_.Exception.InnerException) {
                    Write-Log "  Inner exception: $($_.Exception.InnerException.Message)" 'WARN'
                }
                Write-Log "  PrintService event log (last 30s):" 'WARN'
                Write-PrintServiceEvents -LastSeconds 30
                if ($attempt -lt $maxAttempts) {
                    Write-Log "Waiting 10 seconds before retry..."
                    Start-Sleep -Seconds 10
                }
            }
        }

        if (-not $driverRegistered) {
            # Log print-driver policy registry values to identify what is blocking Add-PrinterDriver.
            Write-Log 'Checking print-driver policy registry keys for diagnostics...' 'WARN'
            foreach ($regKey in @(
                'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers',
                'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers\RPC',
                'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers\PointAndPrint'
            )) {
                if (Test-Path $regKey) {
                    $regProps = Get-ItemProperty $regKey -ErrorAction SilentlyContinue
                    if ($regProps) {
                        $regProps.PSObject.Properties |
                            Where-Object { $_.Name -notlike 'PS*' } |
                            ForEach-Object {
                                Write-Log "  Policy $($regKey.Split('\')[-1])\$($_.Name) = $($_.Value)" 'WARN'
                            }
                    } else {
                        Write-Log "  Policy key exists but has no values: $regKey"
                    }
                } else {
                    Write-Log "  Policy key not present: $regKey"
                }
            }
            throw "Driver '$DriverName' could not be registered with the Print Spooler after $maxAttempts attempts."
        }
    } else {
        Write-Log "Driver '$DriverName' is already registered with the Print Spooler."
    }

    # -- Step 8b: Verify driver is registered before proceeding ----
    if (-not (Test-DriverInSpooler -Name $DriverName)) {
        throw "Driver '$DriverName' not visible in Print Spooler after registration. Cannot create printer queue."
    }

    # -- Step 9: Create TCP/IP port ----
    $CurrentStep = 'Step 9: Create TCP/IP port'
    if (-not (Test-PrinterPortExists -Name $PortName)) {
        Write-Log "Creating TCP/IP port '$PortName' (${PrinterIPAddress}:${PortNumber})..."
        # Add-PrinterPort -SNMPEnabled is not supported on any PS bitness on Windows 11 24H2+.
        # Omit it entirely; SNMPEnabled is applied unconditionally via the registry after creation.
        $portArgs = @{
            Name                 = $PortName
            PrinterHostAddress   = $PrinterIPAddress
            PortNumber           = $PortNumber
            ErrorAction          = 'Stop'
        }
        if ($LprQueueName) { $portArgs['LprHostAddress'] = $PrinterIPAddress; $portArgs['LprQueueName'] = $LprQueueName; $portArgs['Protocol'] = 'LPR' }
        Add-PrinterPort @portArgs
        $CreatedPort = $true
        Write-Log "Port '$PortName' created successfully."

        # Add-PrinterPort -SNMPEnabled is unreliable across Windows versions (removed in 24H2+).
        # Write SNMPEnabled directly to the Standard TCP/IP Port monitor registry key.
        $portRegKey = "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Monitors\Standard TCP/IP Port\Ports\$PortName"
        if (Test-Path $portRegKey) {
            try {
                Set-ItemProperty -Path $portRegKey -Name 'SNMPEnabled' -Value ([int]$SNMPEnabled) -Type DWord -ErrorAction Stop
                Write-Log "SNMP applied via registry: SNMPEnabled=$SNMPEnabled"
            } catch {
                Write-Log "Registry SNMP configuration warning: $($_.Exception.Message)" 'WARN'
            }
        } else {
            Write-Log "Port registry key not found - SNMP configuration skipped." 'WARN'
        }

        # Write SNMPCommunity to registry: Add-PrinterPort does not expose -SNMPCommunity in
        # any PS bitness. Windows defaults to 'public'; write explicitly so the configured value
        # is applied regardless of the caller's intent.
        if ($SNMPEnabled) {
            $portRegKeyComm = "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Monitors\Standard TCP/IP Port\Ports\$PortName"
            if (Test-Path $portRegKeyComm) {
                try {
                    Set-ItemProperty -Path $portRegKeyComm -Name 'SNMPCommunity' -Value $SNMPCommunity -Type String -ErrorAction Stop
                    Write-Log "SNMP community string set: '$SNMPCommunity'"
                } catch {
                    Write-Log "SNMPCommunity registry write warning: $($_.Exception.Message)" 'WARN'
                }
            } else {
                Write-Log 'Port registry key not found -- SNMPCommunity not set.' 'WARN'
            }
        }

        $portDetail = Get-PrinterPort -Name $PortName -ErrorAction SilentlyContinue
        if ($portDetail) {
            # On 32-bit PS, Get-PrinterPort does not return SNMPEnabled correctly.
            # Read it directly from the registry when running 32-bit so the log is accurate.
            if (-not [Environment]::Is64BitProcess) {
                $portRegKey = "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Monitors\Standard TCP/IP Port\Ports\$PortName"
                $snmpRaw = Get-ItemProperty -Path $portRegKey -Name 'SNMPEnabled' -ErrorAction SilentlyContinue
                $snmpDisplay = if ($snmpRaw) { [bool]$snmpRaw.SNMPEnabled } else { $portDetail.SNMPEnabled }
            } else {
                $snmpDisplay = $portDetail.SNMPEnabled
            }
            Write-Log "Port details: Address='$($portDetail.PrinterHostAddress)'  Port=$($portDetail.PortNumber)  Protocol=$($portDetail.Protocol)  SNMP=$snmpDisplay"
        }
    } else {
        Write-Log "Port '$PortName' already exists - skipping creation."
    }
    Write-Log "Step 9 elapsed: $($stopwatch.Elapsed.TotalSeconds.ToString('F1'))s"

    # -- Step 10: Create printer queue ----
    $CurrentStep = 'Step 10: Create printer queue'
    if (-not (Test-PrinterExists -Name $PrinterName)) {
        Write-Log "Creating printer queue '$PrinterName' (driver='$DriverName', port='$PortName')..."
        Add-Printer -Name $PrinterName -DriverName $DriverName -PortName $PortName -ErrorAction Stop
        $CreatedPrinter = $true
        Write-Log "Printer queue '$PrinterName' created successfully."
        $printerDetail = Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue
        if ($printerDetail) {
            Write-Log "Printer details: Driver='$($printerDetail.DriverName)'  Port='$($printerDetail.PortName)'  Status='$($printerDetail.PrinterStatus)'  Shared=$($printerDetail.Shared)"
        }
        # Confirm registry key written by the Spooler
        $prtRegKey = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Printers\$PrinterName"
        Write-Log "Printer registry key: $(if (Test-Path $prtRegKey) { 'EXISTS' } else { 'NOT FOUND (unexpected)' })"
    } else {
        Write-Log "Printer queue '$PrinterName' already exists - skipping creation."
    }
    Write-Log "Step 10 elapsed: $($stopwatch.Elapsed.TotalSeconds.ToString('F1'))s"

    # -- Step 11: Configure printer (warning on failure, no rollback) ----
    $CurrentStep = 'Step 11: Configure printer'
    try {
        # 11a: Printer properties via Set-Printer
        $printerProps = @{}
        if ($PSBoundParameters.ContainsKey('Location')) { $printerProps['Location'] = $Location }
        if ($PSBoundParameters.ContainsKey('Comment'))  { $printerProps['Comment']  = $Comment }
        if ($Shared) {
            $printerProps['Shared'] = $true
            if ($PSBoundParameters.ContainsKey('ShareName')) { $printerProps['ShareName'] = $ShareName }
        }

        if ($printerProps.Count -gt 0) {
            Write-Log 'Applying printer properties...'
            foreach ($key in $printerProps.Keys) { Write-Log "  $key = $($printerProps[$key])" }
            Set-Printer -Name $PrinterName @printerProps -ErrorAction Stop
            Write-Log 'Printer properties applied.'
        }

        # 11b: Print configuration via Set-PrintConfiguration
        $configParams = @{}
        if ($PSBoundParameters.ContainsKey('DuplexMode')) {
            $configParams['DuplexingMode'] = $DuplexMode
        }
        if ($PSBoundParameters.ContainsKey('ColorMode')) {
            $configParams['Color'] = ($ColorMode -eq 'Color')
        }
        if ($PSBoundParameters.ContainsKey('Collate')) {
            $configParams['Collated'] = ($Collate -eq 'Enabled')
        }

        if ($configParams.Count -gt 0) {
            Write-Log 'Applying print configuration...'
            foreach ($key in $configParams.Keys) { Write-Log "  $key = $($configParams[$key])" }
            Set-PrintConfiguration -PrinterName $PrinterName @configParams -ErrorAction Stop
            Write-Log 'Print configuration applied.'
            # Readback: confirm what was actually written
            $readback = Get-PrintConfiguration -PrinterName $PrinterName -ErrorAction SilentlyContinue
            if ($readback) {
                Write-Log "Config readback: DuplexingMode='$($readback.DuplexingMode)'  Color=$($readback.Color)  Collated=$($readback.Collated)"
            }
        }

        # 11c: Paper size via PrintTicket XML
        if ($PSBoundParameters.ContainsKey('PaperSize')) {
            $keyword = $PaperSizeKeywords[$PaperSize]
            Write-Log "Setting paper size to '$PaperSize' ($keyword)..."

            $currentConfig = Get-PrintConfiguration -PrinterName $PrinterName -ErrorAction Stop
            [xml]$ticket = $currentConfig.PrintTicketXml

            $nsMgr = [System.Xml.XmlNamespaceManager]::new($ticket.NameTable)
            $nsMgr.AddNamespace('psf', 'http://schemas.microsoft.com/windows/2003/08/printing/printschemaframework')
            $nsMgr.AddNamespace('psk', 'http://schemas.microsoft.com/windows/2003/08/printing/printschemakeywords')

            $feature = $ticket.SelectSingleNode("//psf:Feature[@name='psk:PageMediaSize']", $nsMgr)
            if ($feature) {
                $option = $feature.SelectSingleNode('psf:Option', $nsMgr)
                if ($option) {
                    $option.SetAttribute('name', $keyword)
                    $scored = $option.SelectNodes('psf:ScoredProperty', $nsMgr)
                    foreach ($sp in $scored) { $option.RemoveChild($sp) | Out-Null }
                    Set-PrintConfiguration -PrinterName $PrinterName -PrintTicketXml $ticket.OuterXml -ErrorAction Stop
                    Write-Log "Paper size set to '$PaperSize'."
                } else {
                    Write-Log 'PageMediaSize Option element not found in PrintTicket.' 'WARN'
                }
            } else {
                Write-Log 'PageMediaSize feature not found in PrintTicket - driver may not support it.' 'WARN'
            }
        }

        # 11d: Custom PrintTicketXml (applied last, can override everything)
        if ($PSBoundParameters.ContainsKey('PrintTicketXml') -and -not [string]::IsNullOrWhiteSpace($PrintTicketXml)) {
            Write-Log 'Applying custom PrintTicketXml...'
            Set-PrintConfiguration -PrinterName $PrinterName -PrintTicketXml $PrintTicketXml -ErrorAction Stop
            Write-Log 'Custom PrintTicketXml applied.'
        }

    } catch {
        Write-Log "Printer configuration warning: $($_.Exception.Message)" 'WARN'
        Write-Log 'Printer is functional but some configuration settings may not have been applied.' 'WARN'
    }

    # -- Step 12: Final verification ----
    $CurrentStep = 'Step 12: Final verification'
    $driverOk  = Test-DriverInSpooler   -Name $DriverName
    $portOk    = Test-PrinterPortExists -Name $PortName
    $printerOk = Test-PrinterExists     -Name $PrinterName

    Write-Log "Verification - Driver in Spooler : $driverOk"
    Write-Log "Verification - Port exists       : $portOk"
    Write-Log "Verification - Printer exists    : $printerOk"

    if ($driverOk -and $portOk -and $printerOk) {
        Write-Log 'Verification PASSED - all components are present.'
    } else {
        if (-not $driverOk)    { Write-Log "Driver '$DriverName' not visible via Get-PrinterDriver." 'WARN' }
        if (-not $portOk)      { Write-Log "Port '$PortName' not found after installation." 'WARN' }
        if (-not $printerOk)   { Write-Log "Printer '$PrinterName' not found after installation." 'WARN' }
    }

    $stopwatch.Stop()
    Write-Log "Elapsed time        : $($stopwatch.Elapsed.TotalSeconds.ToString('F1')) seconds"
    Write-Log '==== Install-Printer END (SUCCESS) ===='
    exit 0

} catch {
    # -- Rollback ----
    if ($CreatedPrinter) {
        Write-Log "Rolling back: removing printer queue '$PrinterName'..." 'WARN'
        Remove-Printer -Name $PrinterName -ErrorAction SilentlyContinue
    }
    if ($CreatedPort) {
        Write-Log "Rolling back: removing port '$PortName'..." 'WARN'
        Remove-PrinterPort -Name $PortName -Confirm:$false -ErrorAction SilentlyContinue
    }

    Write-Log "Failed at         : $CurrentStep" 'ERROR'
    Write-ExceptionDetail -ErrorRecord $_
    Write-Log '--- PrintService event log (last 60s) ---' 'ERROR'
    Write-PrintServiceEvents -LastSeconds 60
    if ($stopwatch) { Write-Log "Elapsed time: $($stopwatch.Elapsed.TotalSeconds.ToString('F1')) seconds" }
    Write-Log '==== Install-Printer END (FAILURE) ===='
    exit 1
}
#endregion
