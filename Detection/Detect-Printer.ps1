<#
.SYNOPSIS
    Detection script for an Intune Win32 app that installs a network printer.

.DESCRIPTION
    Upload this script in the Intune Win32 app wizard under:
      App information > Detection rules > Rules format: "Use a custom detection script"

    Intune runs the script with NO arguments and interprets the result as:
      Exit 0  + any STDOUT text  ->  App is INSTALLED  (detected)
      Exit 0  + no STDOUT text   ->  App is NOT installed
      Non-zero exit               ->  Detection error (Intune flags this)

    Detection logic (both checks must pass, registry-based - no Print Spooler dependency):
      1. Printer queue in registry  (HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Printers\<name>)
      2. Driver in registry         (HKLM:\SYSTEM\CurrentControlSet\Control\Print\Environments\...\<name>)

    Driver Store presence is logged for diagnostics but is not required for detection.

    CONFIGURATION: Set $PrinterName and $DriverName below before packaging.

.NOTES
    Log: %ProgramData%\Microsoft\IntuneManagementExtension\Logs\Detect-Printer.log
    Tested on Windows 10 21H2+ and Windows 11 (x64 and ARM64).
    ARM64: Set "Run script as 32-bit process on 64-bit clients" to No in Intune to ensure
           native 64-bit/ARM64 PowerShell is used.

    Version history:
        1.2.1  2026-03-28  Changed log file handling from Clear-Content (overwrite on each run) to
                           append with '=== NEW RUN ===' separator, consistent with CLAUDE.md guideline
                           of preserving rolling log history across detection cycles.
        1.2.0  2026-03-25  Switched to registry-based detection for printer queue and driver;
                           eliminates Print Spooler dependency. Driver Store check retained as
                           supplementary log info only.
        1.1.0  2026-03-25  Added script versioning and extended system info to log output.
        1.0.0              Initial release.
#>

# ==============================================================================
#  CONFIGURE BEFORE PACKAGING
#  Set $PrinterName to the exact printer queue name used in the install command.
#  Set $DriverName to the exact driver name from the INF file.
#  Both values must match the parameters used in Install-Printer.ps1.
#  Examples:
#    $PrinterName = 'Office Printer 3rd Floor'
#    $DriverName  = 'HP Universal Printing PCL 6'
# ==============================================================================
$PrinterName   = 'Brother MFC-J6540DW'
$DriverName    = 'Brother MFC-J6540DW Printer'
$ScriptVersion = '1.2.1'
# ==============================================================================

$ErrorActionPreference = 'SilentlyContinue'

#region -- Logging ----------------------------------------------------------------
# Logging must be defined BEFORE placeholder validation so we can log the error.
$LogDir  = Join-Path $env:ProgramData 'Microsoft\IntuneManagementExtension\Logs'
$LogFile = Join-Path $LogDir 'Detect-Printer.log'

if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
}
# Append separator instead of clearing - preserves rolling detection history
if (Test-Path $LogFile) {
    Add-Content -Path $LogFile -Value ("`n" + ('=' * 80) + "`n[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] NEW RUN`n" + ('=' * 80)) -Encoding UTF8 -ErrorAction SilentlyContinue
}

function Write-Log {
    param(
        [Parameter(Mandatory = $true)][string] $Message,
        [ValidateSet('INFO','WARN','ERROR')][string] $Level = 'INFO'
    )
    $line = '[{0}] [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    # Do NOT use Write-Host or Write-Error -- all stdout/stderr is reserved for Intune detection.
    Add-Content -Path $LogFile -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue
}
#endregion

#region -- Placeholder validation -------------------------------------------------
# Guard: detect if the administrator forgot to set $PrinterName / $DriverName.
if ($PrinterName -match '^<.*>$' -or $DriverName -match '^<.*>$') {
    Write-Log 'ERROR: $PrinterName and/or $DriverName still contain placeholder values.' 'ERROR'
    Write-Log "  PrinterName = '$PrinterName'"
    Write-Log "  DriverName  = '$DriverName'"
    Write-Log 'Open Detect-Printer.ps1 and set both variables before uploading to Intune.' 'ERROR'
    Write-Log '========== Detect-Printer END (CONFIGURATION ERROR) =========='
    # Exit 0 with no STDOUT = Intune treats as NOT installed (safest default).
    exit 0
}
#endregion

#region -- Helpers ----------------------------------------------------------------
function Resolve-PnpUtil {
    $candidates = @(
        (Join-Path $env:windir 'sysnative\pnputil.exe'),
        (Join-Path $env:windir 'System32\pnputil.exe')
    )
    return ($candidates | Where-Object { Test-Path $_ } | Select-Object -First 1)
}

# Registry-based printer detection: does not require the Print Spooler to be running.
# Key is present whenever the printer queue is installed, regardless of Spooler state.
function Test-PrinterInRegistry {
    param([string] $Name)
    $key = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Printers\$Name"
    return (Test-Path $key)
}

# Registry-based driver detection: checks all printer environments (x64, ARM64, x86).
# Does not require the Print Spooler to be running.
function Test-DriverInRegistry {
    param([string] $Name)
    $environments = @('Windows x64', 'Windows ARM64', 'Windows NT x86')
    $driverVersions = @('Version-3', 'Version-4') # Add Version-4 for class drivers
    foreach ($env in $environments) {
        foreach ($version in $driverVersions) {
            $key = "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Environments\$env\Drivers\$version\$Name"
            if (Test-Path $key) { return $true }
        }
    }
    return $false
}

function Test-DriverInDriverStore {
    param([string] $Name, [string] $PnpUtilPath)
    if (-not $PnpUtilPath) { return $false }

    # Method 1: block-based pnputil /enum-drivers parsing
    $output = & $PnpUtilPath /enum-drivers 2>&1
    if ($LASTEXITCODE -eq 0 -and $output) {
        $blocks = ($output -join "`n") -split '(?m)^\s*$' | Where-Object { $_.Trim() }
        foreach ($block in $blocks) {
            if ($block -match [regex]::Escape($Name)) { return $true }
        }
    }

    # Method 2: scan staged INF files in the Driver Store FileRepository
    $repoPath = Join-Path $env:windir 'System32\DriverStore\FileRepository'
    if (Test-Path $repoPath) {
        $match = Get-ChildItem -Path $repoPath -Recurse -Filter '*.inf' -ErrorAction SilentlyContinue |
                 Select-String -Pattern $Name -SimpleMatch -List -ErrorAction SilentlyContinue |
                 Select-Object -First 1
        if ($match) { return $true }
    }

    return $false
}
#endregion

#region -- Detection --------------------------------------------------------------
Write-Log '========== Detect-Printer START =========='
Write-Log "Script version : $ScriptVersion"
Write-Log "PrinterName    : $PrinterName"
Write-Log "DriverName     : $DriverName"
Write-Log "Running as     : $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"
Write-Log "Computer name  : $env:COMPUTERNAME"
Write-Log "OS arch        : $env:PROCESSOR_ARCHITECTURE"
Write-Log "OS version     : $([System.Environment]::OSVersion.VersionString)"
Write-Log "PS version     : $($PSVersionTable.PSVersion)"
Write-Log "PS 64-bit      : $([Environment]::Is64BitProcess)"

$pnputil = Resolve-PnpUtil
Write-Log "pnputil     : $pnputil"

# Primary checks use registry -- no Print Spooler dependency.
$printerInRegistry = Test-PrinterInRegistry -Name $PrinterName
$driverInRegistry  = Test-DriverInRegistry  -Name $DriverName
# Driver Store check is supplementary diagnostic info only.
$inStore           = Test-DriverInDriverStore -Name $DriverName -PnpUtilPath $pnputil

Write-Log "Printer in registry  : $printerInRegistry"
Write-Log "Driver in registry   : $driverInRegistry"
Write-Log "Driver in Store      : $inStore"

if ($printerInRegistry -and $driverInRegistry) {
    # -- DETECTED ----------------------------------------------------------------
    # Both printer queue and driver are present in the registry. Any stdout + exit 0 = installed.
    $msg = "DETECTED: Printer '$PrinterName' with driver '$DriverName' is installed (PrinterRegistry=$printerInRegistry, DriverRegistry=$driverInRegistry, Store=$inStore)."
    Write-Log $msg
    Write-Log '========== Detect-Printer END (DETECTED) =========='
    Write-Output $msg
    exit 0

} else {
    # -- NOT DETECTED ------------------------------------------------------------
    # Exit 0 with NO stdout = Intune triggers installation.
    $reason = @()
    if (-not $printerInRegistry) { $reason += "printer queue '$PrinterName' not found in registry" }
    if (-not $driverInRegistry)  { $reason += "driver '$DriverName' not found in registry" }
    Write-Log "NOT DETECTED: $($reason -join '; ')."
    if ($inStore) {
        Write-Log "Note: driver '$DriverName' is present in the Driver Store but printer queue or driver registry key is missing." 'WARN'
    }
    Write-Log '========== Detect-Printer END (NOT DETECTED) =========='
    exit 0
}
#endregion
