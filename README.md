# Install Printer via Microsoft Intune

Deploy complete network printers to Windows endpoints using Microsoft Intune Win32 app deployment. The scripts install the printer driver, create a TCP/IP port, create the printer queue, and optionally configure printer settings (paper size, duplex, color, collation) -- all automated and brand/model agnostic.

---

## Table of Contents

- [Prerequisites](#prerequisites)
- [Step 1 -- Download Printer Drivers](#step-1----download-printer-drivers)
  - [Reviewing the INF File](#reviewing-the-inf-file)
  - [Finding the Driver Name](#finding-the-driver-name)
- [Step 2 -- Gather Deployment Information](#step-2----gather-deployment-information)
- [Package Structure](#package-structure)
- [Scripts](#scripts)
  - [Install-Printer.ps1](#install-printerps1)
  - [Detect-Printer.ps1](#detect-printerps1)
  - [Uninstall-Printer.ps1](#uninstall-printerps1)
- [Intune Win32 App Configuration](#intune-win32-app-configuration)
- [Printer Configuration Reference](#printer-configuration-reference)
  - [PaperSize Values](#papersize-values)
  - [DuplexMode Values](#duplexmode-values)
  - [ColorMode Values](#colormode-values)
  - [Collate Values](#collate-values)
  - [PrintTicketXml (Advanced)](#printticketxml-advanced)
- [ARM64 Devices](#arm64-devices)
- [Log Locations](#log-locations)
- [Exit Codes](#exit-codes)
- [Troubleshooting](#troubleshooting)

---

## Prerequisites

- Microsoft Intune subscription
- Windows 10 21H2 or later / Windows 11 (x64 or ARM64)
- [Microsoft Win32 Content Prep Tool](https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool) (`IntuneWinAppUtil.exe`) to package the `.intunewin` file
- Printer driver package (INF, CAT, DLL/SYS files) from the manufacturer
- Network printer IP address and desired printer queue name

---

## Step 1 -- Download Printer Drivers

Download the compatible printer drivers for your printer model from the manufacturer's website. The examples in this guide use the **HP Universal Printing PCL 6** driver. The same steps apply to any printer driver from any manufacturer (HP, Kyocera, Xerox, Brother, Canon, Ricoh, etc.).

After downloading:

1. Extract the contents of the ZIP file and open the folder that contains the driver package.
2. You will find several files including `.DLL`, `.CAB`, a Security Catalog (`.CAT`) file, and a Setup Information (`.INF`) file.

Every driver package includes a **Setup Information (INF) file** that contains details required by the system to install the printer driver:

- Driver files
- Registry entries
- Device IDs
- Catalog file information
- Printer name

### Reviewing the INF File

Open the INF file and review its contents to identify the files it references. At the beginning of the file, you will find a reference such as:

```
CatalogFile=x3UNIVX.cat
```

You **must** include this catalog file in your Intune deployment along with the INF file. The catalog file is used to verify that the driver package has not been altered after publication by validating its digital signature.

Look for the `[SourceDisksNames]` section in the INF file, which lists the additional files required for driver installation. For more information about driver package components, refer to [Components of a Driver Package -- Windows Drivers | Microsoft Learn](https://learn.microsoft.com/en-us/windows-hardware/drivers/install/components-of-a-driver-package).


### Finding the Driver Name

The **driver name** is required in three places:

- `-DriverName` parameter in the Intune install command
- `-DriverName` parameter in the Intune uninstall command
- `$DriverName` variable in `Detect-Printer-Driver.ps1` before upload

The name is declared in the INF's `[Models.*]` or `[DriverName.*]` section. Each line in that section follows the pattern:

```
DriverDisplayName = InstallSection, HardwareID, ...
```

Different manufacturers format the name differently:

| Format | Example |
|---|---|
| Quoted (HP, Xerox, Kyocera) | `"HP Universal Printing PCL 6" = HPCU215U.GPD, UNIDRV` |
| %Variable% (HP Smart Universal, some Xerox/Ricoh) | `%HPODPRINTERS% = INSTALL_SECTION, ...` — actual name is in the `[Strings]` section at the bottom of the INF |
| Unquoted (Brother, Canon) | `Brother MFC-J6540DW Printer = BRSMJ6540DW20A.DSI, ...` |

**Manually:** Open the INF in a text editor, press **Ctrl+F**, search for `[Models` or `[DriverName`, and read the name from the lines that follow. If the name uses a `%variable%` reference (e.g. `%HPODPRINTERS%`), search for `[Strings]` at the bottom of the INF to find the actual value (e.g. `HPODPRINTERS = "HP Smart Universal Printing"`).

**PowerShell — works for quoted, unquoted, and %variable% formats:**

```powershell
$inf = "C:\Path\To\driver.inf"
$lines = Get-Content -LiteralPath $inf

# ── 1. Parse the [Strings] section so we can resolve %variable% references ──
$strings = @{}
$inStrings = $false
foreach ($line in $lines) {
    if ($line -match '^\s*\[Strings\]')  { $inStrings = $true; continue }
    elseif ($line -match '^\s*\[')       { $inStrings = $false }
    if (-not $inStrings) { continue }
    if ($line -match '^\s*(\w+)\s*=\s*"([^"]+)"') {
        $strings[$Matches[1].Trim()] = $Matches[2].Trim()
    }
}

# ── 2. Scan the [Models.*] / [DriverName.*] section for driver names ────────
$inSection = $false
$names = foreach ($line in $lines) {
    if ($line -match '^\s*\[(Models|DriverName)') { $inSection = $true; continue }
    elseif ($line -match '^\s*\[') { $inSection = $false }
    if (-not $inSection) { continue }
    # Format 1 — quoted:    "HP Universal Printing PCL 6" = ...
    if     ($line -match '^\s*"([^"]+)"\s*=')              { $Matches[1].Trim() }
    # Format 2 — variable:  %HPODPRINTERS% = ...
    elseif ($line -match '^\s*%(\w+)%\s*=') {
        $key = $Matches[1]
        if ($strings.ContainsKey($key)) { $strings[$key] } else { "%$key%" }
    }
    # Format 3 — unquoted:  Brother MFC-J6540DW Printer = ...
    elseif ($line -match '^\s*(\S[^=;]{2,80}?)\s*=')      { $Matches[1].Trim() }
}
$names | Where-Object { $_ } | Sort-Object -Unique
```

Replace `C:\Path\To\driver.inf` with the actual path to your INF file. If the INF covers multiple models (common with Brother and HP universal drivers), the command lists every model — pick the one matching your specific printer.

The three formats the script handles:

| Format | Example | Used by |
|---|---|---|
| Quoted | `"HP Universal Printing PCL 6" = HPCU215U.GPD, ...` | HP (PCL 6), Xerox, Kyocera |
| %Variable% | `%HPODPRINTERS% = INSTALL_SECTION, ...` → resolved from `[Strings]` | HP (Smart Universal), some Xerox/Ricoh |
| Unquoted | `Brother MFC-J6540DW Printer = BRSMJ6540DW20A.DSI, ...` | Brother, Canon |

---

## Step 2 -- Gather Deployment Information

Before packaging, collect the following information for your printer deployment:

| Information | Example | Where to find it |
|---|---|---|
| **Driver name** | `HP Universal Printing PCL 6` | INF file (see [Finding the Driver Name](#finding-the-driver-name)) |
| **INF file path** | `.\Driver\hpcu215u.inf` | Relative to the script in the package |
| **Printer name** | `Office Printer 3rd Floor` | Choose a descriptive display name |
| **Printer IP address** | `192.168.1.100` | Network configuration / DHCP reservation |
| **Paper size** | `A4` | Organization standard |
| **Duplex mode** | `TwoSidedLongEdge` | Organization policy |
| **Color mode** | `Color` or `Monochrome` | Organization policy |

The printer name and IP address are **required**. All configuration settings (paper size, duplex, color, collation) are **optional** -- the printer will use the driver's defaults if not specified.

---

## Package Structure

Create a folder with the following structure before packaging with the IntuneWinAppUtil tool:

```
MyPrinter\
+-- Install-Printer.ps1
+-- Uninstall-Printer.ps1
+-- Driver\
    +-- driver.inf
    +-- driver.cat
    +-- *.dll / *.sys / *.cab
```

Package the root folder:

```cmd
IntuneWinAppUtil.exe -c "C:\MyPrinter" -s Install-Printer.ps1 -o "C:\Output"
```

This produces `Install-Printer.intunewin`.

> **Note:** The detection script (`Detect-Printer.ps1`) is uploaded separately in the Intune portal -- it is not included in the `.intunewin` package.

---

## Scripts

### Install-Printer.ps1

Stages the printer driver into the Windows Driver Store, registers it with the Print Spooler, creates a Standard TCP/IP port, creates the printer queue, and optionally configures printer settings.

#### Parameters

| Parameter | Required | Default | Description |
|---|---|---|---|
| `DriverName` | Yes | -- | Exact driver name as declared in the INF file (e.g. `"HP Universal Printing PCL 6"`). |
| `PrinterName` | Yes | -- | Display name for the printer queue (e.g. `"Office Printer 3rd Floor"`). |
| `PrinterIPAddress` | Yes | -- | IP address of the network printer (e.g. `"192.168.1.100"`). |
| `InfPath` | No | Auto-discovered | Path to the driver INF file. Accepts absolute or relative to `$PSScriptRoot`. If omitted, the script searches for any `.inf` file under its own directory. |
| `DriverSourceFolder` | No | Folder containing InfPath | Folder with the INF and all accompanying driver files (SYS, DLL, CAT, etc.). |
| `PortName` | No | `IP_<PrinterIPAddress>` | Custom TCP/IP port name. |
| `PortNumber` | No | `9100` | TCP port number for the printer port. Valid range: 1-65535. |
| `PaperSize` | No | Driver default | Default paper size. See [PaperSize Values](#papersize-values). |
| `DuplexMode` | No | Driver default | Duplex printing mode. See [DuplexMode Values](#duplexmode-values). |
| `ColorMode` | No | Driver default | Color or monochrome printing. See [ColorMode Values](#colormode-values). |
| `Collate` | No | Driver default | Collation setting. See [Collate Values](#collate-values). |
| `PrintTicketXml` | No | -- | Raw PrintTicket XML for advanced driver-specific settings. See [PrintTicketXml](#printticketxml-advanced). |
| `Location` | No | -- | Physical location text (e.g. `"Building A, Room 201"`). |
| `Comment` | No | -- | Description text for the printer. |
| `Shared` | No | `$false` | Share the printer on the network. |
| `ShareName` | No | -- | Network share name (used with `-Shared`). |
| `SNMPEnabled` | No | `$true` | Enable SNMP status monitoring on the TCP/IP port. |
| `SNMPCommunity` | No | `'public'` | SNMP community string for the TCP/IP port. Change when the printer uses a non-default community string (e.g. `'printers'`). Written to the Standard TCP/IP Port monitor registry key after port creation. |
| `SpoolerTimeoutSeconds` | No | `60` | Seconds to wait for the Print Spooler service. Valid range: 10-300. |

#### Intune Install Command Examples

**Minimal** -- driver, printer name, and IP only:

```
powershell.exe -ExecutionPolicy Bypass -File Install-Printer.ps1 -DriverName "HP Universal Printing PCL 6" -InfPath ".\Driver\hpcu215u.inf" -PrinterName "Office Printer" -PrinterIPAddress "192.168.1.100"
```

**With configuration** -- paper size, duplex, and color:

```
powershell.exe -ExecutionPolicy Bypass -File Install-Printer.ps1 -DriverName "HP Universal Printing PCL 6" -InfPath ".\Driver\hpcu215u.inf" -PrinterName "Office Printer" -PrinterIPAddress "192.168.1.100" -PaperSize A4 -DuplexMode TwoSidedLongEdge -ColorMode Color
```

**Full example** -- all options:

```
powershell.exe -ExecutionPolicy Bypass -File Install-Printer.ps1 -DriverName "Kyocera TASKalfa 3554ci KX" -InfPath ".\Driver\kyocera.inf" -DriverSourceFolder ".\Driver" -PrinterName "Finance Printer" -PrinterIPAddress "10.0.1.50" -PaperSize A4 -DuplexMode TwoSidedLongEdge -ColorMode Monochrome -Collate Enabled -Location "Building A, Room 201" -Comment "Finance department printer"
```

#### Installation Steps (what the script does)

1. Resolves and validates the INF file path
2. Resolves the driver source folder
3. Locates `pnputil.exe` (handles 32-bit PowerShell on 64-bit/ARM64 OS via `sysnative`)
4. Ensures the Print Spooler service is running
5. Validates the printer IP address and derives the default port name
6. Checks if the printer already exists (skips install if already present)
7. Stages the driver into the Driver Store with `pnputil /add-driver`
8. Restarts the Print Spooler and waits 5 seconds for the driver to fully register in the Driver Store
9. Registers the driver with `Add-PrinterDriver` (retry loop -- 5 attempts, 10s apart)
10. Creates the Standard TCP/IP port (skips if already exists)
11. Creates the printer queue (skips if already exists)
12. Configures printer settings: location/comment, duplex/color/collation, paper size via PrintTicket XML, and custom PrintTicketXml
13. Verifies all components (driver, port, printer) are present

**Rollback:** If port or printer queue creation fails, the script automatically cleans up any resources it created during the current run. Configuration failures (step 12) do **not** trigger rollback -- the printer is functional, just unconfigured.

**Idempotent:** Running the script a second time with the same parameters succeeds without errors -- existing components are detected and skipped.

---

### Detect-Printer.ps1

Custom detection script for Intune. Checks that both the **printer queue** and the **printer driver** are present. Both must exist for the app to be detected as installed.

#### Configuration

Before uploading to Intune, open the script and set both variables:

```powershell
$PrinterName = 'Office Printer 3rd Floor'
$DriverName  = 'HP Universal Printing PCL 6'
```

These values must match the `-PrinterName` and `-DriverName` parameters used in the install command.

#### Detection Logic

Detection uses **registry keys only** -- the Print Spooler does not need to be running.

| Condition | Intune Result |
|---|---|
| Printer queue in registry **AND** driver in registry | App is **INSTALLED** (detected) |
| Printer queue missing **OR** driver missing | App is **NOT installed** |
| Non-zero exit | Detection **error** |

Registry keys checked:

| Component | Registry path |
|---|---|
| Printer queue | `HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Printers\<PrinterName>` |
| Driver | `HKLM:\SYSTEM\CurrentControlSet\Control\Print\Environments\<arch>\Drivers\Version-3\<DriverName>` |

The script also checks the Driver Store as supplementary diagnostic information (logged to file), but this check is **not required** for detection.

#### Intune Detection Rule Settings

| Setting | Value |
|---|---|
| Rules format | Use a custom detection script |
| Script file | `Detect-Printer.ps1` |
| Run script as 32-bit process on 64-bit clients | **No** |
| Enforce script signature check | No *(unless your org signs scripts)* |

---

### Uninstall-Printer.ps1

Removes the printer queue, TCP/IP port, printer driver, and optionally the OEM INF package from the Windows Driver Store.

Safe-guards applied before removal:

- The port name is captured from the printer object before the queue is deleted
- The port is only removed if no other printer queues reference it
- The driver is only removed if no other printer queues use it
- Any remaining queues using the driver are deleted before driver removal
- The OEM INF entry is located in the Driver Store and deleted with `pnputil /delete-driver`

#### Parameters

| Parameter | Required | Default | Description |
|---|---|---|---|
| `PrinterName` | Yes | -- | Display name of the printer queue to remove. |
| `DriverName` | Yes | -- | Exact driver name as it appears in the Print Spooler. Must match the value used during installation. |
| `KeepDriver` | No | Off | When specified, keeps the printer driver in the Print Spooler after removing the queue. |
| `KeepPort` | No | Off | When specified, keeps the TCP/IP port after removing the queue. |
| `KeepDriverStore` | No | Off | When specified, keeps the OEM INF package in the Driver Store. Use when multiple printer models share the same INF. |
| `SpoolerTimeoutSeconds` | No | `60` | Seconds to wait for the Print Spooler service. Valid range: 10-300. |

#### Intune Uninstall Command Examples

**Standard removal** -- removes printer, port, driver, and Driver Store entry:

```
powershell.exe -ExecutionPolicy Bypass -File Uninstall-Printer.ps1 -PrinterName "Office Printer" -DriverName "HP Universal Printing PCL 6"
```

**Remove printer queue only** -- keep driver and port:

```
powershell.exe -ExecutionPolicy Bypass -File Uninstall-Printer.ps1 -PrinterName "Office Printer" -DriverName "HP Universal Printing PCL 6" -KeepDriver -KeepPort
```

**Keep INF in Driver Store** -- shared INF scenario:

```
powershell.exe -ExecutionPolicy Bypass -File Uninstall-Printer.ps1 -PrinterName "Finance Printer" -DriverName "Kyocera TASKalfa 3554ci KX" -KeepDriverStore
```

#### Uninstallation Steps (what the script does)

1. Ensures the Print Spooler service is running
2. Locates the printer queue and captures its port name
3. Removes the printer queue
4. Removes the TCP/IP port (unless `-KeepPort` is specified, or another printer shares the port)
5. Removes the printer driver from the Print Spooler (unless `-KeepDriver` is specified; falls back to `printui.dll` for stubborn drivers)
6. Removes the OEM INF package from the Driver Store with `pnputil /delete-driver` (unless `-KeepDriverStore` is specified)

**Idempotent:** Running the script when the printer is already absent succeeds without errors (exit 0).

---

## Intune Win32 App Configuration

| Field | Value |
|---|---|
| App type | Windows app (Win32) |
| Package file | `Install-Printer.intunewin` |
| Install command | `powershell.exe -ExecutionPolicy Bypass -File Install-Printer.ps1 -DriverName "..." -InfPath ".\Driver\..." -PrinterName "..." -PrinterIPAddress "..."` |
| Uninstall command | `powershell.exe -ExecutionPolicy Bypass -File Uninstall-Printer.ps1 -PrinterName "..." -DriverName "..."` |
| Install behavior | System |
| Device restart behavior | No specific action |
| Detection rules | Custom script -- upload `Detect-Printer.ps1` |
| Run script as 32-bit process | **No** |
| Enforce script signature check | No |

> **Important:** Always use **double quotes** (`"..."`) around parameter values in the install and uninstall commands. Single quotes (`'...'`) are PowerShell syntax only -- they are not recognized by the Windows command line parser (`CommandLineToArgvW`) and will cause parameter binding errors (exit code 1 / error 0x80070001) when Intune executes the command.

---

## Printer Configuration Reference

All configuration parameters are optional. When omitted, the printer uses the driver's built-in defaults.

### PaperSize Values

| Value | Description |
|---|---|
| `A3` | ISO A3 (297 x 420 mm) |
| `A4` | ISO A4 (210 x 297 mm) |
| `A5` | ISO A5 (148 x 210 mm) |
| `B4` | JIS B4 (257 x 364 mm) |
| `B5` | JIS B5 (182 x 257 mm) |
| `Letter` | US Letter (8.5 x 11 in) |
| `Legal` | US Legal (8.5 x 14 in) |
| `Executive` | US Executive (7.25 x 10.5 in) |
| `Tabloid` | US Tabloid (11 x 17 in) |
| `Statement` | US Statement (5.5 x 8.5 in) |
| `Folio` | Folio (8.5 x 13 in) |
| `Quarto` | Quarto (8.47 x 10.83 in) |
| `Note` | US Note (8.5 x 11 in) |

Paper size is applied by modifying the printer's PrintTicket XML (`PageMediaSize` feature). The driver recalculates the physical dimensions automatically.

### DuplexMode Values

| Value | Description |
|---|---|
| `OneSided` | Single-sided printing (simplex) |
| `TwoSidedLongEdge` | Double-sided, flip on long edge (standard duplex for portrait) |
| `TwoSidedShortEdge` | Double-sided, flip on short edge (standard duplex for landscape) |

### ColorMode Values

| Value | Description |
|---|---|
| `Color` | Color printing enabled |
| `Monochrome` | Black and white only |

### Collate Values

| Value | Description |
|---|---|
| `Enabled` | Collation on (pages printed in order: 1-2-3, 1-2-3) |
| `Disabled` | Collation off (pages grouped: 1-1, 2-2, 3-3) |

### PrintTicketXml (Advanced)

The `-PrintTicketXml` parameter allows you to configure **driver-specific settings** that are not covered by the standard parameters (PaperSize, DuplexMode, ColorMode, Collate). This includes settings like input tray selection, output bin, stapling, hole punching, watermarks, n-up printing, and any other capability the driver exposes.

PrintTicketXml is applied **after** all standard settings, so it can override them if the XML contains the same features.

#### When to use PrintTicketXml

| Use PrintTicketXml | Use standard parameters instead |
|---|---|
| Input tray / output bin selection | Paper size (`-PaperSize A4`) |
| Stapling, hole punching, folding | Duplex mode (`-DuplexMode TwoSidedLongEdge`) |
| Watermarks, overlays | Color mode (`-ColorMode Color`) |
| Secure print / PIN printing | Collation (`-Collate Enabled`) |
| Vendor-specific driver options | Any setting covered by standard params |
| Overriding a standard setting with a driver-specific variant | -- |

**Rule of thumb:** If a standard parameter covers the setting, use it -- it is simpler and more readable. Use PrintTicketXml only for settings that standard parameters cannot configure.

#### How to obtain a PrintTicket XML

**Step 1** -- Install the printer on a test machine with the same driver you will deploy via Intune.

**Step 2** -- Open **Settings > Printers & Scanners**, select the printer, and configure the desired settings (input tray, output bin, stapling, etc.) under **Printing preferences**.

**Step 3** -- Export the current PrintTicket XML with PowerShell:

```powershell
$config = Get-PrintConfiguration -PrinterName "Your Printer Name"
$config.PrintTicketXml | Set-Content -Path "C:\Temp\PrintTicket.xml" -Encoding UTF8
```

**Step 4** -- Open `PrintTicket.xml` in a text editor. The file contains `<psf:Feature>` elements for each configured setting. You can use the **entire file content** as the `-PrintTicketXml` value, or extract only the specific features you need.

#### Real-world examples

**Example 1** -- Set input tray to Tray 2:

```
powershell.exe -ExecutionPolicy Bypass -File Install-Printer.ps1 ^
  -DriverName "HP Universal Printing PCL 6" ^
  -InfPath ".\Driver\hpcu215u.inf" ^
  -PrinterName "Office Printer" ^
  -PrinterIPAddress "192.168.1.100" ^
  -PrintTicketXml "<psf:PrintTicket xmlns:psf='http://schemas.microsoft.com/windows/2003/08/printing/printschemaframework' xmlns:psk='http://schemas.microsoft.com/windows/2003/08/printing/printschemakeywords' version='1'><psf:Feature name='psk:JobInputBin'><psf:Option name='psk:Tray2' /></psf:Feature></psf:PrintTicket>"
```

**Example 2** -- Set output bin to Stacker:

```
powershell.exe -ExecutionPolicy Bypass -File Install-Printer.ps1 ^
  -DriverName "HP Universal Printing PCL 6" ^
  -InfPath ".\Driver\hpcu215u.inf" ^
  -PrinterName "Office Printer" ^
  -PrinterIPAddress "192.168.1.100" ^
  -PrintTicketXml "<psf:PrintTicket xmlns:psf='http://schemas.microsoft.com/windows/2003/08/printing/printschemaframework' xmlns:psk='http://schemas.microsoft.com/windows/2003/08/printing/printschemakeywords' version='1'><psf:Feature name='psk:JobOutputBin'><psf:Option name='psk:Stacker' /></psf:Feature></psf:PrintTicket>"
```

**Example 3** -- Enable edge-to-edge stapling (top-left corner):

```
powershell.exe -ExecutionPolicy Bypass -File Install-Printer.ps1 ^
  -DriverName "Kyocera TASKalfa 3554ci KX" ^
  -InfPath ".\Driver\kyocera.inf" ^
  -PrinterName "Finance Printer" ^
  -PrinterIPAddress "10.0.1.50" ^
  -PrintTicketXml "<psf:PrintTicket xmlns:psf='http://schemas.microsoft.com/windows/2003/08/printing/printschemaframework' xmlns:psk='http://schemas.microsoft.com/windows/2003/08/printing/printschemakeywords' version='1'><psf:Feature name='psk:JobStapleAllDocuments'><psf:Option name='psk:StapleTopLeft' /></psf:Feature></psf:PrintTicket>"
```

**Example 4** -- Combine standard parameters with PrintTicketXml (A4, duplex, color + Tray 2):

```
powershell.exe -ExecutionPolicy Bypass -File Install-Printer.ps1 ^
  -DriverName "HP Universal Printing PCL 6" ^
  -InfPath ".\Driver\hpcu215u.inf" ^
  -PrinterName "Office Printer" ^
  -PrinterIPAddress "192.168.1.100" ^
  -PaperSize A4 ^
  -DuplexMode TwoSidedLongEdge ^
  -ColorMode Color ^
  -PrintTicketXml "<psf:PrintTicket xmlns:psf='http://schemas.microsoft.com/windows/2003/08/printing/printschemaframework' xmlns:psk='http://schemas.microsoft.com/windows/2003/08/printing/printschemakeywords' version='1'><psf:Feature name='psk:JobInputBin'><psf:Option name='psk:Tray2' /></psf:Feature></psf:PrintTicket>"
```

#### Common PrintTicket feature names

These are standard [Print Schema keywords](https://learn.microsoft.com/en-us/windows/win32/printdocs/print-schema) supported by most drivers:

| Feature name | Common options | Description |
|---|---|---|
| `psk:JobInputBin` | `psk:AutoSelect`, `psk:Tray1`, `psk:Tray2`, `psk:Manual` | Input tray |
| `psk:JobOutputBin` | `psk:AutoSelect`, `psk:Stacker`, `psk:Mailbox` | Output bin |
| `psk:JobStapleAllDocuments` | `psk:None`, `psk:StapleTopLeft`, `psk:StapleTopRight` | Stapling |
| `psk:JobHolePunch` | `psk:None`, `psk:LeftEdge`, `psk:TopEdge` | Hole punching |
| `psk:JobNUpAllDocumentsContiguously` | `psk:PagesPerSheet` (with scored property) | N-up printing |
| `psk:PageWatermark` | Vendor-specific | Watermark |

> **Note:** Feature names and option values are driver-dependent. Not every driver supports every feature, and some drivers use vendor-specific names (e.g. `ns0000:CustomTray`). Always export the PrintTicket XML from a working printer with the **same driver** to get the correct names.

> **Tip:** For very long XML strings, you can store the PrintTicket in a file and read it at deployment time. However, for Intune Win32 app deployment, the command line is the simplest approach -- most PrintTicket snippets for a single feature fit on one line.

---

## ARM64 Devices

This toolkit is compatible with **ARM64 Windows devices** (e.g. Snapdragon-based laptops and Surface Pro X). The following notes apply:

### pnputil path resolution

Intune Management Extension (IME) launches install/uninstall scripts as a **32-bit (x86) process**. On both x64 and ARM64 Windows, the scripts automatically resolve `pnputil.exe` via the `sysnative` path (which points to the real native 64-bit/ARM64 `System32`) when running under a 32-bit PowerShell host. No extra configuration is needed.

Each script logs the architecture on startup to help with triage:

```
[INFO] OS arch     : ARM64
[INFO] PS 64-bit   : False
```

### Detection script

The detection script uses registry keys for its primary checks (no Spooler dependency). However, the supplementary Driver Store scan uses `pnputil.exe`, and it is still recommended to run the detection script as a **64-bit/native process** on ARM64 to ensure the correct `Windows ARM64` driver environment is enumerated. Ensure this Intune setting is configured:

| Setting | Value |
|---|---|
| Run script as 32-bit process on 64-bit clients | **No** |

Setting this to **Yes** on an ARM64 device would run the detection script as an x86 process and might not enumerate ARM64 native drivers correctly in the supplementary Driver Store check.

### Driver INF architecture support

The driver INF file must declare architecture support for the target platform. Open the INF and verify the `[SourceDisksFiles.*]` sections cover your deployment targets:

- `[SourceDisksFiles.amd64]` -- required for x64 devices
- `[SourceDisksFiles.arm64]` -- required for ARM64 devices

If the manufacturer's INF does not include an `arm64` section, the driver cannot be staged on ARM64 devices with `pnputil`. Contact the manufacturer for an ARM64-compatible driver package.

### ARM64 driver signing requirements

ARM64 Windows enforces stricter driver signing than x64. Drivers must be **WHQL-signed** (Windows Hardware Quality Labs certification). An EV (Extended Validation) Code Signing certificate alone is **insufficient** on ARM64 -- the driver will stage successfully with `pnputil` but `Add-PrinterDriver` will fail with error `0xE0000242`.

The install script logs a startup warning on ARM64 devices:

```
[WARN] ARM64 device: driver must be WHQL-signed. EV Code Signing alone is insufficient on ARM64 and causes error 0xE0000242.
```

**Resolution:** Obtain a WHQL-certified driver from the manufacturer's website or Windows Update. Most enterprise printers (HP, Brother, Kyocera, Canon, Xerox) publish WHQL-signed ARM64 drivers, but you may need to download a newer driver package specifically labeled for ARM64.

---

## Log Locations

| Script | Log File |
|---|---|
| Install-Printer.ps1 | `%ProgramData%\Microsoft\IntuneManagementExtension\Logs\Install-Printer.log` |
| Detect-Printer.ps1 | `%ProgramData%\Microsoft\IntuneManagementExtension\Logs\Detect-Printer.log` |
| Uninstall-Printer.ps1 | `%ProgramData%\Microsoft\IntuneManagementExtension\Logs\Uninstall-Printer.log` |

All three scripts append to their log with a `=== NEW RUN ===` separator, preserving a rolling history across retries, reinstalls, and detection cycles.

Every run logs the following diagnostic header before any work begins:

```
[INFO] Script version      : 1.7.1
[INFO] Running as          : NT AUTHORITY\SYSTEM
[INFO] Computer name       : DESKTOP-ABC123
[INFO] OS arch             : AMD64
[INFO] OS version          : Microsoft Windows NT 10.0.22621.0
[INFO] PS version          : 5.1.22621.4391
[INFO] PS 64-bit           : False
[INFO] Process ID          : 4812
[INFO] Script root         : C:\Windows\IMECache\{GUID}
```

> **Note:** `PS 64-bit : False` is expected -- IME launches PowerShell as a 32-bit process. Print management cmdlets work correctly from 32-bit PowerShell on Windows 10/11. Only `pnputil.exe` requires native access, which the scripts resolve via `sysnative` automatically.

Install-Printer.ps1 also logs the contents of the Driver folder (file names and sizes), per-step elapsed time markers (after pnputil, Spooler restart, Add-PrinterDriver, port creation, and queue creation), Spooler uptime and PID after restart (used for the smart initialization wait), and a one-line detail dump for the registered driver, created port, and created printer queue.

To view a log from the default Intune log directory:

```powershell
Get-Content "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs\Install-Printer.log"
```

---

## Testing Locally as SYSTEM

Intune runs install, uninstall, and detection scripts as **NT AUTHORITY\SYSTEM** in a 32-bit PowerShell process. Reproducing this context locally is essential for catching permission, path, and bitness issues before deploying to production.

### Using PSExec (Sysinternals)

[PSExec](https://learn.microsoft.com/en-us/sysinternals/downloads/psexec) can launch a process as SYSTEM on the local machine. Download PSExec from the Sysinternals page.

**Step 1** -- Open an elevated (`Run as administrator`) Command Prompt.

**Step 2** -- Launch a 32-bit PowerShell shell as SYSTEM, exactly as IME does:

```cmd
psexec -i -s "C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
```

> `-i` keeps the session interactive so you can see output. `-s` runs as SYSTEM.

**Step 3** -- In the new PowerShell window, verify you are SYSTEM and 32-bit:

```powershell
[System.Security.Principal.WindowsIdentity]::GetCurrent().Name   # -> NT AUTHORITY\SYSTEM
[Environment]::Is64BitProcess                                      # -> False
```

**Step 4** -- Run the install script directly:

```powershell
cd "C:\Path\To\AppInstaller"
.\Install-Printer.ps1 -DriverName "Brother MFC-J6540DW Printer" -InfPath ".\Driver\BRPRI20A.INF" -PrinterName "Brother MFC-J6540DW" -PrinterIPAddress "192.168.1.100"
```

**Step 5** -- Review the log at `%ProgramData%\Microsoft\IntuneManagementExtension\Logs\Install-Printer.log` for the full step-by-step trace.

### Testing the detection script

The detection script should run as a 64-bit process (Intune setting "Run script as 32-bit process" = No). To test it locally:

```powershell
# From a 64-bit PowerShell session running as SYSTEM (or your own user for basic testing):
& "C:\Path\To\Detection\Detect-Printer.ps1"
echo "Exit code: $LASTEXITCODE"
```

Exit 0 with output = detected. Exit 0 with no output = not detected.

---

## Exit Codes

| Code | Meaning |
|---|---|
| `0` | Success (or components were already absent -- idempotent) |
| `1` | Failure -- see the log file for details |
| `259` | `pnputil`: driver already staged / up-to-date -- treated as success |

---

## Troubleshooting

### Printer installs but paper size / duplex is not applied

Configuration settings are applied via `Set-PrintConfiguration` and PrintTicket XML after the printer queue is created. If the driver does not support a particular setting, the script logs a warning but still exits with code 0 (the printer is functional).

**Resolution:** Check the log file for warnings at step 12. Verify the driver supports the requested setting by configuring it manually via Windows Settings > Printers & Scanners.

### Detection script reports "NOT DETECTED" after successful install

The detection script requires **both** the printer queue and the driver registry keys to be present.

**Resolution:**
1. Check the printer registry key (replace with your printer name):
   ```powershell
   Test-Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Printers\Office Printer"
   ```
2. Check the driver registry key (replace with your driver name and architecture):
   ```powershell
   Test-Path "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Environments\Windows x64\Drivers\Version-3\HP Universal Printing PCL 6"
   ```
3. Verify the `$PrinterName` and `$DriverName` values in `Detect-Printer.ps1` match **exactly** (case-sensitive) the values used in the install command. A single character difference results in NOT DETECTED.

### pnputil returns error on ARM64 devices

The scripts handle 32-bit/64-bit path resolution automatically, but the detection script must run as a 64-bit process.

**Resolution:** In the Intune detection rule settings, set "Run script as 32-bit process on 64-bit clients" to **No**.

### Printer queue created but cannot print (network error)

The script does not test network connectivity during installation -- the printer is created regardless. This allows the printer to be pre-deployed before the physical device is connected to the network.

**Resolution:** Verify the printer is powered on and reachable at the configured IP address: `Test-NetConnection -ComputerName "192.168.1.100" -Port 9100`

### Uninstall fails with "driver is in use"

The uninstall script automatically removes printer queues that use the driver before attempting driver removal. If removal still fails, the script falls back to `printui.dll`.

**Resolution:** Check the log for other printer queues using the same driver. Manually remove them with `Remove-Printer -Name "Queue Name"` and re-run the uninstall.

### Add-PrinterDriver fails with HRESULT 0x80070032 / Event ID 817 on every attempt

Event ID 817 in the `Microsoft-Windows-PrintService/Operational` log fires on the very first `Add-PrinterDriver` attempt and all retries fail immediately with the same error. The v1.6.5 fix (omitting `-InfPath`) provides no relief.

This is caused by the Intune policy **Allow Print Spooler to accept client connections** being set to **Disabled**. This writes `RegisterSpoolerRemoteRpcEndPoint = 2` to the registry, which instructs the Spooler to start but **not register its RPC endpoint** with the Windows endpoint mapper. `Add-PrinterDriver` communicates with the Spooler exclusively via CIM over RPC -- with no endpoint registered, every call fails regardless of which code path is used.

**Confirm the cause:**

```powershell
(Get-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers' -Name 'RegisterSpoolerRemoteRpcEndPoint' -ErrorAction SilentlyContinue).RegisterSpoolerRemoteRpcEndPoint
```

A return value of `2` confirms this is the cause.

**Distinguish from Mode A (INF code path restriction -- fixed in v1.6.5):**

| Symptom | Mode A (INF path blocked) | Mode B (endpoint absent) |
|---|---|---|
| `Add-PrinterDriver -Name` (no `-InfPath`) | **Succeeds** | **Fails** |
| Event ID 817 timing | May vary | Fires on first attempt immediately |
| Registry value | `= 1` or absent | `= 2` |

**Resolution:** In the Intune policy that controls printing settings (e.g. `WIN - SC - Device Security - Printing Settings`), change:

```
Administrative Templates > Printers
Allow Print Spooler to accept client connections
From:  Disabled
To:    Enabled
```

**This change is safe when the following five mitigations remain enabled in the same policy** -- together they block the complete PrintNightmare exploit chain:

| Mitigation | Policy setting | What it blocks |
|---|---|---|
| Redirection Guard | Configure Redirection Guard = Enabled | Path-redirection attacks hijacking driver file loading |
| Queue-specific file restriction | Manage processing of Queue-specific files = Limit to Color profiles | Arbitrary `CopyFiles` DLL execution (the core PrintNightmare primitive) |
| Driver install to admins only | Limits print driver installation to Administrators = Enabled | Non-admin driver installation via Point and Print |
| Point and Print server restriction | Point and Print Restrictions = Enabled + trusted servers | Connections to unauthorized print servers |
| RPC authentication | Configure RPC listener settings = Negotiate | Unauthenticated inbound RPC connections to the Spooler |

After the policy change is applied and the endpoint checks in, the next Intune sync will run the install script and `Add-PrinterDriver` will succeed. If the driver was already staged by a previous run (`pnputil` reported "Already exists in the system"), the install will proceed directly to `Add-PrinterDriver` without re-staging the files.

### Add-PrinterDriver fails with HRESULT 0x80070032 -- zero PrintService events (Mode C: TrustedServers policy)

All `Add-PrinterDriver` attempts fail with HRESULT 0x80070032 and **no PrintService events** are written to the Admin or Operational log on any attempt. The Mode B fix (enabling `Allow Print Spooler to accept client connections`) provides no relief. This zero-events signature means the call never reaches the Spooler at all.

**Root cause:** The policy `Users can only point and print to these servers` is set to **Enabled/True** with an empty server list. This configures `PointAndPrint\TrustedServers = 1` + empty `ServerList`, which causes the MSFT_PrinterDriver CIM provider to reject all `Add-PrinterDriver` calls in the provider layer before any Spooler RPC call is made.

**Confirm the cause:**

```powershell
Get-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers\PointAndPrint' -ErrorAction SilentlyContinue | Select-Object TrustedServers, ServerList
```

`TrustedServers = 1` with an empty `ServerList` confirms Mode C.

**Resolution:** In the Intune printing policy (e.g. `WIN - SC - Device Security - Printing Settings`), change:

```
Administrative Templates > Printers
Users can only point and print to these servers
From:  Enabled (True) with empty server list
To:    Disabled (False)
```

After the policy propagates and the device syncs, the next Intune deployment will succeed.

**How to distinguish Mode A, B, and C:**

| Symptom | Mode A | Mode B | Mode C |
|---|---|---|---|
| PrintService Event ID 817 | Fires (with `-InfPath`) | Fires on first attempt | **Never fires** |
| `-Name` only (no `-InfPath`) | Succeeds | Fails | Fails |
| `RegisterSpoolerRemoteRpcEndPoint` | `1` or absent | `2` | `1` or absent |
| `PointAndPrint\TrustedServers` | Any | Any | `1` + empty ServerList |

### Add-PrinterDriver blocked by RPC/security policy (HRESULT 0x80070032 / Event 817) -- INF code path restricted (Mode A)

The Print Spooler needs time to process a newly staged driver. If `Add-PrinterDriver` is called immediately after `pnputil /add-driver`, it may fail with a generic error even though the driver files are in the Driver Store.

**Resolution:** The script automatically restarts the Spooler, waits 5 seconds, and retries `Add-PrinterDriver` up to 5 times (10s apart). If the issue persists, check the log for the specific error message on each attempt.

### Install fails with error 0x80070001 (exit code 1)

Error 0x80070001 maps to Win32 exit code 1. It means PowerShell started but exited with code 1. Check the log file first:

```
%ProgramData%\Microsoft\IntuneManagementExtension\Logs\Install-Printer.log
```

**Common causes:**

1. **Smart quotes or curly quotes in the Intune command.** When you copy the install command from a browser or a PDF into the Intune portal, word processors and browsers often replace straight double quotes (`"`) with curly/smart quotes (`"` `"`). PowerShell and the Windows command line do not recognise curly quotes as string delimiters -- they are passed as literal characters, causing parameter binding to fail. Always type the quotes manually or paste into Notepad first to strip formatting before pasting into the Intune portal.
2. **Single quotes in the Intune command.** Windows only recognises double quotes (`"`) as string delimiters at the command line. Single quotes (`'`) cause arguments with spaces to be split into separate tokens, breaking parameter binding.
3. **Incorrect source folder during packaging.** The IntuneWinAppUtil source folder must be the folder containing `Install-Printer.ps1`, not the project root.
4. **Wrong install behavior.** Must be set to **System**, not User.

**Resolution:** Check the log file. Look for the parameter echo block at the top -- if it is absent, the failure occurred before the script body ran (quoting issue). Verify the install command uses plain straight double quotes, repackage from the correct source folder, and confirm install behavior is System.

### Intune shows "Detection error" (non-zero exit)

The detection script should always exit with code 0. A non-zero exit indicates an unexpected error.

**Resolution:** Check the detection log at `%ProgramData%\Microsoft\IntuneManagementExtension\Logs\Detect-Printer.log` for error details. As of v1.2.0 the detection script uses registry-based checks and does not depend on the Print Spooler. Common causes include permission errors accessing the registry (unusual for SYSTEM) or a misconfiguration of the `$PrinterName` / `$DriverName` variables that triggers the placeholder guard.

### pnputil succeeds (exit 0) but Add-PrinterDriver still fails -- Windows Protected Print Mode (Windows 11 24H2+)

`pnputil /add-driver` stages the INF successfully (exit 0, driver appears in Driver Store), but `Add-PrinterDriver` fails on every attempt across all 5 retries. No HRESULT 0x80070032 from the RPC path -- instead the driver simply cannot be registered with the Spooler.

As of v1.7.0 the script checks for this policy **at startup** and logs a warning before any driver work begins:

```
[WARN] POLICY BLOCKER DETECTED: Windows Protected Print Mode is ENABLED (EnableWindowsProtectedPrint = 1).
[WARN] All third-party printer drivers are blocked on Windows 11 24H2+. Add-PrinterDriver will fail.
[WARN] Fix: set "Configure Windows protected print" to Disabled in your Intune printing policy.
```

**Root cause:** `Configure Windows protected print` is set to **Enabled** in the Intune printing policy. On Windows 11 24H2 and later, this mode restricts printing to a curated subset of Microsoft inbox drivers. All third-party manufacturer drivers (Brother, HP, Kyocera, Xerox, Canon, Ricoh, etc.) are blocked from registration with the Spooler regardless of driver signing or pnputil success.

**Confirm the cause:**

```powershell
(Get-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers\WPP' -Name 'EnableWindowsProtectedPrint' -ErrorAction SilentlyContinue).EnableWindowsProtectedPrint
```

A return value of `1` confirms Windows Protected Print Mode is enforced by policy.

**Resolution:** In the Intune printing policy (e.g. `WIN - SC - Device Security - Printing Settings`), change:

```
Administrative Templates > Printers
Configure Windows protected print
From:  Enabled
To:    Disabled
```

This is an independent blocker. It can occur on a device that also has Mode B or Mode C issues -- fix all applicable policy settings together before re-triggering the Intune deployment.