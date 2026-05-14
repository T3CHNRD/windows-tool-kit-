# T3CHNRD'S Windows Tool Kit

PowerShell-based desktop toolkit with a unified WinForms launcher for maintenance, repair, cleanup, updates, deployment workflows, and imported legacy scripts.

## What This Project Includes
- Desktop UI launcher (`ToolkitLauncher.ps1`)
- Background execution model (`BackgroundWorker`) so UI stays responsive
- Progress bar and status updates for each task
- Unified logging to `Logs/`
- Core maintenance tasks:
  - Network diagnostics and network stack reset
  - KillerTools-inspired local network scan, domain lookup, and MAC vendor lookup
  - Temp/junk cleanup and C: drive free-space cleanup
  - Disk space monitoring
  - Data Transfer Wizard using `robocopy` for old-to-new PC or external-drive transfers
  - Clone Disk Wizard that inventories disks, saves a source/destination clone plan, and launches Windows disk/imaging workflows
  - Update all apps with `winget`
  - Windows repair checks (`DISM` + `SFC`)
  - Debloat helper (KillerTools DEBLOAT-inspired installed app inventory export)
  - Security audit tools including Defender, privacy, persistence, browser-extension, encryption/hash, BitLocker, Secure Boot 2023 certificate update, and open-port reviews
- External integrations:
  - `mallockey/Install-Microsoft365` wrapper
  - Microsoft official Media Creation Tool workflow wrapper
- Imported scripts from ZIP preserved in `LegacyScripts/` and exposed in launcher

## Credits
- Toolkit project owner: T3CHNRD
- KillerTools and KillerScan inspiration/reference: https://killertools.net/ and https://github.com/SteveTheKiller/killer-tools-site
- Install-Microsoft365 integration reference: https://github.com/mallockey/Install-Microsoft365
- Windows Media Creation Tool workflow: Microsoft official documentation and download pages

## Folder Structure
```text
T3CHNRD'S Windows Tool Kit/
|- ToolkitLauncher.ps1
|- README.md
|- PROJECT_PLAN.md
|- Build/
|  |- Build-PortableExe.ps1
|  |- Build-Msi.ps1
|  `- Installer/Product.wxs
|- Config/Toolkit.Settings.psd1
|- Modules/MaintenanceToolkit/
|- Integrations/
|- Scripts/Tasks/
|- LegacyScripts/
|- Logs/
`- Docs/
```

## Quick Start
```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
cd "<project-root>"
.\ToolkitLauncher.ps1
```

## Build Outputs
### Portable EXE (USB recommended)
Outputs:
- `dist\portable\T3CHNRD'S Windows Tool Kit\T3CHNRD'S Windows Tool Kit.exe`
- `dist\T3CHNRD'S Windows Tool Kit-portable.zip`

### MSI (optional)
Requires WiX Toolset v3.11.
```powershell
cd "<project-root>\Build"
.\Build-Msi.ps1
```
Output:
- `dist\msi\T3CHNRD'S Windows Tool Kit.msi`

## USB Notes
- Copy the full portable folder to USB, not just the EXE.
- Run the EXE from inside that folder so relative paths resolve.
- For admin-required tasks, run from an elevated context.

## Secure Boot Report Note
The Secure Boot 2023 Certificate Update task is integrated under `Security`. The toolkit copy writes its report and local log to `Desktop\T3CHNRD-SecureBoot2023` instead of the original `P:` drive path.
