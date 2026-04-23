# T3CHNRD'S Windows Tool Kit

PowerShell-based desktop toolkit with a unified WinForms launcher for maintenance, repair, cleanup, updates, deployment workflows, and imported legacy scripts.

## What This Project Includes
- Desktop UI launcher (`ToolkitLauncher.ps1`)
- Background execution model (`BackgroundWorker`) so UI stays responsive
- Progress bar and status updates for each task
- Unified logging to `Logs/`
- Core maintenance tasks:
  - Network diagnostics and network stack reset
  - Temp/junk cleanup and C: drive free-space cleanup
  - Disk space monitoring
  - Update all apps with `winget`
  - Windows repair checks (`DISM` + `SFC`)
  - Debloat helper (installed app inventory export)
- External integrations:
  - `mallockey/Install-Microsoft365` wrapper
  - Microsoft official Media Creation Tool workflow wrapper
- Imported scripts from ZIP preserved in `LegacyScripts/` and exposed in launcher

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
cd "C:\Users\tvedders\Documents\T3CHNRD'S Windows Tool Kit"
.\ToolkitLauncher.ps1
```

## Build Outputs
### Portable EXE (USB recommended)
```powershell
cd "C:\Users\tvedders\Documents\T3CHNRD'S Windows Tool Kit\Build"
.\Build-PortableExe.ps1
```
Outputs:
- `dist\portable\T3CHNRD'S Windows Tool Kit\T3CHNRD'S Windows Tool Kit.exe`
- `dist\T3CHNRD'S Windows Tool Kit-portable.zip`

### MSI (optional)
Requires WiX Toolset v3.11.
```powershell
cd "C:\Users\tvedders\Documents\T3CHNRD'S Windows Tool Kit\Build"
.\Build-Msi.ps1
```
Output:
- `dist\msi\T3CHNRD'S Windows Tool Kit.msi`

## USB Notes
- Copy the full portable folder to USB, not just the EXE.
- Run the EXE from inside that folder so relative paths resolve.
- For admin-required tasks, run from an elevated context.
