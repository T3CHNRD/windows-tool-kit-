# Windows Maintenance Toolkit

PowerShell-based desktop toolkit with a unified WinForms launcher for maintenance, repair, cleanup, updates, deployment workflows, and imported legacy scripts.

## What This Project Includes
- Modern desktop UI launcher (`ToolkitLauncher.ps1`)
- Background execution model using `BackgroundWorker` so UI stays responsive
- Progress bar and status updates for each task
- Unified logging to `Logs/`
- Core maintenance tasks:
  - Network diagnostics
  - Network stack reset
  - Temp/junk cleanup
  - C: drive free-space cleanup
  - Disk space monitoring
  - Bulk app updates (`winget`)
  - Windows health checks (`DISM` + `SFC`)
  - Debloat helper (installed app inventory export)
- External integrations:
  - `mallockey/Install-Microsoft365` download + launch wrapper
  - Microsoft official Media Creation Tool workflow helper
- Imported scripts from ZIP preserved in `LegacyScripts/` and exposed in launcher

## Folder Structure
```text
windows script toolkit/
├─ ToolkitLauncher.ps1
├─ PROJECT_PLAN.md
├─ README.md
├─ Build/
│  ├─ Build-PortableExe.ps1
│  ├─ Build-Msi.ps1
│  └─ Installer/
│     └─ Product.wxs
├─ Config/
│  └─ Toolkit.Settings.psd1
├─ Modules/
│  └─ MaintenanceToolkit/
│     ├─ MaintenanceToolkit.psd1
│     └─ MaintenanceToolkit.psm1
├─ Integrations/
│  ├─ Invoke-InstallMicrosoft365.ps1
│  └─ Invoke-MediaCreationWorkflow.ps1
├─ Scripts/
│  └─ Tasks/
│     ├─ Invoke-NetworkMaintenance.ps1
│     ├─ Invoke-ResetNetworkStack.ps1
│     ├─ Invoke-ClearTempJunk.ps1
│     ├─ Invoke-FreeCDriveSpace.ps1
│     ├─ Invoke-UpdateAllApps.ps1
│     ├─ Invoke-WindowsRepairChecks.ps1
│     ├─ Invoke-DiskSpaceMonitor.ps1
│     └─ Invoke-DebloatInventory.ps1
├─ LegacyScripts/
│  └─ (14 imported scripts from ZIP)
├─ Logs/
└─ Docs/
   ├─ Script-Adaptation-Notes.md
   └─ Screenshots/
```

## Requirements
- Windows 10/11
- Windows PowerShell 5.1+
- `winget` (for app update task)
- Admin PowerShell session for admin-required tasks
- Internet access for integration downloads

## Quick Start
1. Open PowerShell.
2. Optional execution policy for current session:
   ```powershell
   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
   ```
3. Launch toolkit:
   ```powershell
   cd "C:\Users\tvedders\Documents\windows script toolkit"
   .\ToolkitLauncher.ps1
   ```

## Build for USB or Installer
### Portable EXE (recommended for USB)
Builds a folder containing `WindowsMaintenanceToolkit.exe` plus all required subfolders, and also creates a portable zip.
```powershell
cd "C:\Users\tvedders\Documents\windows script toolkit\Build"
.\Build-PortableExe.ps1
```
Output:
- `dist\portable\WindowsMaintenanceToolkit\WindowsMaintenanceToolkit.exe`
- `dist\WindowsMaintenanceToolkit-portable.zip`

### MSI Installer (optional)
Requires WiX Toolset v3.11 installed.
```powershell
cd "C:\Users\tvedders\Documents\windows script toolkit\Build"
.\Build-Msi.ps1
```
Output:
- `dist\msi\WindowsMaintenanceToolkit.msi`

## USB Usage Notes
- Copy the full portable folder to the USB drive, not only the `.exe`.
- Run the `.exe` from that folder so relative paths to modules/scripts resolve correctly.
- For admin-required tasks, launch the EXE from an elevated PowerShell/Explorer context.

## Running as Administrator
Some tasks require elevation. The UI indicates this in the `Requires Admin` column and prevents execution in a non-elevated session.

## Logging
- Daily logs are written to:
  - `Logs\Toolkit-YYYYMMDD.log`
- UI also shows live execution logs.

## Notes on External Integrations
### Microsoft 365 Installer
- Uses `mallockey/Install-Microsoft365` GitHub project ZIP download.
- Wrapper script: `Integrations\Invoke-InstallMicrosoft365.ps1`

### Windows Installation Media Workflow
- Uses Microsoft official URLs:
  - Microsoft support workflow page for installation media
  - Windows Media Creation Tool direct official link
- Wrapper script: `Integrations\Invoke-MediaCreationWorkflow.ps1`

## Extending the Toolkit
Add new tasks in `Modules\MaintenanceToolkit\MaintenanceToolkit.psm1`:
1. Add a new task object in `Get-ToolkitTaskCatalog`.
2. Implement task handler with clear progress steps.
3. Mark `RequiresAdmin = $true` when needed.
4. Reuse shared helpers (`Invoke-ToolkitCommand`, `Invoke-TaskStep`, logging).

## GitHub Push Checklist
1. Verify project files:
   ```powershell
   Get-ChildItem -Recurse
   ```
2. Initialize and push:
   ```powershell
   git init
   git add .
   git commit -m "Initial Windows Maintenance Toolkit"
   git remote add origin <your-repo-url>
   git push -u origin main
   ```
