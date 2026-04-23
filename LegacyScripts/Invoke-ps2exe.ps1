<#
.SYNOPSIS
    Builder script - Simple Version
    - Compiles Launch_OneNote.ps1 into a no-console EXE using splatting.
    - Uses local desktop paths for reliability.
    - Creates a desktop shortcut pointing to the new EXE.
#>

$ErrorActionPreference = "Stop"

# 1. Verify ps2exe is installed
if (-not (Get-Module -ListAvailable ps2exe) -and -not (Get-Command "ps2exe" -ErrorAction SilentlyContinue)) {
    Write-Host "ps2exe not found. Installing (CurrentUser)..." -ForegroundColor Yellow
    Install-Module -Name ps2exe -Scope CurrentUser -Force
}

# 2. Define Paths & Setup (Removed runtime40 to fix error)
$setup = @{
    inputFile  = "C:\Users\tvedders\Desktop\personal PS scripts\Launch_OneNote.ps1"
    outputFile = "C:\Users\tvedders\Desktop\ALBLCO IT Documentation.exe"
    iconFile   = "C:\Users\tvedders\Desktop\personal PS scripts\aluminum_blanking_32x32.ico"
    noConsole  = $true
}

# Ensure source files exist before trying to compile
if (!(Test-Path $setup.inputFile)) { throw "Source script not found: $($setup.inputFile)" }
if (!(Test-Path $setup.iconFile))  { throw "Icon file not found: $($setup.iconFile)" }

# 3. Run Compilation
Write-Host "Compiling EXE with embedded icon..." -ForegroundColor Cyan
ps2exe @setup

# 4. Create Desktop Shortcut
Write-Host "Creating Desktop Shortcut..." -ForegroundColor Cyan
try {
    $DesktopPath  = [Environment]::GetFolderPath("Desktop")
    $ShortcutPath = Join-Path $DesktopPath "ALBLCO IT Documentation.lnk"
    
    $WshShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($ShortcutPath)
    
    $Shortcut.TargetPath       = $setup.outputFile
    $Shortcut.WorkingDirectory = Split-Path $setup.outputFile
    $Shortcut.IconLocation     = "$($setup.outputFile),0"
    
    $Shortcut.Save()
    Write-Host "Compilation complete and Shortcut created on Desktop." -ForegroundColor Green
}
catch {
    Write-Host "EXE created, but failed to create shortcut: $($_.Exception.Message)" -ForegroundColor Yellow
}

Write-Host "Done." -ForegroundColor Cyan