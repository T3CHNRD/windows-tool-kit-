[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$toolkitRoot = Split-Path -Parent $PSScriptRoot
$settingsPath = Join-Path $toolkitRoot 'Config\Toolkit.Settings.psd1'
$settings = Import-PowerShellDataFile -Path $settingsPath

$downloadUrl = $settings.Integrations.Microsoft365RepoZip
$targetRoot = Join-Path $toolkitRoot 'Integrations\Install-Microsoft365'
$installerDownloadPath = Join-Path $targetRoot 'OfficeDeploymentTool'
$zipPath = Join-Path $env:TEMP ("Install-Microsoft365-{0}.zip" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))

# FIX: MED-15 - use the maintained mallockey workflow repo instead of a hardcoded ODT binary URL.
Write-Output "Downloading Install-Microsoft365 from $downloadUrl"
Invoke-WebRequest -Uri $downloadUrl -OutFile $zipPath -UseBasicParsing
Write-Output "Downloaded workflow archive to $zipPath"

if (Test-Path $targetRoot) {
    Write-Output "Removing previous workflow copy from $targetRoot"
    Remove-Item -Path $targetRoot -Recurse -Force
}
New-Item -Path $targetRoot -ItemType Directory -Force | Out-Null

Write-Output "Extracting Install-Microsoft365 workflow to $targetRoot"
Expand-Archive -Path $zipPath -DestinationPath $targetRoot -Force
Remove-Item -Path $zipPath -Force -ErrorAction SilentlyContinue

$candidateScript = Get-ChildItem -Path $targetRoot -Filter 'Install-Microsoft365.ps1' -Recurse |
    Select-Object -First 1

if (-not $candidateScript) {
    throw 'Could not locate Install-Microsoft365.ps1 in the downloaded mallockey workflow.'
}

New-Item -Path $installerDownloadPath -ItemType Directory -Force | Out-Null

Write-Output "Launching installer script in a visible PowerShell window: $($candidateScript.FullName)"
Write-Output "Office Deployment Tool download/work folder: $installerDownloadPath"
Write-Output 'The visible Install-Microsoft365 window is the mallockey workflow. Use that window for prompts and install progress.'
$arguments = @(
    '-NoProfile',
    '-ExecutionPolicy', 'Bypass',
    '-File', ('"{0}"' -f $candidateScript.FullName),
    '-DisplayInstall',
    '-OfficeInstallDownloadPath', ('"{0}"' -f $installerDownloadPath)
) -join ' '

# FIX: MARKET-03 - do not wait on the interactive mallockey installer window.
# The toolkit task should report that the official workflow launched, while the
# visible installer window owns the interactive prompts and install lifecycle.
$process = Start-Process -FilePath 'powershell.exe' -ArgumentList $arguments -PassThru -WindowStyle Normal
Write-Output "Install-Microsoft365 workflow launched in process ID $($process.Id)."
Write-Output 'Follow the instructions in the visible Microsoft 365 installer window.'
