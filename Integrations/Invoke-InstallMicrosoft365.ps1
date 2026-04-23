[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$toolkitRoot = Split-Path -Parent $PSScriptRoot
$settingsPath = Join-Path $toolkitRoot 'Config\Toolkit.Settings.psd1'
$settings = Import-PowerShellDataFile -Path $settingsPath

$downloadUrl = $settings.Integrations.Microsoft365RepoZip
$targetRoot = Join-Path $toolkitRoot 'Integrations\Install-Microsoft365'
$zipPath = Join-Path $env:TEMP ("Install-Microsoft365-{0}.zip" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))

Write-Host "Downloading Install-Microsoft365 from $downloadUrl"
Invoke-WebRequest -Uri $downloadUrl -OutFile $zipPath -UseBasicParsing

if (Test-Path $targetRoot) {
    Remove-Item -Path $targetRoot -Recurse -Force
}
New-Item -Path $targetRoot -ItemType Directory -Force | Out-Null

Expand-Archive -Path $zipPath -DestinationPath $targetRoot -Force
Remove-Item -Path $zipPath -Force -ErrorAction SilentlyContinue

$candidateScript = Get-ChildItem -Path $targetRoot -Filter '*.ps1' -Recurse |
    Where-Object { $_.Name -match 'install|office|microsoft365' } |
    Select-Object -First 1

if (-not $candidateScript) {
    throw 'Could not locate an installer entry script in Install-Microsoft365.'
}

Write-Host "Launching installer script: $($candidateScript.FullName)"
Write-Host 'Follow the script prompts to choose Microsoft 365 install options.'
powershell.exe -NoProfile -ExecutionPolicy Bypass -File $candidateScript.FullName
