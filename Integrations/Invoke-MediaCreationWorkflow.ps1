[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$toolkitRoot = Split-Path -Parent $PSScriptRoot
$settingsPath = Join-Path $toolkitRoot 'Config\Toolkit.Settings.psd1'
$settings = Import-PowerShellDataFile -Path $settingsPath

$supportPage = $settings.Integrations.WindowsMediaSupportPage
$mediaToolUrl = $settings.Integrations.Windows10MediaToolUrl

$downloadTarget = Join-Path ([Environment]::GetFolderPath('Downloads')) 'MediaCreationTool22H2.exe'

Write-Host "Downloading official Microsoft Media Creation Tool to $downloadTarget"
Invoke-WebRequest -Uri $mediaToolUrl -OutFile $downloadTarget -UseBasicParsing

Write-Host "Opening official Microsoft support workflow page: $supportPage"
Start-Process $supportPage

Write-Host 'Starting Media Creation Tool...'
Start-Process -FilePath $downloadTarget
