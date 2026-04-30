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

Write-Output "Downloading official Microsoft Media Creation Tool to $downloadTarget"
Invoke-WebRequest -Uri $mediaToolUrl -OutFile $downloadTarget -UseBasicParsing

# FIX: MED-16 - the Media Creation Tool is intentionally interactive; launch it visibly.
Write-Output "Opening official Microsoft support workflow page: $supportPage"
Start-Process $supportPage

Write-Output 'Starting Media Creation Tool in a visible wizard window. Complete the wizard there, then return to the toolkit.'
Start-Process -FilePath $downloadTarget
