[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$toolkitRoot = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
$logDir = Join-Path $toolkitRoot 'Logs'
if (-not (Test-Path $logDir)) {
    New-Item -Path $logDir -ItemType Directory -Force | Out-Null
}

$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$wingetOut = Join-Path $logDir "InstalledApps-winget-$stamp.txt"
$appxOut = Join-Path $logDir "InstalledApps-appx-$stamp.csv"

Write-Output 'Collecting installed desktop apps (winget)...'
winget.exe list | Out-File -FilePath $wingetOut -Encoding utf8

Write-Output 'Collecting AppX packages...'
Get-AppxPackage | Select-Object Name, PackageFullName, Publisher, Version |
    Export-Csv -Path $appxOut -NoTypeInformation -Encoding utf8

Write-Output 'Opening Apps & Features for safe uninstall review...'
Start-Process 'ms-settings:appsfeatures'

Write-Output "Debloat inventory complete. Reports:"
Write-Output $wingetOut
Write-Output $appxOut
