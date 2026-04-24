[CmdletBinding()]
param(
    [string]$SelectionFile
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$toolkitRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$logDir = Join-Path $toolkitRoot 'Logs'
if (-not (Test-Path $logDir)) {
    New-Item -Path $logDir -ItemType Directory -Force | Out-Null
}

$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$wingetOut = Join-Path $logDir "InstalledApps-winget-$stamp.txt"
$appxOut = Join-Path $logDir "InstalledApps-appx-$stamp.csv"
$selectedOut = Join-Path $logDir "SelectedApps-review-$stamp.txt"

Write-Output 'Collecting installed desktop apps (winget)...'
winget.exe list | Out-File -FilePath $wingetOut -Encoding utf8

Write-Output 'Collecting AppX packages...'
Get-AppxPackage | Select-Object Name, PackageFullName, Publisher, Version |
    Export-Csv -Path $appxOut -NoTypeInformation -Encoding utf8

if ($SelectionFile -and (Test-Path $SelectionFile)) {
    Write-Output 'Loading selected app review list...'
    $selectedApps = @(Get-Content -LiteralPath $SelectionFile | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($selectedApps.Count -gt 0) {
        "Selected apps for review:" | Out-File -FilePath $selectedOut -Encoding utf8
        $selectedApps | Sort-Object -Unique | Out-File -FilePath $selectedOut -Encoding utf8 -Append
        Write-Output "Saved selected app review list: $selectedOut"
    }
    else {
        Write-Output 'No apps were selected in the UI review list.'
    }
}

Write-Output 'Opening Apps & Features for safe uninstall review...'
try {
    Start-Process 'ms-settings:appsfeatures' -ErrorAction Stop
}
catch {
    Write-Output "Could not open Apps & Features automatically: $($_.Exception.Message)"
}

Write-Output "Debloat inventory complete. Reports:"
Write-Output $wingetOut
Write-Output $appxOut
if (Test-Path $selectedOut) {
    Write-Output $selectedOut
}
