[CmdletBinding()]
param(
    [string]$OutputRoot,
    [string]$AppName = "T3CHNRD'S Windows Tool Kit"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (-not $OutputRoot) {
    $OutputRoot = Join-Path $PSScriptRoot '..\dist'
}

$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$distRoot = (Resolve-Path $OutputRoot -ErrorAction SilentlyContinue)
if (-not $distRoot) {
    New-Item -Path $OutputRoot -ItemType Directory -Force | Out-Null
    $distRoot = Resolve-Path $OutputRoot
}
$portableRoot = Join-Path $distRoot.Path 'portable'
$appRoot = Join-Path $portableRoot $AppName
$exePath = Join-Path $appRoot "$AppName.exe"
$rootExePath = Join-Path $projectRoot "$AppName.exe"
$fallbackRootExePath = Join-Path $projectRoot "$AppName (updated).exe"
$zipPath = Join-Path $distRoot.Path "$AppName-portable.zip"

Write-Host "Project root: $projectRoot"
Write-Host "Portable root: $appRoot"

if (Test-Path $appRoot) {
    Remove-Item -LiteralPath $appRoot -Recurse -Force
}
New-Item -Path $appRoot -ItemType Directory -Force | Out-Null

$copyItems = @(
    'Config',
    'Docs',
    'Integrations',
    'LegacyScripts',
    'Logs',
    'Modules',
    'Scripts',
    'ToolkitLauncher.ps1',
    'README.md',
    'PROJECT_PLAN.md'
)

foreach ($item in $copyItems) {
    $source = Join-Path $projectRoot $item
    if (Test-Path $source) {
        Copy-Item -Path $source -Destination (Join-Path $appRoot $item) -Recurse -Force
    }
}

if (-not (Get-Module -ListAvailable -Name ps2exe)) {
    Write-Host 'Installing ps2exe module (CurrentUser)...'
    Install-Module -Name ps2exe -Scope CurrentUser -Force -AllowClobber
}

Import-Module ps2exe -Force

$launcherPs1 = Join-Path $appRoot 'ToolkitLauncher.ps1'
if (-not (Test-Path $launcherPs1)) {
    throw "Launcher script missing at $launcherPs1"
}

Write-Host 'Compiling launcher to EXE...'
ps2exe -inputFile $launcherPs1 -outputFile $exePath -noConsole -title "T3CHNRD'S Windows Tool Kit" -version '1.0.0.0'

try {
    Copy-Item -LiteralPath $exePath -Destination $rootExePath -Force
    Write-Host "Root EXE updated: $rootExePath"
}
catch {
    Write-Warning "Could not overwrite the root EXE. It is likely still open: $rootExePath"
    Copy-Item -LiteralPath $exePath -Destination $fallbackRootExePath -Force
    Write-Host "Updated EXE copied instead to: $fallbackRootExePath"
}

if (Test-Path $zipPath) {
    Remove-Item -LiteralPath $zipPath -Force
}
Compress-Archive -Path "$appRoot\*" -DestinationPath $zipPath -Force

Write-Host ''
Write-Host "Build complete."
Write-Host "Portable folder: $appRoot"
Write-Host "Portable EXE: $exePath"
Write-Host "Root EXE: $rootExePath"
if (Test-Path -LiteralPath $fallbackRootExePath) {
    Write-Host "Fallback root EXE: $fallbackRootExePath"
}
Write-Host "Portable ZIP: $zipPath"
