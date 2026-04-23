[CmdletBinding()]
param(
    [string]$OutputRoot = (Join-Path $PSScriptRoot '..\dist'),
    [string]$AppName = "T3CHNRD'S Windows Tool Kit",
    [switch]$SkipPortableBuild
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$outputRootResolved = Resolve-Path $OutputRoot -ErrorAction SilentlyContinue
if (-not $outputRootResolved) {
    New-Item -Path $OutputRoot -ItemType Directory -Force | Out-Null
    $outputRootResolved = Resolve-Path $OutputRoot
}

$portableRoot = Join-Path $outputRootResolved.Path "portable\$AppName"
$msiOutDir = Join-Path $outputRootResolved.Path 'msi'
$objDir = Join-Path $outputRootResolved.Path 'obj'
$installerDir = Join-Path $PSScriptRoot 'Installer'
$productWxs = Join-Path $installerDir 'Product.wxs'
$harvestWxs = Join-Path $installerDir 'HarvestedFiles.wxs'

if (-not $SkipPortableBuild) {
    & (Join-Path $PSScriptRoot 'Build-PortableExe.ps1') -OutputRoot $OutputRoot -AppName $AppName
}

if (-not (Test-Path $portableRoot)) {
    throw "Portable app root not found: $portableRoot"
}

if (-not (Test-Path (Join-Path $portableRoot "$AppName.exe"))) {
    throw "Expected EXE not found under $portableRoot"
}

if (-not (Test-Path $msiOutDir)) { New-Item -Path $msiOutDir -ItemType Directory -Force | Out-Null }
if (-not (Test-Path $objDir)) { New-Item -Path $objDir -ItemType Directory -Force | Out-Null }

function Get-WixToolPath {
    param([Parameter(Mandatory = $true)][string]$ToolName)

    $candidates = @()
    if ($env:WIX) {
        $candidates += (Join-Path $env:WIX "$ToolName.exe")
    }
    if (${env:ProgramFiles(x86)}) {
        $candidates += (Join-Path ${env:ProgramFiles(x86)} "WiX Toolset v3.11\bin\$ToolName.exe")
    }
    if ($env:ProgramFiles) {
        $candidates += (Join-Path $env:ProgramFiles "WiX Toolset v3.11\bin\$ToolName.exe")
    }

    $candidates = @($candidates | Where-Object { $_ -and (Test-Path $_) })

    if (@($candidates).Count -gt 0) {
        return @($candidates)[0]
    }

    $cmd = Get-Command "$ToolName.exe" -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }

    throw "WiX tool '$ToolName.exe' not found. Install WiX v3.11 first: https://wixtoolset.org/releases/"
}

$heat = Get-WixToolPath -ToolName 'heat'
$candle = Get-WixToolPath -ToolName 'candle'
$light = Get-WixToolPath -ToolName 'light'

Write-Host 'Harvesting files for MSI...'
& $heat dir $portableRoot -cg CG_MaintToolkit -dr INSTALLFOLDER -srd -gg -var var.SourceDir -out $harvestWxs

Write-Host 'Compiling MSI sources...'
$productWixobj = Join-Path $objDir 'Product.wixobj'
$harvestWixobj = Join-Path $objDir 'HarvestedFiles.wixobj'

& $candle -dSourceDir="$portableRoot" -out $objDir\ $productWxs $harvestWxs

if (-not (Test-Path $productWixobj) -or -not (Test-Path $harvestWixobj)) {
    throw 'WiX candle step did not produce expected .wixobj files.'
}

$msiPath = Join-Path $msiOutDir "$AppName.msi"
Write-Host 'Linking MSI package...'
& $light -ext WixUIExtension -out $msiPath $productWixobj $harvestWixobj

if (-not (Test-Path $msiPath)) {
    throw 'MSI output file was not created.'
}

Write-Host ''
Write-Host "MSI build complete: $msiPath"
