[CmdletBinding()]
param(
    [string]$SkipSelectionFile
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'VendorUpdate.Common.ps1')

Write-Output 'Starting Windows Update workflow.'
if ($SkipSelectionFile) {
    Write-Output ("Using skip selection file: {0}" -f $SkipSelectionFile)
}

Invoke-WindowsUpdateInstallation -SkipSelectionFile $SkipSelectionFile
Write-Output 'Windows Update workflow completed.'
