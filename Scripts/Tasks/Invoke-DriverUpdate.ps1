[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'VendorUpdate.Common.ps1')

Write-Output 'Starting vendor driver update workflow.'
Write-Output 'Supported manufacturers for driver automation: Dell, HP, Lenovo, Framework.'
Invoke-VendorMaintenanceUpdate -Mode Drivers -PromptForManufacturer
