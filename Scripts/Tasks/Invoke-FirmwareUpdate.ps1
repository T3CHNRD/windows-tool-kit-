[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'VendorUpdate.Common.ps1')

Write-Output 'Starting vendor firmware update workflow.'
Write-Output 'Supported manufacturers for firmware automation: Dell, HP, Lenovo, Framework.'
Invoke-VendorMaintenanceUpdate -Mode Firmware -PromptForManufacturer
