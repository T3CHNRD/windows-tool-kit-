[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'VendorUpdate.Common.ps1')

Write-Output 'Starting vendor BIOS update workflow.'
Write-Output 'Supported manufacturers for BIOS automation: Dell, HP, Lenovo, Framework.'
Invoke-VendorMaintenanceUpdate -Mode BIOS -PromptForManufacturer
