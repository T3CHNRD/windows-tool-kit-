Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $root 'Modules\StorageTools\StorageTools.psm1') -Force
Write-Output 'Opening Data Transfer Wizard. It will detect external drives, or let you choose a network/source folder.'
$result = Show-TtkDataTransferWizard
Write-Output "Transfer wizard result: $($result.Summary)"
if ($result.LogPath) { Write-Output "Robocopy log: $($result.LogPath)" }
if ($result.Completed) { exit 0 }
if ($result.Cancelled) { Write-Output 'Transfer wizard closed without starting a transfer.'; exit 0 }
exit 1
