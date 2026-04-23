[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Write-Output 'Step 1/3: DISM ScanHealth'
dism.exe /Online /Cleanup-Image /ScanHealth | Out-String | Write-Output

Write-Output 'Step 2/3: DISM RestoreHealth'
dism.exe /Online /Cleanup-Image /RestoreHealth | Out-String | Write-Output

Write-Output 'Step 3/3: SFC ScanNow'
sfc.exe /scannow | Out-String | Write-Output

Write-Output 'Windows repair checks completed.'
