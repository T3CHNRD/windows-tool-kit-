[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Write-Output 'Starting network maintenance diagnostics...'
Write-Output 'Step 1/4: IP configuration'
ipconfig.exe /all | Out-String | Write-Output

Write-Output 'Step 2/4: DNS cache display'
ipconfig.exe /displaydns | Out-String | Write-Output

Write-Output 'Step 3/4: Internet ping test'
ping.exe 8.8.8.8 -n 3 | Out-String | Write-Output

Write-Output 'Step 4/4: Route table snapshot'
route.exe print | Out-String | Write-Output

Write-Output 'Network maintenance diagnostics complete.'
