[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Write-Output 'Resetting Winsock...'
netsh.exe winsock reset | Out-String | Write-Output

Write-Output 'Resetting TCP/IP stack...'
netsh.exe int ip reset | Out-String | Write-Output

Write-Output 'Flushing DNS cache...'
ipconfig.exe /flushdns | Out-String | Write-Output

Write-Output 'Network stack reset complete. Reboot recommended.'
