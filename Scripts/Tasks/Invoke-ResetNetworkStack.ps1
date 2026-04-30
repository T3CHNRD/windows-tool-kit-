[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Test-IsAdmin {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-IsAdmin)) {
    throw 'Network stack reset requires administrator rights.'
}

Write-Output 'Resetting Winsock...'
netsh.exe winsock reset | Out-String | Write-Output

Write-Output 'Resetting TCP/IP stack...'
netsh.exe int ip reset | Out-String | Write-Output

Write-Output 'Resetting IPv6 stack...'
netsh.exe int ipv6 reset | Out-String | Write-Output

Write-Output 'Resetting Windows Firewall policy...'
netsh.exe advfirewall reset | Out-String | Write-Output

Write-Output 'Flushing DNS cache...'
ipconfig.exe /flushdns | Out-String | Write-Output

Write-Output 'Waiting 3 seconds for network services and adapters to stabilize...'
Start-Sleep -Seconds 3

Write-Output 'Network stack reset complete. Reboot recommended.'
