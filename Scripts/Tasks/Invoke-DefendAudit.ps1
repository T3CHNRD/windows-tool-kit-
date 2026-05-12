Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'
Write-Output 'DEFEND-inspired security hardening audit. This tool reports status and does not make changes.'
$mp = Get-MpComputerStatus -ErrorAction SilentlyContinue
if ($mp) {
    Write-Output "Defender antivirus enabled: $($mp.AntivirusEnabled)"
    Write-Output "Real-time protection enabled: $($mp.RealTimeProtectionEnabled)"
    Write-Output "Behavior monitor enabled: $($mp.BehaviorMonitorEnabled)"
    Write-Output "Antispyware signatures age: $($mp.AntispywareSignatureAge)"
} else { Write-Output 'Defender status unavailable, possibly due to third-party AV or unsupported edition.' }
Write-Output 'Firewall profiles:'
Get-NetFirewallProfile -ErrorAction SilentlyContinue | Select Name,Enabled,DefaultInboundAction,DefaultOutboundAction | Format-Table -AutoSize | Out-String | Write-Output
Write-Output 'Controlled folder access:'
Get-MpPreference -ErrorAction SilentlyContinue | Select EnableControlledFolderAccess,PUAProtection,DisableRealtimeMonitoring | Format-List | Out-String | Write-Output
Write-Output 'Completed: Defender/firewall baseline checked. Skipped: no changes applied. Failed: see warnings above if any.'
exit 0
