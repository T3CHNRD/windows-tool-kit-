Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'
Write-Output 'ORCA-inspired security report: combines endpoint posture, open ports, and risky local configuration indicators.'
Write-Output '--- Defender ---'
Get-MpComputerStatus -ErrorAction SilentlyContinue | Select AntivirusEnabled,RealTimeProtectionEnabled,BehaviorMonitorEnabled,IoavProtectionEnabled,NISEnabled | Format-List | Out-String | Write-Output
Write-Output '--- Firewall ---'
Get-NetFirewallProfile -ErrorAction SilentlyContinue | Select Name,Enabled,DefaultInboundAction | Format-Table -AutoSize | Out-String | Write-Output
Write-Output '--- Listening Ports ---'
Get-NetTCPConnection -State Listen -ErrorAction SilentlyContinue | Select LocalAddress,LocalPort,OwningProcess,@{n='Process';e={(Get-Process -Id $_.OwningProcess -ErrorAction SilentlyContinue).Name}} | Sort LocalPort | Format-Table -AutoSize | Out-String | Write-Output
Write-Output '--- SMBv1 Optional Feature ---'
Get-WindowsOptionalFeature -Online -FeatureName SMB1Protocol -ErrorAction SilentlyContinue | Select FeatureName,State | Format-List | Out-String | Write-Output
Write-Output 'Completed: ORCA-style report. Skipped: remediation. Failed: see unavailable sections above.'
exit 0
