[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Finding {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$Status,
        [Parameter(Mandatory = $true)][string]$Detail,
        [string]$Recommendation = ''
    )

    Write-Output ("[{0}] {1}: {2}" -f $Status, $Name, $Detail)
    if ($Recommendation) {
        Write-Output ("    Recommendation: {0}" -f $Recommendation)
    }
}

function Invoke-CheckedAudit {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][scriptblock]$Action
    )

    Write-Output "Checking: $Name"
    try {
        & $Action
    }
    catch {
        Write-Finding -Name $Name -Status 'WARN' -Detail "Could not complete check: $($_.Exception.Message)"
    }
}

Write-Output 'Starting Windows security baseline audit.'
Write-Output 'Scope: local defensive checks only. This does not run offensive actions or scan third-party systems.'

Invoke-CheckedAudit -Name 'Microsoft Defender status' -Action {
    if (-not (Get-Command Get-MpComputerStatus -ErrorAction SilentlyContinue)) {
        Write-Finding -Name 'Microsoft Defender status' -Status 'WARN' -Detail 'Defender PowerShell cmdlets are not available on this system.' -Recommendation 'Confirm Microsoft Defender Antivirus is installed or managed by another endpoint security product.'
        return
    }

    $status = Get-MpComputerStatus
    Write-Finding -Name 'Defender real-time protection' -Status $(if ($status.RealTimeProtectionEnabled) { 'OK' } else { 'RISK' }) -Detail "RealTimeProtectionEnabled=$($status.RealTimeProtectionEnabled)" -Recommendation $(if (-not $status.RealTimeProtectionEnabled) { 'Enable real-time protection or verify your managed EDR policy.' } else { '' })
    Write-Finding -Name 'Defender behavior monitor' -Status $(if ($status.BehaviorMonitorEnabled) { 'OK' } else { 'RISK' }) -Detail "BehaviorMonitorEnabled=$($status.BehaviorMonitorEnabled)"
    Write-Finding -Name 'Defender signatures' -Status 'INFO' -Detail "AntivirusSignatureAge=$($status.AntivirusSignatureAge), AntispywareSignatureAge=$($status.AntispywareSignatureAge), QuickScanAge=$($status.QuickScanAge)"
    if ($status.PSObject.Properties.Name -contains 'IsTamperProtected') {
        Write-Finding -Name 'Defender tamper protection' -Status $(if ($status.IsTamperProtected) { 'OK' } else { 'WARN' }) -Detail "IsTamperProtected=$($status.IsTamperProtected)"
    }
}

Invoke-CheckedAudit -Name 'Windows Firewall profiles' -Action {
    if (-not (Get-Command Get-NetFirewallProfile -ErrorAction SilentlyContinue)) {
        Write-Finding -Name 'Windows Firewall profiles' -Status 'WARN' -Detail 'Get-NetFirewallProfile is unavailable.'
        return
    }

    foreach ($profile in Get-NetFirewallProfile) {
        $state = if ($profile.Enabled) { 'OK' } else { 'RISK' }
        Write-Finding -Name "Firewall profile $($profile.Name)" -Status $state -Detail "Enabled=$($profile.Enabled), DefaultInboundAction=$($profile.DefaultInboundAction), DefaultOutboundAction=$($profile.DefaultOutboundAction)" -Recommendation $(if (-not $profile.Enabled) { 'Turn on this firewall profile unless another managed firewall is intentionally replacing it.' } else { '' })
    }
}

Invoke-CheckedAudit -Name 'BitLocker protection' -Action {
    if (-not (Get-Command Get-BitLockerVolume -ErrorAction SilentlyContinue)) {
        Write-Finding -Name 'BitLocker protection' -Status 'WARN' -Detail 'BitLocker cmdlets are unavailable on this edition or session.'
        return
    }

    foreach ($volume in Get-BitLockerVolume -ErrorAction Stop) {
        $status = if ($volume.ProtectionStatus -eq 'On') { 'OK' } else { 'RISK' }
        Write-Finding -Name "BitLocker $($volume.MountPoint)" -Status $status -Detail "ProtectionStatus=$($volume.ProtectionStatus), VolumeStatus=$($volume.VolumeStatus)" -Recommendation $(if ($volume.ProtectionStatus -ne 'On') { 'Enable BitLocker on fixed drives that contain sensitive data.' } else { '' })
    }
}

Invoke-CheckedAudit -Name 'Secure Boot' -Action {
    if (-not (Get-Command Confirm-SecureBootUEFI -ErrorAction SilentlyContinue)) {
        Write-Finding -Name 'Secure Boot' -Status 'WARN' -Detail 'Confirm-SecureBootUEFI is unavailable or this device is not booted in UEFI mode.'
        return
    }

    $enabled = Confirm-SecureBootUEFI
    Write-Finding -Name 'Secure Boot' -Status $(if ($enabled) { 'OK' } else { 'RISK' }) -Detail "Enabled=$enabled" -Recommendation $(if (-not $enabled) { 'Enable Secure Boot in firmware settings when supported.' } else { '' })
}

Invoke-CheckedAudit -Name 'SMBv1 protocol' -Action {
    $feature = Get-WindowsOptionalFeature -Online -FeatureName SMB1Protocol -ErrorAction SilentlyContinue
    if ($feature) {
        Write-Finding -Name 'SMBv1 protocol' -Status $(if ($feature.State -eq 'Disabled') { 'OK' } else { 'RISK' }) -Detail "State=$($feature.State)" -Recommendation $(if ($feature.State -ne 'Disabled') { 'Disable SMBv1 unless a documented legacy dependency exists.' } else { '' })
    }
    else {
        Write-Finding -Name 'SMBv1 protocol' -Status 'INFO' -Detail 'SMB1Protocol optional feature was not found.'
    }
}

Invoke-CheckedAudit -Name 'Remote Desktop exposure' -Action {
    $rdpValue = Get-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Terminal Server' -Name fDenyTSConnections -ErrorAction Stop
    $rdpEnabled = ($rdpValue.fDenyTSConnections -eq 0)
    Write-Finding -Name 'Remote Desktop' -Status $(if ($rdpEnabled) { 'WARN' } else { 'OK' }) -Detail "Enabled=$rdpEnabled" -Recommendation $(if ($rdpEnabled) { 'Keep RDP disabled unless needed, restrict it with firewall/VPN/MFA, and monitor logons.' } else { '' })
}

Invoke-CheckedAudit -Name 'Local administrators' -Action {
    if (-not (Get-Command Get-LocalGroupMember -ErrorAction SilentlyContinue)) {
        Write-Finding -Name 'Local administrators' -Status 'WARN' -Detail 'Get-LocalGroupMember is unavailable.'
        return
    }

    $members = @(Get-LocalGroupMember -Group 'Administrators' -ErrorAction Stop)
    Write-Finding -Name 'Local administrators' -Status $(if ($members.Count -le 3) { 'OK' } else { 'WARN' }) -Detail ("{0} member(s): {1}" -f $members.Count, (($members | Select-Object -ExpandProperty Name) -join ', ')) -Recommendation $(if ($members.Count -gt 3) { 'Review whether every local administrator still needs admin rights.' } else { '' })
}

Invoke-CheckedAudit -Name 'Listening ports summary' -Action {
    if (-not (Get-Command Get-NetTCPConnection -ErrorAction SilentlyContinue)) {
        Write-Finding -Name 'Listening ports summary' -Status 'WARN' -Detail 'Get-NetTCPConnection is unavailable.'
        return
    }

    $listeners = @(Get-NetTCPConnection -State Listen -ErrorAction SilentlyContinue)
    $externalListeners = @($listeners | Where-Object { $_.LocalAddress -notin @('127.0.0.1', '::1') })
    Write-Finding -Name 'Listening ports summary' -Status $(if ($externalListeners.Count -le 10) { 'OK' } else { 'WARN' }) -Detail ("{0} non-loopback listener(s) detected." -f $externalListeners.Count) -Recommendation $(if ($externalListeners.Count -gt 10) { 'Run the Open Ports and Services Audit tool and close unnecessary listeners.' } else { '' })
}

Write-Output 'Windows security baseline audit complete.'
