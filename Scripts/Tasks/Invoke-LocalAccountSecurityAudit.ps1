[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-AuditLine {
    param(
        [Parameter(Mandatory = $true)][string]$Status,
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$Detail,
        [string]$Recommendation = ''
    )

    Write-Output ("[{0}] {1}: {2}" -f $Status, $Name, $Detail)
    if ($Recommendation) {
        Write-Output ("    Recommendation: {0}" -f $Recommendation)
    }
}

Write-Output 'Starting local account and password policy audit.'
Write-Output 'Defensive objective: find weak local account settings that increase unauthorized access risk.'

Write-Output 'Checking local password policy with net accounts...'
$netAccounts = @(cmd.exe /c 'net accounts' 2>&1)
foreach ($line in $netAccounts) {
    if (-not [string]::IsNullOrWhiteSpace($line)) {
        Write-Output ("Policy: {0}" -f $line.Trim())
    }
}

$lockoutLine = $netAccounts | Where-Object { $_ -match 'Lockout threshold' } | Select-Object -First 1
if ($lockoutLine -and $lockoutLine -match 'Never') {
    Write-AuditLine -Status 'WARN' -Name 'Account lockout policy' -Detail 'Lockout threshold appears to be Never.' -Recommendation 'Consider an account lockout threshold to slow password guessing on local accounts.'
}

$minPasswordLine = $netAccounts | Where-Object { $_ -match 'Minimum password length' } | Select-Object -First 1
if ($minPasswordLine -and $minPasswordLine -match '(\d+)') {
    $minLength = [int]$Matches[1]
    Write-AuditLine -Status $(if ($minLength -ge 12) { 'OK' } else { 'WARN' }) -Name 'Minimum password length' -Detail "Minimum length=$minLength" -Recommendation $(if ($minLength -lt 12) { 'Use 12+ characters where possible, or follow your organization password policy.' } else { '' })
}

if (Get-Command Get-LocalUser -ErrorAction SilentlyContinue) {
    Write-Output 'Checking local user accounts...'
    foreach ($user in (Get-LocalUser | Sort-Object Name)) {
        $status = 'INFO'
        $recommendation = ''
        if ($user.Enabled -and $user.Name -match 'guest') {
            $status = 'WARN'
            $recommendation = 'Disable Guest-style accounts unless there is a documented business need.'
        }
        elseif ($user.Enabled -and $user.PasswordRequired -eq $false) {
            $status = 'RISK'
            $recommendation = 'Require a password for enabled local accounts.'
        }
        elseif ($user.Enabled) {
            $status = 'OK'
        }

        Write-AuditLine -Status $status -Name "Local user $($user.Name)" -Detail "Enabled=$($user.Enabled), PasswordRequired=$($user.PasswordRequired), LastLogon=$($user.LastLogon)" -Recommendation $recommendation
    }
}
else {
    Write-AuditLine -Status 'WARN' -Name 'Local user audit' -Detail 'Get-LocalUser is unavailable on this system.'
}

if (Get-Command Get-LocalGroupMember -ErrorAction SilentlyContinue) {
    Write-Output 'Checking local Administrators group membership...'
    $admins = @(Get-LocalGroupMember -Group 'Administrators' -ErrorAction Stop)
    foreach ($admin in $admins) {
        Write-AuditLine -Status 'REVIEW' -Name 'Administrator member' -Detail "$($admin.Name) [$($admin.ObjectClass)]"
    }
    if ($admins.Count -gt 3) {
        Write-AuditLine -Status 'WARN' -Name 'Administrator group size' -Detail "$($admins.Count) member(s) found." -Recommendation 'Review whether each account still requires local administrator rights.'
    }
}
else {
    Write-AuditLine -Status 'WARN' -Name 'Administrator group audit' -Detail 'Get-LocalGroupMember is unavailable on this system.'
}

Write-Output 'Local account and password policy audit complete.'
