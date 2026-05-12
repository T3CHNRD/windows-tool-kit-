#Requires -Version 5.1
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Step {
    param([string]$Message, [string]$Level = 'INFO')
    Write-Output "[$Level] $Message"
}

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

try {
    Write-Step 'Starting BitLocker disable/decryption workflow.'
    if (-not (Test-IsAdministrator)) {
        throw 'Administrator privileges are required to disable BitLocker.'
    }

    if (-not (Get-Command Get-BitLockerVolume -ErrorAction SilentlyContinue)) {
        throw 'BitLocker PowerShell cmdlets are not available on this Windows edition.'
    }

    $volumes = @(Get-BitLockerVolume -ErrorAction Stop | Where-Object { $_.MountPoint })
    if (-not $volumes) {
        Write-Step 'No BitLocker-capable volumes were returned by Windows.' 'WARN'
        exit 0
    }

    $completed = 0
    $skipped = 0
    $failed = 0

    foreach ($volume in $volumes) {
        try {
            $mount = $volume.MountPoint
            Write-Step "Reviewing BitLocker state for ${mount}: protection=$($volume.ProtectionStatus), volume=$($volume.VolumeStatus)."

            if ($volume.ProtectionStatus -eq 'Off' -and $volume.VolumeStatus -eq 'FullyDecrypted') {
                $skipped++
                Write-Step "$mount is already fully decrypted. Skipping."
                continue
            }

            Write-Step "Disabling BitLocker on $mount. Decryption may continue in the background after this task finishes." 'WARN'
            Disable-BitLocker -MountPoint $mount -ErrorAction Stop
            Start-Sleep -Seconds 2
            $updated = Get-BitLockerVolume -MountPoint $mount -ErrorAction Stop
            Write-Step "$mount updated state: protection=$($updated.ProtectionStatus), volume=$($updated.VolumeStatus), encryption=$($updated.EncryptionPercentage) percent."
            $completed++
        } catch {
            $failed++
            Write-Step "Failed to disable BitLocker on $($volume.MountPoint): $($_.Exception.Message)" 'ERROR'
        }
    }

    Write-Step "BitLocker disable summary: completed=$completed skipped=$skipped failed=$failed."
    if ($failed -gt 0) { exit 1 }
    exit 0
} catch {
    Write-Step $_.Exception.Message 'ERROR'
    exit 1
}

