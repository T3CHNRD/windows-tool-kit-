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
    Write-Step 'Starting BitLocker recovery key backup workflow.'
    if (-not (Test-IsAdministrator)) {
        throw 'Administrator privileges are required to read and back up BitLocker recovery protectors.'
    }

    if (-not (Get-Command Get-BitLockerVolume -ErrorAction SilentlyContinue)) {
        throw 'BitLocker PowerShell cmdlets are not available on this Windows edition.'
    }

    $removableTargets = @(Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=2' -ErrorAction SilentlyContinue |
        Where-Object { $_.DeviceID -and (Test-Path ($_.DeviceID + '\')) })

    if (-not $removableTargets) {
        throw 'No writable removable/external media was detected. Insert a USB drive or external disk and run this tool again.'
    }

    $volumes = @(Get-BitLockerVolume -ErrorAction Stop | Where-Object { $_.MountPoint })
    if (-not $volumes) {
        Write-Step 'No BitLocker-capable volumes were returned by Windows.' 'WARN'
        exit 0
    }

    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $computer = $env:COMPUTERNAME
    $reportLines = New-Object System.Collections.Generic.List[string]
    $reportLines.Add('T3CHNRD BitLocker Recovery Key Backup')
    $reportLines.Add("Computer: $computer")
    $reportLines.Add("Created: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    $reportLines.Add('Keep this file offline and protected. Anyone with recovery keys can unlock protected drives.')
    $reportLines.Add('')

    $completed = 0
    $skipped = 0
    $failed = 0

    foreach ($volume in $volumes) {
        try {
            $mount = $volume.MountPoint
            Write-Step "Reviewing BitLocker protectors for $mount."
            $protectors = @($volume.KeyProtector | Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' })

            if (-not $protectors -and $volume.ProtectionStatus -ne 'Off') {
                Write-Step "No recovery-password protector found for $mount. Creating one before backup." 'WARN'
                Add-BitLockerKeyProtector -MountPoint $mount -RecoveryPasswordProtector -ErrorAction Stop | Out-Null
                $volume = Get-BitLockerVolume -MountPoint $mount -ErrorAction Stop
                $protectors = @($volume.KeyProtector | Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' })
            }

            if (-not $protectors) {
                $skipped++
                $reportLines.Add("Volume: $mount")
                $reportLines.Add("Status: skipped - no recovery-password protector found and protection is $($volume.ProtectionStatus)")
                $reportLines.Add('')
                Write-Step "Skipped $mount because no recovery-password protector is available." 'WARN'
                continue
            }

            $reportLines.Add("Volume: $mount")
            $reportLines.Add("Volume Status: $($volume.VolumeStatus)")
            $reportLines.Add("Protection Status: $($volume.ProtectionStatus)")
            foreach ($protector in $protectors) {
                $reportLines.Add("Key Protector ID: $($protector.KeyProtectorId)")
                $reportLines.Add("Recovery Password: $($protector.RecoveryPassword)")
            }
            $reportLines.Add('')
            $completed++
            Write-Step "Captured recovery protector metadata for $mount. Recovery password is written only to the external backup file."
        } catch {
            $failed++
            $reportLines.Add("Volume: $($volume.MountPoint)")
            $reportLines.Add("Status: failed - $($_.Exception.Message)")
            $reportLines.Add('')
            Write-Step "Failed to process $($volume.MountPoint): $($_.Exception.Message)" 'ERROR'
        }
    }

    $writtenFiles = New-Object System.Collections.Generic.List[string]
    foreach ($target in $removableTargets) {
        try {
            $folder = Join-Path ($target.DeviceID + '\') 'T3CHNRD-BitLocker-Backup'
            New-Item -Path $folder -ItemType Directory -Force | Out-Null
            $file = Join-Path $folder "BitLocker-RecoveryKeys-$computer-$timestamp.txt"
            Set-Content -LiteralPath $file -Value $reportLines -Encoding UTF8 -Force
            $writtenFiles.Add($file)
            Write-Step "Backup file written to external media: $file"
        } catch {
            $failed++
            Write-Step "Failed to write backup to $($target.DeviceID): $($_.Exception.Message)" 'ERROR'
        }
    }

    Write-Step "BitLocker key backup summary: completed=$completed skipped=$skipped failed=$failed external_files=$($writtenFiles.Count)."
    if ($completed -gt 0 -and $writtenFiles.Count -eq 0) {
        throw 'Recovery keys were collected, but no external backup file could be written.'
    }
    if ($failed -gt 0) { exit 1 }
    exit 0
} catch {
    Write-Step $_.Exception.Message 'ERROR'
    exit 1
}
