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

function Invoke-ChkdskScan {
    param([Parameter(Mandatory = $true)][string]$Drive)
    Write-Step "Running chkdsk online scan for $Drive."
    & chkdsk.exe $Drive /scan 2>&1 | ForEach-Object { Write-Step ([string]$_) }
    return $LASTEXITCODE
}

try {
    Write-Step 'Starting drive scan and repair workflow.'
    if (-not (Test-IsAdministrator)) {
        throw 'Administrator privileges are required for drive repair operations.'
    }

    Write-Step 'Collecting disk and partition inventory.'
    Get-Disk -ErrorAction SilentlyContinue |
        Sort-Object Number |
        ForEach-Object { Write-Step ("Disk {0}: {1}, {2}, {3}, size {4:N1} GB" -f $_.Number, $_.FriendlyName, $_.PartitionStyle, $_.HealthStatus, ($_.Size / 1GB)) }

    $volumes = @(Get-Volume -ErrorAction Stop |
        Where-Object { $_.DriveLetter -and $_.DriveType -in @('Fixed', 'Removable') } |
        Sort-Object DriveLetter)

    if (-not $volumes) {
        Write-Step 'No internal or external drive-letter volumes were found to scan.' 'WARN'
        exit 0
    }

    $systemDrive = $env:SystemDrive.TrimEnd(':')
    $completed = 0
    $skipped = 0
    $failed = 0

    foreach ($volume in $volumes) {
        $letter = [string]$volume.DriveLetter
        $drive = "$letter`:"
        try {
            Write-Step "Scanning $drive ($($volume.FileSystemLabel)) type=$($volume.DriveType) filesystem=$($volume.FileSystem) health=$($volume.HealthStatus)."

            if ($volume.FileSystem -notin @('NTFS', 'ReFS', 'FAT32', 'exFAT')) {
                $skipped++
                Write-Step "Skipping $drive because filesystem '$($volume.FileSystem)' is not supported by this Windows repair workflow." 'WARN'
                continue
            }

            if (Get-Command Repair-Volume -ErrorAction SilentlyContinue) {
                if ($volume.FileSystem -in @('NTFS', 'ReFS')) {
                    Write-Step "Running Repair-Volume -Scan for $drive."
                    $scanResult = Repair-Volume -DriveLetter $letter -Scan -ErrorAction SilentlyContinue
                    if ($scanResult) { Write-Step "Repair-Volume scan result for ${drive}: $scanResult" }
                }
            }

            $exitCode = Invoke-ChkdskScan -Drive $drive
            Write-Step "chkdsk /scan exit code for ${drive}: $exitCode."

            if ($letter -ieq $systemDrive) {
                Write-Step "$drive is the Windows system drive. Online scan completed; offline fixes are not forced from the toolkit." 'WARN'
            } elseif ($volume.DriveType -eq 'Removable') {
                Write-Step "Attempting non-system removable drive repair with chkdsk $drive /f. Close files on this drive if Windows asks to lock it." 'WARN'
                & chkdsk.exe $drive /f 2>&1 | ForEach-Object { Write-Step ([string]$_) }
                Write-Step "chkdsk /f exit code for ${drive}: $LASTEXITCODE."
            } else {
                Write-Step "$drive is a non-system internal drive. Scan completed; run again after closing open files if repairs are needed." 
            }

            $completed++
        } catch {
            $failed++
            Write-Step "Failed to scan/repair ${drive}: $($_.Exception.Message)" 'ERROR'
        }
    }

    Write-Step "Drive scan and repair summary: completed=$completed skipped=$skipped failed=$failed."
    if ($failed -gt 0) { exit 1 }
    exit 0
} catch {
    Write-Step $_.Exception.Message 'ERROR'
    exit 1
}

