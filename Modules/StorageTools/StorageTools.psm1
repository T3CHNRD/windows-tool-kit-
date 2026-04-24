Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-TtkDriveReport {
    [CmdletBinding()]
    param()

    Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3" |
        ForEach-Object {
            $freeGb = [math]::Round($_.FreeSpace / 1GB, 2)
            $sizeGb = [math]::Round($_.Size / 1GB, 2)
            $pctFree = if ($_.Size -gt 0) { [math]::Round(($_.FreeSpace / $_.Size) * 100, 1) } else { 0 }
            [pscustomobject]@{
                Drive = $_.DeviceID
                VolumeName = $_.VolumeName
                FileSystem = $_.FileSystem
                FreeGB = $freeGb
                SizeGB = $sizeGb
                PercentFree = $pctFree
                Severity = if ($pctFree -lt 15) { 'LOW SPACE' } else { 'OK' }
            }
        }
}

function Start-TtkFileTransfer {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Source,
        [Parameter(Mandatory = $true)][string]$Destination,
        [switch]$Move,
        [string]$LogPath
    )

    if (-not (Test-Path -LiteralPath $Source)) {
        throw "Source path not found: $Source"
    }

    $null = New-Item -ItemType Directory -Force -Path $Destination
    if (-not $LogPath) {
        $LogPath = Join-Path $env:TEMP ("robocopy-{0:yyyyMMdd-HHmmss}.log" -f (Get-Date))
    }

    $copyFlags = @('/E','/R:2','/W:2','/COPY:DAT','/DCOPY:DAT','/MT:8','/TEE','/LOG+:' + $LogPath)
    if ($Move) {
        $copyFlags += '/MOVE'
    }

    $arguments = @("`"$Source`"","`"$Destination`"") + $copyFlags
    $process = Start-Process -FilePath 'robocopy.exe' -ArgumentList $arguments -Wait -PassThru -NoNewWindow
    if ($process.ExitCode -ge 8) {
        throw "Robocopy failed with exit code $($process.ExitCode). See $LogPath"
    }

    return [pscustomobject]@{
        ExitCode = $process.ExitCode
        LogPath = $LogPath
        MoveMode = [bool]$Move
    }
}

function Get-TtkDiskInventory {
    [CmdletBinding()]
    param()

    Get-Disk | Select-Object Number, FriendlyName, PartitionStyle, Size, OperationalStatus, BusType
}

function Invoke-TtkCloneDiskGuide {
    [CmdletBinding()]
    param()

    return @(
        'Disk cloning is orchestrated as a guided workflow.',
        '1. Back up important data before continuing.',
        '2. Connect the destination disk and confirm it is the correct target.',
        '3. Use manufacturer or imaging software for the actual sector copy when required.',
        '4. Validate boot order and disk health after the clone completes.'
    )
}

function Invoke-TtkNewComputerSetupChecklist {
    [CmdletBinding()]
    param()

    return @(
        'Apply Windows updates',
        'Install manufacturer updates',
        'Install Microsoft 365 and required apps',
        'Set power and BitLocker policies',
        'Transfer user data',
        'Validate backup and restore points'
    )
}

Export-ModuleMember -Function @(
    'Get-TtkDriveReport',
    'Start-TtkFileTransfer',
    'Get-TtkDiskInventory',
    'Invoke-TtkCloneDiskGuide',
    'Invoke-TtkNewComputerSetupChecklist'
)
