[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Write-Output 'Collecting disk usage details...'
$drives = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3"

foreach ($drive in $drives) {
    $freeGb = [math]::Round($drive.FreeSpace / 1GB, 2)
    $sizeGb = [math]::Round($drive.Size / 1GB, 2)
    $pctFree = if ($drive.Size -gt 0) { [math]::Round(($drive.FreeSpace / $drive.Size) * 100, 1) } else { 0 }
    $severity = if ($pctFree -lt 15) { 'LOW SPACE' } else { 'OK' }
    Write-Output ("Drive {0}: {1} GB free / {2} GB total ({3}% free) [{4}]" -f $drive.DeviceID, $freeGb, $sizeGb, $pctFree, $severity)
}

Write-Output 'Disk space monitoring complete.'
