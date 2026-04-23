[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Write-Output 'Running DISM component cleanup to free C: space...'
dism.exe /Online /Cleanup-Image /StartComponentCleanup | Out-String | Write-Output

Write-Output 'Running Storage Sense trigger (if available)...'
try {
    Start-Process -FilePath cleanmgr.exe -ArgumentList '/VERYLOWDISK' -Wait -NoNewWindow
    Write-Output 'Disk cleanup completed.'
}
catch {
    Write-Output "cleanmgr did not run: $($_.Exception.Message)"
}

Write-Output 'C: drive free-space task completed.'
