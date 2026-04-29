[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Finding {
    param(
        [Parameter(Mandatory = $true)][string]$Area,
        [Parameter(Mandatory = $true)][string]$Detail,
        [string]$Status = 'REVIEW'
    )

    Write-Output ("[{0}] {1}: {2}" -f $Status, $Area, $Detail)
}

Write-Output 'Starting startup items review.'
Write-Output 'Review objective: list common autorun locations and highlight unusual auto-start entries.'

$runKeys = @(
    'HKLM:\Software\Microsoft\Windows\CurrentVersion\Run',
    'HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce',
    'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run',
    'HKCU:\Software\Microsoft\Windows\CurrentVersion\RunOnce',
    'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Run'
)

foreach ($keyPath in $runKeys) {
    Write-Output "Checking registry autorun key: $keyPath"
    if (-not (Test-Path -LiteralPath $keyPath)) {
        Write-Finding -Area 'Registry autorun' -Detail "$keyPath not present." -Status 'INFO'
        continue
    }

    $item = Get-ItemProperty -LiteralPath $keyPath -ErrorAction SilentlyContinue
    foreach ($property in $item.PSObject.Properties) {
        if ($property.Name -in @('PSPath', 'PSParentPath', 'PSChildName', 'PSDrive', 'PSProvider')) {
            continue
        }
        Write-Finding -Area 'Registry autorun' -Detail "$keyPath -> $($property.Name) = $($property.Value)"
    }
}

$startupFolders = @(
    [Environment]::GetFolderPath('Startup'),
    (Join-Path $env:ProgramData 'Microsoft\Windows\Start Menu\Programs\Startup')
) | Where-Object { $_ -and (Test-Path -LiteralPath $_) }

foreach ($folder in $startupFolders) {
    Write-Output "Checking startup folder: $folder"
    $items = @(Get-ChildItem -LiteralPath $folder -Force -ErrorAction SilentlyContinue)
    if ($items.Count -eq 0) {
        Write-Finding -Area 'Startup folder' -Detail "$folder is empty." -Status 'OK'
    }
    foreach ($item in $items) {
        Write-Finding -Area 'Startup folder' -Detail "$($item.FullName)"
    }
}

Write-Output 'Checking scheduled tasks that are enabled and not Microsoft-authored where possible...'
$tasks = @(Get-ScheduledTask -ErrorAction SilentlyContinue |
    Where-Object { $_.State -ne 'Disabled' } |
    Sort-Object TaskPath, TaskName)

foreach ($task in ($tasks | Select-Object -First 80)) {
    $taskInfo = $null
    try { $taskInfo = Get-ScheduledTaskInfo -TaskName $task.TaskName -TaskPath $task.TaskPath -ErrorAction SilentlyContinue } catch {}
    $actionText = ($task.Actions | ForEach-Object { "$($_.Execute) $($_.Arguments)" }) -join ' | '
    $status = if ($task.TaskPath -like '\Microsoft\*') { 'INFO' } else { 'REVIEW' }
    Write-Finding -Area 'Scheduled task' -Status $status -Detail "$($task.TaskPath)$($task.TaskName); LastRun=$($taskInfo.LastRunTime); Action=$actionText"
}

Write-Output 'Checking automatic services outside Windows directory...'
$services = @(Get-CimInstance Win32_Service -Filter "StartMode='Auto'" -ErrorAction SilentlyContinue |
    Sort-Object Name)
foreach ($service in $services) {
    $pathName = [string]$service.PathName
    $status = 'INFO'
    if ($pathName -and $pathName -notmatch '^[`"]?C:\\Windows\\' -and $pathName -notmatch '^[`"]?%SystemRoot%\\') {
        $status = 'REVIEW'
    }
    Write-Finding -Area 'Auto service' -Status $status -Detail "$($service.Name) [$($service.State)] -> $pathName"
}

Write-Output 'Startup items review complete.'
