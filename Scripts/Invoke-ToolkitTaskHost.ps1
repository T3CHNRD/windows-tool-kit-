[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$ToolkitRoot,
    [Parameter(Mandatory = $true)][string]$TaskId,
    [string]$SelectionFile
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-HostToken {
    param(
        [Parameter(Mandatory = $true)][string]$Kind,
        [string[]]$Parts = @()
    )

    $segments = @($Kind)
    foreach ($part in $Parts) {
        $segments += ([string]$part -replace '\|', '/')
    }

    Write-Output ('__TTK__|' + ($segments -join '|'))
}

$modulePath = Join-Path $ToolkitRoot 'Modules\MaintenanceToolkit\MaintenanceToolkit.psm1'
Import-Module $modulePath -Force

$context = New-ToolkitTaskContext `
    -ReportProgress {
        param($Percent, $Status)
        Write-HostToken -Kind 'PROGRESS' -Parts @([string][int]$Percent, [string]$Status)
    } `
    -WriteStatus {
        param($Status)
        Write-HostToken -Kind 'STATUS' -Parts @([string]$Status)
    } `
    -WriteLog {
        param($Message, $Level)
        $line = Write-ToolkitLog -Message $Message -Level $Level
        Write-HostToken -Kind 'LOG' -Parts @([string]$line)
    }

$context | Add-Member -NotePropertyName UserInput -NotePropertyValue @{} -Force
if ($SelectionFile) {
    $context.UserInput.SelectionFile = $SelectionFile
}

$task = Get-ToolkitTaskCatalog | Where-Object { $_.Id -eq $TaskId } | Select-Object -First 1
if (-not $task) {
    throw "Task '$TaskId' was not found."
}

try {
    & $context.WriteLog "Starting task: $($task.Name)" 'INFO'
    Invoke-ToolkitTaskById -TaskId $TaskId -Context $context
    & $context.WriteLog "Task completed successfully: $($task.Name)" 'INFO'
    Write-HostToken -Kind 'RESULT' -Parts @('SUCCESS', $task.Name)
    exit 0
}
catch {
    $message = $_.Exception.Message
    & $context.WriteLog "Task failed: $message" 'ERROR'
    Write-HostToken -Kind 'RESULT' -Parts @('FAIL', $task.Name, $message)
    exit 1
}
