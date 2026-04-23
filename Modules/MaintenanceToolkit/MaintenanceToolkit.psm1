Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-ToolkitRoot {
    $modulePath = $PSScriptRoot
    return (Resolve-Path (Join-Path $modulePath '..\..')).Path
}

function Get-ToolkitSettings {
    $settingsPath = Join-Path (Get-ToolkitRoot) 'Config\Toolkit.Settings.psd1'
    if (Test-Path $settingsPath) {
        return Import-PowerShellDataFile -Path $settingsPath
    }

    return @{
        LogRoot = 'Logs'
        Integrations = @{
            Microsoft365RepoZip = 'https://github.com/mallockey/Install-Microsoft365/archive/refs/heads/main.zip'
            WindowsMediaSupportPage = 'https://support.microsoft.com/en-us/windows/create-installation-media-for-windows-99a58364-8c02-206f-aa6f-40c3b507420d'
            Windows10MediaToolUrl = 'https://go.microsoft.com/fwlink/?LinkId=691209'
        }
    }
}

function Test-ToolkitIsAdmin {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Write-ToolkitLog {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level = 'INFO'
    )

    $settings = Get-ToolkitSettings
    $root = Get-ToolkitRoot
    $logDir = Join-Path $root $settings.LogRoot
    if (-not (Test-Path $logDir)) {
        New-Item -Path $logDir -ItemType Directory -Force | Out-Null
    }

    $logPath = Join-Path $logDir ("Toolkit-{0:yyyyMMdd}.log" -f (Get-Date))
    $line = "[{0:yyyy-MM-dd HH:mm:ss}] [{1}] {2}" -f (Get-Date), $Level, $Message
    Add-Content -Path $logPath -Value $line
    return $line
}

function New-ToolkitTaskContext {
    param(
        [Parameter(Mandatory = $true)][scriptblock]$ReportProgress,
        [Parameter(Mandatory = $true)][scriptblock]$WriteStatus,
        [Parameter(Mandatory = $true)][scriptblock]$WriteLog
    )

    $root = Get-ToolkitRoot
    $settings = Get-ToolkitSettings

    return [pscustomobject]@{
        ToolkitRoot = $root
        LegacyScriptsPath = Join-Path $root 'LegacyScripts'
        IntegrationsPath = Join-Path $root 'Integrations'
        Settings = $settings
        ReportProgress = $ReportProgress
        WriteStatus = $WriteStatus
        WriteLog = $WriteLog
    }
}

function Invoke-ToolkitTaskById {
    param(
        [Parameter(Mandatory = $true)][string]$TaskId,
        [Parameter(Mandatory = $true)]$Context
    )

    $task = Get-ToolkitTaskCatalog | Where-Object { $_.Id -eq $TaskId } | Select-Object -First 1
    if (-not $task) {
        throw "Task '$TaskId' was not found."
    }

    if ($task.RequiresAdmin -and -not (Test-ToolkitIsAdmin)) {
        throw "Task '$($task.Name)' requires an elevated PowerShell session."
    }

    & $task.Handler $Context $task
}

function Invoke-TaskStep {
    param(
        [Parameter(Mandatory = $true)]$Context,
        [Parameter(Mandatory = $true)][int]$Percent,
        [Parameter(Mandatory = $true)][string]$Status,
        [string]$LogMessage,
        [ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level = 'INFO'
    )

    & $Context.ReportProgress $Percent $Status
    if ($LogMessage) {
        & $Context.WriteLog $LogMessage $Level
    }
}

function Invoke-ToolkitCommand {
    param(
        [Parameter(Mandatory = $true)]$Context,
        [Parameter(Mandatory = $true)][string]$FilePath,
        [string]$Arguments = '',
        [string]$StepName = 'Executing command',
        [int]$StartPercent = 0,
        [int]$EndPercent = 100,
        [switch]$RequireSuccessExitCode
    )

    $commandLog = "{0}: {1} {2}" -f $StepName, $FilePath, $Arguments
    Invoke-TaskStep -Context $Context -Percent $StartPercent -Status $StepName -LogMessage $commandLog

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $FilePath
    $psi.Arguments = $Arguments
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $psi
    [void]$process.Start()

    while (-not $process.HasExited) {
        $next = [Math]::Min(($StartPercent + 5), ($EndPercent - 5))
        if ($next -gt $StartPercent) {
            Invoke-TaskStep -Context $Context -Percent $next -Status "$StepName (running)" -LogMessage "$StepName is still running."
        }
        Start-Sleep -Milliseconds 700
    }

    $stdout = $process.StandardOutput.ReadToEnd()
    $stderr = $process.StandardError.ReadToEnd()
    if ($stdout) { & $Context.WriteLog $stdout 'INFO' }
    if ($stderr) { & $Context.WriteLog $stderr 'WARN' }

    if ($RequireSuccessExitCode -and $process.ExitCode -ne 0) {
        throw "$StepName failed with exit code $($process.ExitCode)."
    }

    Invoke-TaskStep -Context $Context -Percent $EndPercent -Status "$StepName complete" -LogMessage "$StepName finished with exit code $($process.ExitCode)."
}

function Invoke-ToolkitScriptTask {
    param(
        [Parameter(Mandatory = $true)]$Context,
        [Parameter(Mandatory = $true)][string]$ScriptName,
        [Parameter(Mandatory = $true)][string]$StepName
    )

    $scriptPath = Join-Path $Context.ToolkitRoot "Scripts\Tasks\$ScriptName"
    if (-not (Test-Path $scriptPath)) {
        throw "Task script not found: $scriptPath"
    }

    $args = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
    Invoke-ToolkitCommand -Context $Context -FilePath 'powershell.exe' -Arguments $args -StepName $StepName -StartPercent 5 -EndPercent 100 -RequireSuccessExitCode
}

function Get-LegacyScriptTasks {
    $legacyRoot = Join-Path (Get-ToolkitRoot) 'LegacyScripts'
    if (-not (Test-Path $legacyRoot)) {
        return @()
    }

    $scripts = Get-ChildItem -Path $legacyRoot -Filter '*.ps1' | Sort-Object Name
    $tasks = @()
    foreach ($script in $scripts) {
        $safeId = ($script.BaseName -replace '[^a-zA-Z0-9]', '')
        $tasks += [pscustomobject]@{
            Id = "Legacy.$safeId"
            Name = "Legacy: $($script.BaseName)"
            Category = 'Legacy Imported Scripts'
            Description = "Runs imported script '$($script.Name)' in background mode."
            RequiresAdmin = $false
            Handler = {
                param($Context, $Task)
                $scriptName = ($Task.Name -replace '^Legacy: ', '') + '.ps1'
                $path = Join-Path $Context.LegacyScriptsPath $scriptName
                if (-not (Test-Path $path)) {
                    throw "Legacy script '$scriptName' was not found at $path"
                }

                Invoke-TaskStep -Context $Context -Percent 5 -Status "Launching $scriptName" -LogMessage "Launching legacy script $path"
                $elapsed = 0
                $job = Start-Job -ScriptBlock {
                    param($FilePath)
                    try {
                        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $FilePath
                    }
                    catch {
                        throw $_
                    }
                } -ArgumentList $path

                while ($job.State -eq 'Running') {
                    $elapsed += 1
                    $percent = [Math]::Min(95, (10 + ($elapsed * 2)))
                    Invoke-TaskStep -Context $Context -Percent $percent -Status "$scriptName is running" -LogMessage "$scriptName still running..."
                    Start-Sleep -Seconds 1
                }

                $output = Receive-Job -Job $job -Keep -ErrorAction SilentlyContinue
                if ($output) {
                    foreach ($line in $output) {
                        & $Context.WriteLog "$line" 'INFO'
                    }
                }

                if ($job.State -ne 'Completed') {
                    throw "Legacy script $scriptName did not complete successfully. Job state: $($job.State)"
                }

                Remove-Job -Job $job -Force
                Invoke-TaskStep -Context $Context -Percent 100 -Status "$scriptName complete" -LogMessage "$scriptName finished."
            }.GetNewClosure()
        }
    }

    return $tasks
}

function Get-ToolkitTaskCatalog {
    $coreTasks = @(
        [pscustomobject]@{
            Id = 'Network.Diagnostics'
            Name = 'Network Diagnostics'
            Category = 'Network'
            Description = 'Collects network health diagnostics, DNS status, and adapter details.'
            RequiresAdmin = $false
            Handler = {
                param($Context, $Task)
                Invoke-ToolkitScriptTask -Context $Context -ScriptName 'Invoke-NetworkMaintenance.ps1' -StepName 'Network diagnostics and maintenance'
            }
        },
        [pscustomobject]@{
            Id = 'Network.ResetStack'
            Name = 'Reset Network Stack'
            Category = 'Network'
            Description = 'Resets Winsock and TCP/IP stack for troubleshooting persistent network failures.'
            RequiresAdmin = $true
            Handler = {
                param($Context, $Task)
                Invoke-ToolkitScriptTask -Context $Context -ScriptName 'Invoke-ResetNetworkStack.ps1' -StepName 'Reset network stack'
                & $Context.WriteStatus 'Network stack reset complete. A reboot is recommended.'
            }
        },
        [pscustomobject]@{
            Id = 'Cleanup.TempFiles'
            Name = 'Clear Temp and Junk Files'
            Category = 'Cleanup'
            Description = 'Clears temp files from user and system temp folders.'
            RequiresAdmin = $true
            Handler = {
                param($Context, $Task)
                Invoke-ToolkitScriptTask -Context $Context -ScriptName 'Invoke-ClearTempJunk.ps1' -StepName 'Clear temp and junk files'
            }
        },
        [pscustomobject]@{
            Id = 'Cleanup.FreeSpace'
            Name = 'Free Space on C Drive'
            Category = 'Cleanup'
            Description = 'Runs component cleanup to reclaim disk space.'
            RequiresAdmin = $true
            Handler = {
                param($Context, $Task)
                Invoke-ToolkitScriptTask -Context $Context -ScriptName 'Invoke-FreeCDriveSpace.ps1' -StepName 'Free space on C drive'
            }
        },
        [pscustomobject]@{
            Id = 'Disk.Monitor'
            Name = 'Disk Space Monitor'
            Category = 'Storage'
            Description = 'Reports free space for local drives and flags low-space disks.'
            RequiresAdmin = $false
            Handler = {
                param($Context, $Task)
                Invoke-ToolkitScriptTask -Context $Context -ScriptName 'Invoke-DiskSpaceMonitor.ps1' -StepName 'Disk space monitoring'
            }
        },
        [pscustomobject]@{
            Id = 'Update.AllApps'
            Name = 'Update All Installed Apps'
            Category = 'Updates'
            Description = 'Runs winget upgrade for all supported packages.'
            RequiresAdmin = $true
            Handler = {
                param($Context, $Task)
                Invoke-ToolkitScriptTask -Context $Context -ScriptName 'Invoke-UpdateAllApps.ps1' -StepName 'Update all installed apps'
            }
        },
        [pscustomobject]@{
            Id = 'Repair.WindowsHealth'
            Name = 'Windows Repair Checks'
            Category = 'Repair'
            Description = 'Runs DISM health checks and SFC integrity validation.'
            RequiresAdmin = $true
            Handler = {
                param($Context, $Task)
                Invoke-ToolkitScriptTask -Context $Context -ScriptName 'Invoke-WindowsRepairChecks.ps1' -StepName 'Windows repair checks'
            }
        },
        [pscustomobject]@{
            Id = 'Apps.DebloatHelper'
            Name = 'Debloat / Uninstall Helper'
            Category = 'Applications'
            Description = 'Exports installed application lists for safe manual review and uninstall.'
            RequiresAdmin = $false
            Handler = {
                param($Context, $Task)
                Invoke-ToolkitScriptTask -Context $Context -ScriptName 'Invoke-DebloatInventory.ps1' -StepName 'Debloat and uninstall helper'
            }
        },
        [pscustomobject]@{
            Id = 'Integration.Microsoft365'
            Name = 'Install Microsoft 365 (Integration)'
            Category = 'Deployment'
            Description = 'Downloads and launches mallockey/Install-Microsoft365 workflow.'
            RequiresAdmin = $true
            Handler = {
                param($Context, $Task)
                $scriptPath = Join-Path $Context.IntegrationsPath 'Invoke-InstallMicrosoft365.ps1'
                if (-not (Test-Path $scriptPath)) {
                    throw "Missing integration script: $scriptPath"
                }

                Invoke-TaskStep -Context $Context -Percent 10 -Status 'Preparing Microsoft 365 workflow' -LogMessage 'Launching Install-Microsoft365 integration script.'
                & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $scriptPath
                Invoke-TaskStep -Context $Context -Percent 100 -Status 'Microsoft 365 integration complete' -LogMessage 'Install-Microsoft365 workflow finished.'
            }
        },
        [pscustomobject]@{
            Id = 'Integration.MediaCreationTool'
            Name = 'Windows Media Creation Tool Workflow'
            Category = 'Deployment'
            Description = 'Uses official Microsoft download links and opens Windows install media workflow.'
            RequiresAdmin = $false
            Handler = {
                param($Context, $Task)
                $scriptPath = Join-Path $Context.IntegrationsPath 'Invoke-MediaCreationWorkflow.ps1'
                if (-not (Test-Path $scriptPath)) {
                    throw "Missing integration script: $scriptPath"
                }

                Invoke-TaskStep -Context $Context -Percent 15 -Status 'Preparing media creation workflow' -LogMessage 'Launching Media Creation workflow helper.'
                & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $scriptPath
                Invoke-TaskStep -Context $Context -Percent 100 -Status 'Media creation workflow ready' -LogMessage 'Media creation integration completed.'
            }
        }
    )

    return @($coreTasks + (Get-LegacyScriptTasks))
}

Export-ModuleMember -Function @(
    'Get-ToolkitRoot',
    'Get-ToolkitSettings',
    'Test-ToolkitIsAdmin',
    'Write-ToolkitLog',
    'New-ToolkitTaskContext',
    'Get-ToolkitTaskCatalog',
    'Invoke-ToolkitTaskById'
)
