Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$updateToolsModule = Join-Path (Join-Path $PSScriptRoot '..') 'UpdateTools\UpdateTools.psm1'
$storageToolsModule = Join-Path (Join-Path $PSScriptRoot '..') 'StorageTools\StorageTools.psm1'
$hardwareDiagnosticsModule = Join-Path (Join-Path $PSScriptRoot '..') 'HardwareDiagnostics\HardwareDiagnostics.psm1'

foreach ($moduleToImport in @($updateToolsModule, $storageToolsModule, $hardwareDiagnosticsModule)) {
    if (Test-Path -LiteralPath $moduleToImport) {
        Import-Module $moduleToImport -Force -ErrorAction SilentlyContinue | Out-Null
    }
}

function Resolve-ToolkitPath {
    param(
        [Parameter(Mandatory = $true)][string]$Path
    )

    return [string](Convert-Path -Path $Path)
}

function Get-ToolkitRoot {
    $modulePath = $PSScriptRoot
    return (Resolve-ToolkitPath -Path (Join-Path $modulePath '..\..'))
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
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Message,
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
    $stdoutLines = @()
    $stderrLines = @()
    if ($stdout) {
        $stdoutLines = @($stdout -split "(`r`n|`n|`r)" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        & $Context.WriteLog $stdout 'INFO'
    }
    if ($stderr) {
        $stderrLines = @($stderr -split "(`r`n|`n|`r)" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        & $Context.WriteLog $stderr 'WARN'
    }

    if ($RequireSuccessExitCode -and $process.ExitCode -ne 0) {
        $detailLines = @($stderrLines + $stdoutLines) |
            Where-Object { $_ -and $_ -notmatch '^\s*(At |CategoryInfo|FullyQualifiedErrorId|\+ )' } |
            Select-Object -Last 4
        $detail = if ($detailLines.Count -gt 0) { ' Details: ' + ($detailLines -join ' ') } else { '' }
        throw "$StepName failed with exit code $($process.ExitCode).$detail"
    }

    Invoke-TaskStep -Context $Context -Percent $EndPercent -Status "$StepName complete" -LogMessage "$StepName finished with exit code $($process.ExitCode)."
}

function Invoke-ToolkitScriptTask {
    param(
        [Parameter(Mandatory = $true)]$Context,
        [Parameter(Mandatory = $true)][string]$ScriptName,
        [Parameter(Mandatory = $true)][string]$StepName,
        [string]$AdditionalArguments = ''
    )

    $scriptPath = Join-Path $Context.ToolkitRoot "Scripts\Tasks\$ScriptName"
    if (-not (Test-Path $scriptPath)) {
        throw "Task script not found: $scriptPath"
    }

    $args = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
    if ($AdditionalArguments) {
        $args = "$args $AdditionalArguments"
    }
    Invoke-ToolkitCommand -Context $Context -FilePath 'powershell.exe' -Arguments $args -StepName $StepName -StartPercent 5 -EndPercent 100 -RequireSuccessExitCode
}

function Get-ToolkitCategoryOrder {
    return @(
        'Applications',
        'Cleanup',
        'Deployment',
        'Hardware Diagnostics',
        'Misc/Utility',
        'Network Tools',
        'OneNote & Documents',
        'Repair',
        'Storage / Setup',
        'Update Tools'
    )
}

function Get-LegacyScriptDefinition {
    param(
        [Parameter(Mandatory = $true)][string]$ScriptBaseName
    )

    $name = $ScriptBaseName.ToLowerInvariant()

    switch ($name) {
        'bootcertchcker' {
            return @{
                Category = 'Misc/Utility'
                Description = 'Checks current Secure Boot and Windows UEFI CA 2023 certificate presence.'
                RequiresAdmin = $false
            }
        }
        'check-securebootcert' {
            return @{
                Category = 'Misc/Utility'
                Description = 'Runs a detailed Secure Boot certificate audit and reports the outcome.'
                RequiresAdmin = $true
            }
        }
        'invoke-ps2exe' {
            return @{
                Category = 'Misc/Utility'
                Description = 'Legacy PS2EXE builder workflow for packaging PowerShell scripts into executables.'
                RequiresAdmin = $false
            }
        }
        'the network access script' {
            return @{
                Category = 'Network Tools'
                Description = 'Legacy quick-connect helper for opening a remote computer administrative share.'
                RequiresAdmin = $false
            }
        }
        'fix_onenote_duplicates' {
            return @{
                Category = 'OneNote & Documents'
                Description = 'Removes duplicate OneNote content from the configured notebook workflow.'
                RequiresAdmin = $false
            }
        }
        'how_to_guide' {
            return @{
                Category = 'OneNote & Documents'
                Description = 'Reference guide for organizing imported documentation in OneNote.'
                RequiresAdmin = $false
            }
        }
        'invoke-massduplicatecleanup' {
            return @{
                Category = 'OneNote & Documents'
                Description = 'Bulk duplicate cleanup workflow for OneNote documentation pages.'
                RequiresAdmin = $false
            }
        }
        'launch_onenote' {
            return @{
                Category = 'OneNote & Documents'
                Description = 'Launches OneNote and the configured documentation notebook workflow.'
                RequiresAdmin = $false
            }
        }
        'master audit & smart-skip' {
            return @{
                Category = 'OneNote & Documents'
                Description = 'Audits source documentation and skips items already processed into OneNote.'
                RequiresAdmin = $false
            }
        }
        'master_onenote_importer' {
            return @{
                Category = 'OneNote & Documents'
                Description = 'Imports structured documentation content into a OneNote notebook.'
                RequiresAdmin = $false
            }
        }
        'move_to_onenote' {
            return @{
                Category = 'OneNote & Documents'
                Description = 'Moves staged content into OneNote using the original legacy import workflow.'
                RequiresAdmin = $false
            }
        }
        'move_to_onenote_2' {
            return @{
                Category = 'OneNote & Documents'
                Description = 'Secondary OneNote move/import variant from the legacy script set.'
                RequiresAdmin = $false
            }
        }
        'move_to_onenote_selfcheck' {
            return @{
                Category = 'OneNote & Documents'
                Description = 'Validates the OneNote move/import workflow before processing content.'
                RequiresAdmin = $false
            }
        }
        'sortdoc' {
            return @{
                Category = 'OneNote & Documents'
                Description = 'Sorts incoming documents into the expected staging structure.'
                RequiresAdmin = $false
            }
        }
        default {
            return @{
                Category = 'Applications'
                Description = "Runs imported legacy script '$ScriptBaseName.ps1' in background mode."
                RequiresAdmin = $false
            }
        }
    }
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
        $definition = Get-LegacyScriptDefinition -ScriptBaseName $script.BaseName
        $tasks += [pscustomobject]@{
            Id = "Legacy.$safeId"
            Name = "Legacy: $($script.BaseName)"
            Category = $definition.Category
            Description = $definition.Description
            RequiresAdmin = $definition.RequiresAdmin
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
                        if ([string]::IsNullOrWhiteSpace([string]$line)) {
                            continue
                        }
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
            Id = 'Hardware.MouseKeyboardTest'
            Name = 'Mouse & Keyboard Test'
            Category = 'Hardware Diagnostics'
            Description = 'Opens an activity test window that detects mouse movement, clicks, and keyboard input.'
            RequiresAdmin = $false
            Handler = {
                param($Context, $Task)
                Invoke-TaskStep -Context $Context -Percent 10 -Status 'Opening mouse and keyboard test' -LogMessage 'Launching mouse and keyboard diagnostic window.'
                $result = Test-TtkMouseKeyboardActivity
                if ($result) {
                    & $Context.WriteLog $result 'INFO'
                }
                Invoke-TaskStep -Context $Context -Percent 100 -Status 'Mouse and keyboard test complete' -LogMessage 'Mouse and keyboard diagnostic window closed.'
            }
        },
        [pscustomobject]@{
            Id = 'Hardware.MonitorPixelTest'
            Name = 'Monitor Dead Pixel Test'
            Category = 'Hardware Diagnostics'
            Description = 'Launches a full-screen color cycling tool for dead-pixel and stuck-pixel inspection. Press ESC to exit.'
            RequiresAdmin = $false
            Handler = {
                param($Context, $Task)
                Invoke-TaskStep -Context $Context -Percent 10 -Status 'Opening monitor pixel test' -LogMessage 'Launching monitor dead-pixel tester.'
                $result = Start-TtkMonitorPixelTest
                if ($result) {
                    & $Context.WriteLog $result 'INFO'
                }
                Invoke-TaskStep -Context $Context -Percent 100 -Status 'Monitor pixel test complete' -LogMessage 'Monitor dead-pixel tester closed.'
            }
        },
        [pscustomobject]@{
            Id = 'Network.Diagnostics'
            Name = 'Network Diagnostics'
            Category = 'Network Tools'
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
            Category = 'Network Tools'
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
            Category = 'Storage / Setup'
            Description = 'Reports free space for local drives and flags low-space disks.'
            RequiresAdmin = $false
            Handler = {
                param($Context, $Task)
                Invoke-ToolkitScriptTask -Context $Context -ScriptName 'Invoke-DiskSpaceMonitor.ps1' -StepName 'Disk space monitoring'
            }
        },
        [pscustomobject]@{
            Id = 'Storage.CloneGuide'
            Name = 'Clone Disk Guide'
            Category = 'Storage / Setup'
            Description = 'Shows the guided disk-cloning checklist and safety steps before running an imaging workflow.'
            RequiresAdmin = $false
            Handler = {
                param($Context, $Task)
                Invoke-TaskStep -Context $Context -Percent 20 -Status 'Preparing clone disk guidance' -LogMessage 'Opening clone disk guidance.'
                foreach ($line in (Invoke-TtkCloneDiskGuide)) {
                    & $Context.WriteLog $line 'INFO'
                }
                Invoke-TaskStep -Context $Context -Percent 100 -Status 'Clone disk guidance complete' -LogMessage 'Clone disk guidance displayed in the log.'
            }
        },
        [pscustomobject]@{
            Id = 'Storage.NewComputerSetup'
            Name = 'New Computer Setup Checklist'
            Category = 'Storage / Setup'
            Description = 'Writes the standard new-computer setup checklist to the execution log for guided deployment work.'
            RequiresAdmin = $false
            Handler = {
                param($Context, $Task)
                Invoke-TaskStep -Context $Context -Percent 20 -Status 'Preparing new computer setup checklist' -LogMessage 'Opening new computer setup checklist.'
                foreach ($line in (Invoke-TtkNewComputerSetupChecklist)) {
                    & $Context.WriteLog $line 'INFO'
                }
                Invoke-TaskStep -Context $Context -Percent 100 -Status 'New computer setup checklist complete' -LogMessage 'New computer setup checklist written to the log.'
            }
        },
        [pscustomobject]@{
            Id = 'Update.AllApps'
            Name = 'Update All Installed Apps'
            Category = 'Update Tools'
            Description = 'Runs winget upgrade for all supported packages.'
            RequiresAdmin = $true
            Handler = {
                param($Context, $Task)
                Invoke-ToolkitScriptTask -Context $Context -ScriptName 'Invoke-UpdateAllApps.ps1' -StepName 'Update all installed apps'
            }
        },
        [pscustomobject]@{
            Id = 'Update.VendorBIOS'
            Name = 'BIOS Update Tool'
            Category = 'Update Tools'
            Description = 'Vendor BIOS updater for Dell, HP, and Lenovo systems only. Uses vendor-supported update tooling and blocks unsupported manufacturers.'
            RequiresAdmin = $true
            Handler = {
                param($Context, $Task)
                Invoke-ToolkitScriptTask -Context $Context -ScriptName 'Invoke-BiosUpdate.ps1' -StepName 'Vendor BIOS update'
            }
        },
        [pscustomobject]@{
            Id = 'Update.VendorFirmware'
            Name = 'Firmware Update Tool'
            Category = 'Update Tools'
            Description = 'Vendor firmware updater for Dell, HP, and Lenovo systems only. Uses vendor-supported update tooling and blocks unsupported manufacturers.'
            RequiresAdmin = $true
            Handler = {
                param($Context, $Task)
                Invoke-ToolkitScriptTask -Context $Context -ScriptName 'Invoke-FirmwareUpdate.ps1' -StepName 'Vendor firmware update'
            }
        },
        [pscustomobject]@{
            Id = 'Update.VendorDrivers'
            Name = 'Driver Update Tool'
            Category = 'Update Tools'
            Description = 'Vendor driver updater for Dell, HP, and Lenovo systems only. Uses vendor-supported update tooling and blocks unsupported manufacturers.'
            RequiresAdmin = $true
            Handler = {
                param($Context, $Task)
                Invoke-ToolkitScriptTask -Context $Context -ScriptName 'Invoke-DriverUpdate.ps1' -StepName 'Vendor driver update'
            }
        },
        [pscustomobject]@{
            Id = 'Update.WindowsOS'
            Name = 'Windows Update Tool'
            Category = 'Update Tools'
            Description = 'Scans, downloads, and installs Windows updates. Supports skip-file based exclusions for updates you do not want installed.'
            RequiresAdmin = $true
            Handler = {
                param($Context, $Task)
                $extraArguments = ''
                if ($Context.PSObject.Properties.Name -contains 'UserInput') {
                    $selectionFile = $Context.UserInput.SelectionFile
                    if ($selectionFile) {
                        $extraArguments = "-SkipSelectionFile `"$selectionFile`""
                    }
                }
                Invoke-ToolkitScriptTask -Context $Context -ScriptName 'Invoke-WindowsUpdateTool.ps1' -StepName 'Windows Update tool' -AdditionalArguments $extraArguments
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
                $extraArguments = ''
                if ($Context.PSObject.Properties.Name -contains 'UserInput') {
                    $selectionFile = $Context.UserInput.SelectionFile
                    if ($selectionFile) {
                        $extraArguments = "-SelectionFile `"$selectionFile`""
                    }
                }
                Invoke-ToolkitScriptTask -Context $Context -ScriptName 'Invoke-DebloatInventory.ps1' -StepName 'Debloat and uninstall helper' -AdditionalArguments $extraArguments
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

                Invoke-TaskStep -Context $Context -Percent 5 -Status 'Preparing Microsoft 365 workflow' -LogMessage 'Preparing mallockey/Install-Microsoft365 download and launch workflow.'
                Invoke-TaskStep -Context $Context -Percent 15 -Status 'Downloading Microsoft 365 installer workflow' -LogMessage 'The integration script will download https://github.com/mallockey/Install-Microsoft365/archive/refs/heads/main.zip and locate the installer entry point.'
                Invoke-ToolkitCommand -Context $Context -FilePath 'powershell.exe' -Arguments "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`"" -StepName 'mallockey Install-Microsoft365 workflow' -StartPercent 20 -EndPercent 100 -RequireSuccessExitCode
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
    'Get-ToolkitCategoryOrder',
    'Get-ToolkitSettings',
    'Test-ToolkitIsAdmin',
    'Write-ToolkitLog',
    'New-ToolkitTaskContext',
    'Get-ToolkitTaskCatalog',
    'Invoke-ToolkitTaskById',
    'Invoke-TaskStep',
    'Invoke-ToolkitCommand',
    'Invoke-ToolkitScriptTask'
)
