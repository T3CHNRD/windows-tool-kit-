[CmdletBinding()]
param(
    [switch]$ElevatedRelaunch
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Test-LauncherIsAdmin {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-LauncherCommandPath {
    $commandInfo = $MyInvocation.MyCommand
    if ($commandInfo -and $commandInfo.PSObject.Properties['Path']) {
        $commandPath = $commandInfo.Path
        if ($commandPath) {
            return $commandPath
        }
    }

    $commandLineArgs = [Environment]::GetCommandLineArgs()
    if ($commandLineArgs.Count -gt 0) {
        $processPath = $commandLineArgs[0]
        if ($processPath -and (Test-Path -LiteralPath $processPath)) {
            return (Get-Item -LiteralPath $processPath).FullName
        }
    }

    return $null
}

function Get-LauncherRootPath {
    $commandPath = Get-LauncherCommandPath
    if ($commandPath) {
        return (Split-Path -Parent $commandPath)
    }

    return [System.AppDomain]::CurrentDomain.BaseDirectory.TrimEnd('\')
}

function Restart-LauncherElevated {
    param(
        [Parameter(Mandatory = $true)][string]$ToolkitRoot
    )

    $commandPath = Get-LauncherCommandPath
    if (-not $commandPath) {
        throw 'Could not determine the launcher path for elevation.'
    }

    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.Verb = 'runas'
    $startInfo.WorkingDirectory = $ToolkitRoot
    $startInfo.UseShellExecute = $true

    if ([IO.Path]::GetExtension($commandPath).Equals('.exe', [System.StringComparison]::OrdinalIgnoreCase)) {
        $startInfo.FileName = $commandPath
        if ($ElevatedRelaunch) {
            $startInfo.Arguments = '-ElevatedRelaunch'
        }
    }
    else {
        $startInfo.FileName = 'powershell.exe'
        $quotedPath = '"' + $commandPath + '"'
        $startInfo.Arguments = "-NoProfile -ExecutionPolicy Bypass -File $quotedPath -ElevatedRelaunch"
    }

    [void][System.Diagnostics.Process]::Start($startInfo)
}

$toolkitRoot = Get-LauncherRootPath

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

if (-not (Test-LauncherIsAdmin)) {
    try {
        Restart-LauncherElevated -ToolkitRoot $toolkitRoot
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "The toolkit needs administrator rights to run maintenance tasks. Elevation was not granted.`r`n`r`n$($_.Exception.Message)",
            "T3CHNRD'S Windows Tool Kit",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
    }
    return
}

$modulePath = Join-Path $toolkitRoot 'Modules\MaintenanceToolkit\MaintenanceToolkit.psm1'
Import-Module $modulePath -Force

[System.Windows.Forms.Application]::EnableVisualStyles()

$tasks = @(Get-ToolkitTaskCatalog | Sort-Object Category, Name)
$categoryOrder = @(Get-ToolkitCategoryOrder | Where-Object { $tasks.Category -contains $_ })
$tasksByCategory = @{}
foreach ($category in $categoryOrder) {
    $tasksByCategory[$category] = @($tasks | Where-Object { $_.Category -eq $category })
}

$script:SelectedTask = $null
$script:DebloatSelections = @()
$script:DebloatInventory = @()
$script:TaskButtons = @{}
$script:TaskButtonList = New-Object System.Collections.Generic.List[System.Windows.Forms.Button]

$form = New-Object System.Windows.Forms.Form
$form.Text = "T3CHNRD'S Windows Tool Kit"
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object System.Drawing.Size(1380, 860)
$form.MinimumSize = New-Object System.Drawing.Size(1240, 760)
$form.BackColor = [System.Drawing.Color]::FromArgb(244, 247, 252)
$form.Font = New-Object System.Drawing.Font('Segoe UI', 10)
$form.KeyPreview = $true

$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.AutoPopDelay = 12000
$toolTip.InitialDelay = 350
$toolTip.ReshowDelay = 200
$toolTip.ShowAlways = $true

$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Dock = 'Top'
$headerPanel.Height = 110
$headerPanel.BackColor = [System.Drawing.Color]::FromArgb(26, 45, 78)
$form.Controls.Add($headerPanel)

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "T3CHNRD'S Windows Tool Kit"
$titleLabel.ForeColor = [System.Drawing.Color]::White
$titleLabel.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 20, [System.Drawing.FontStyle]::Bold)
$titleLabel.AutoSize = $true
$titleLabel.Location = New-Object System.Drawing.Point(24, 18)
$headerPanel.Controls.Add($titleLabel)

$subLabel = New-Object System.Windows.Forms.Label
$subLabel.Text = 'Unified repair, cleanup, update, deployment, diagnostics, and imported script runner'
$subLabel.ForeColor = [System.Drawing.Color]::FromArgb(202, 216, 240)
$subLabel.Font = New-Object System.Drawing.Font('Segoe UI', 11)
$subLabel.AutoSize = $true
$subLabel.Location = New-Object System.Drawing.Point(26, 62)
$headerPanel.Controls.Add($subLabel)

$contentSplit = New-Object System.Windows.Forms.SplitContainer
$contentSplit.Dock = 'Fill'
$contentSplit.BackColor = [System.Drawing.Color]::FromArgb(232, 237, 245)
$contentSplit.FixedPanel = 'Panel1'
$form.Controls.Add($contentSplit)

$leftPanel = New-Object System.Windows.Forms.Panel
$leftPanel.Dock = 'Fill'
$leftPanel.Padding = New-Object System.Windows.Forms.Padding(16, 16, 12, 16)
$leftPanel.BackColor = [System.Drawing.Color]::FromArgb(244, 247, 252)
$contentSplit.Panel1.Controls.Add($leftPanel)

$leftTitle = New-Object System.Windows.Forms.Label
$leftTitle.Text = 'Available Tools'
$leftTitle.AutoSize = $true
$leftTitle.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 14, [System.Drawing.FontStyle]::Bold)
$leftTitle.Dock = 'Top'
$leftPanel.Controls.Add($leftTitle)

$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = 'Fill'
$tabControl.Padding = New-Object System.Drawing.Point(16, 6)
$leftPanel.Controls.Add($tabControl)
$leftPanel.Controls.SetChildIndex($tabControl, 1)

$rightPanel = New-Object System.Windows.Forms.Panel
$rightPanel.Dock = 'Fill'
$rightPanel.Padding = New-Object System.Windows.Forms.Padding(18, 16, 18, 16)
$rightPanel.BackColor = [System.Drawing.Color]::White
$contentSplit.Panel2.Controls.Add($rightPanel)

$detailsLayout = New-Object System.Windows.Forms.TableLayoutPanel
$detailsLayout.Dock = 'Fill'
$detailsLayout.ColumnCount = 1
$detailsLayout.RowCount = 11
$detailsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$detailsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$detailsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 135)))
$detailsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$detailsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 28)))
$detailsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$detailsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$detailsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 56)))
$detailsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$detailsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$detailsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$rightPanel.Controls.Add($detailsLayout)

$selectedToolLabel = New-Object System.Windows.Forms.Label
$selectedToolLabel.Text = 'Tool Details'
$selectedToolLabel.AutoSize = $true
$selectedToolLabel.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 15, [System.Drawing.FontStyle]::Bold)
$selectedToolLabel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 6)
$detailsLayout.Controls.Add($selectedToolLabel, 0, 0)

$descriptionHintLabel = New-Object System.Windows.Forms.Label
$descriptionHintLabel.Text = 'Hover over a tool to preview it. Click a tool to run it.'
$descriptionHintLabel.AutoSize = $true
$descriptionHintLabel.ForeColor = [System.Drawing.Color]::FromArgb(84, 96, 118)
$descriptionHintLabel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 8)
$detailsLayout.Controls.Add($descriptionHintLabel, 0, 1)

$descriptionBox = New-Object System.Windows.Forms.TextBox
$descriptionBox.Multiline = $true
$descriptionBox.ReadOnly = $true
$descriptionBox.Dock = 'Fill'
$descriptionBox.BorderStyle = 'FixedSingle'
$descriptionBox.BackColor = [System.Drawing.Color]::FromArgb(249, 251, 255)
$descriptionBox.ScrollBars = 'Vertical'
$descriptionBox.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 12)
$detailsLayout.Controls.Add($descriptionBox, 0, 2)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = 'Status: Ready'
$statusLabel.AutoSize = $true
$statusLabel.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 11)
$statusLabel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 6)
$detailsLayout.Controls.Add($statusLabel, 0, 3)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Dock = 'Fill'
$progressBar.Minimum = 0
$progressBar.Maximum = 100
$progressBar.Value = 0
$progressBar.Style = 'Continuous'
$progressBar.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 12)
$detailsLayout.Controls.Add($progressBar, 0, 4)

$requiresAdminLabel = New-Object System.Windows.Forms.Label
$requiresAdminLabel.Text = 'Requires Admin: Yes'
$requiresAdminLabel.AutoSize = $true
$requiresAdminLabel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 4)
$detailsLayout.Controls.Add($requiresAdminLabel, 0, 5)

$inputSummaryLabel = New-Object System.Windows.Forms.Label
$inputSummaryLabel.Text = 'Input: None required'
$inputSummaryLabel.AutoSize = $true
$inputSummaryLabel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 12)
$detailsLayout.Controls.Add($inputSummaryLabel, 0, 6)

$buttonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$buttonPanel.Dock = 'Fill'
$buttonPanel.FlowDirection = 'LeftToRight'
$buttonPanel.WrapContents = $false
$buttonPanel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)
$detailsLayout.Controls.Add($buttonPanel, 0, 7)

$runButton = New-Object System.Windows.Forms.Button
$runButton.Text = 'Run Selected Tool Again'
$runButton.Size = New-Object System.Drawing.Size(190, 38)
$runButton.BackColor = [System.Drawing.Color]::FromArgb(47, 108, 188)
$runButton.ForeColor = [System.Drawing.Color]::White
$runButton.FlatStyle = 'Flat'
$toolTip.SetToolTip($runButton, 'Runs the currently selected tool again.')
$buttonPanel.Controls.Add($runButton)

$chooseAppsButton = New-Object System.Windows.Forms.Button
$chooseAppsButton.Text = 'Choose Apps'
$chooseAppsButton.Size = New-Object System.Drawing.Size(130, 38)
$chooseAppsButton.FlatStyle = 'Flat'
$chooseAppsButton.Visible = $false
$toolTip.SetToolTip($chooseAppsButton, 'Pick the applications to include in the debloat review list.')
$buttonPanel.Controls.Add($chooseAppsButton)

$openLogsButton = New-Object System.Windows.Forms.Button
$openLogsButton.Text = 'Open Logs Folder'
$openLogsButton.Size = New-Object System.Drawing.Size(150, 38)
$openLogsButton.FlatStyle = 'Flat'
$buttonPanel.Controls.Add($openLogsButton)

$openRootButton = New-Object System.Windows.Forms.Button
$openRootButton.Text = 'Open Project Root'
$openRootButton.Size = New-Object System.Drawing.Size(150, 38)
$openRootButton.FlatStyle = 'Flat'
$buttonPanel.Controls.Add($openRootButton)

$logTitle = New-Object System.Windows.Forms.Label
$logTitle.Text = 'Execution Log'
$logTitle.AutoSize = $true
$logTitle.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 11, [System.Drawing.FontStyle]::Bold)
$logTitle.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 6)
$detailsLayout.Controls.Add($logTitle, 0, 8)

$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Multiline = $true
$logBox.ScrollBars = 'Vertical'
$logBox.ReadOnly = $true
$logBox.Dock = 'Fill'
$logBox.BorderStyle = 'FixedSingle'
$logBox.BackColor = [System.Drawing.Color]::FromArgb(247, 249, 252)
$detailsLayout.Controls.Add($logBox, 0, 9)

$footerLabel = New-Object System.Windows.Forms.Label
$footerLabel.Text = 'The launcher starts elevated automatically so maintenance tasks can run without manual relaunch steps.'
$footerLabel.AutoSize = $true
$footerLabel.ForeColor = [System.Drawing.Color]::FromArgb(98, 108, 124)
$footerLabel.Margin = New-Object System.Windows.Forms.Padding(0, 8, 0, 0)
$detailsLayout.Controls.Add($footerLabel, 0, 10)

$worker = New-Object System.ComponentModel.BackgroundWorker
$worker.WorkerReportsProgress = $true
$worker.WorkerSupportsCancellation = $false

function Add-UiLogLine {
    param(
        [Parameter(Mandatory = $true)][string]$Line
    )

    $timestamped = "[{0:HH:mm:ss}] {1}" -f (Get-Date), $Line
    $logBox.AppendText($timestamped + [Environment]::NewLine)
}

function Get-DebloatInventory {
    $apps = New-Object System.Collections.Generic.List[object]

    try {
        $wingetOutput = & winget.exe list --accept-source-agreements 2>$null
        foreach ($line in $wingetOutput) {
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            if ($line -match '^\s*Name\s+Id\s+Version') { continue }
            if ($line -match '^-{3,}') { continue }

            $name = ($line -replace '\s{2,}.*$', '').Trim()
            if (-not $name) { continue }

            $apps.Add([pscustomobject]@{
                Name = $name
                Source = 'Desktop app'
            })
        }
    }
    catch {
    }

    try {
        Get-AppxPackage | ForEach-Object {
            $apps.Add([pscustomobject]@{
                Name = $_.Name
                Source = 'AppX package'
            })
        }
    }
    catch {
    }

    return $apps |
        Sort-Object Name -Unique |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_.Name) }
}

function Show-DebloatChooser {
    if (-not $script:DebloatInventory -or $script:DebloatInventory.Count -eq 0) {
        $script:DebloatInventory = @(Get-DebloatInventory)
    }

    $picker = New-Object System.Windows.Forms.Form
    $picker.Text = 'Choose Apps For Debloat Review'
    $picker.StartPosition = 'CenterParent'
    $picker.Size = New-Object System.Drawing.Size(680, 560)
    $picker.MinimumSize = New-Object System.Drawing.Size(620, 500)
    $picker.Font = $form.Font

    $infoLabel = New-Object System.Windows.Forms.Label
    $infoLabel.Text = 'Check the apps you want included in the review list, then click Save and Run.'
    $infoLabel.AutoSize = $true
    $infoLabel.Location = New-Object System.Drawing.Point(14, 14)
    $picker.Controls.Add($infoLabel)

    $searchBox = New-Object System.Windows.Forms.TextBox
    $searchBox.Location = New-Object System.Drawing.Point(14, 42)
    $searchBox.Size = New-Object System.Drawing.Size(632, 28)
    $picker.Controls.Add($searchBox)

    $appChecklist = New-Object System.Windows.Forms.CheckedListBox
    $appChecklist.Location = New-Object System.Drawing.Point(14, 78)
    $appChecklist.Size = New-Object System.Drawing.Size(632, 374)
    $appChecklist.Anchor = 'Top,Bottom,Left,Right'
    $appChecklist.CheckOnClick = $true
    $picker.Controls.Add($appChecklist)

    $saveButton = New-Object System.Windows.Forms.Button
    $saveButton.Text = 'Save and Run'
    $saveButton.Size = New-Object System.Drawing.Size(140, 36)
    $saveButton.Location = New-Object System.Drawing.Point(356, 470)
    $saveButton.Anchor = 'Bottom,Right'
    $picker.Controls.Add($saveButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = 'Cancel'
    $cancelButton.Size = New-Object System.Drawing.Size(140, 36)
    $cancelButton.Location = New-Object System.Drawing.Point(506, 470)
    $cancelButton.Anchor = 'Bottom,Right'
    $picker.Controls.Add($cancelButton)

    function Populate-DebloatList {
        param([string]$FilterText)

        $appChecklist.Items.Clear()
        foreach ($app in $script:DebloatInventory) {
            if ($FilterText -and $app.Name -notlike "*$FilterText*") {
                continue
            }

            $label = '{0} [{1}]' -f $app.Name, $app.Source
            $isChecked = $script:DebloatSelections -contains $app.Name
            [void]$appChecklist.Items.Add($label, $isChecked)
        }
    }

    Populate-DebloatList -FilterText ''

    $searchBox.Add_TextChanged({
        Populate-DebloatList -FilterText $searchBox.Text.Trim()
    })

    $saveButton.Add_Click({
        $selected = New-Object System.Collections.Generic.List[string]
        foreach ($item in $appChecklist.CheckedItems) {
            $name = ($item -replace '\s\[[^\]]+\]$', '').Trim()
            if ($name) {
                $selected.Add($name)
            }
        }

        $script:DebloatSelections = @($selected | Sort-Object -Unique)
        $picker.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $picker.Close()
    })

    $cancelButton.Add_Click({
        $picker.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $picker.Close()
    })

    return ($picker.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK)
}

function Update-ToolButtonWidths {
    foreach ($button in $script:TaskButtonList) {
        $parentPanel = $button.Parent
        if ($parentPanel) {
            $button.Width = [Math]::Max(300, $parentPanel.ClientSize.Width - 28)
        }
    }
}

function Set-ToolButtonsEnabled {
    param(
        [Parameter(Mandatory = $true)][bool]$Enabled
    )

    foreach ($button in $script:TaskButtonList) {
        $button.Enabled = $Enabled
    }
}

function Update-SelectedToolVisuals {
    foreach ($taskId in $script:TaskButtons.Keys) {
        $button = $script:TaskButtons[$taskId]
        $isSelected = $script:SelectedTask -and ($script:SelectedTask.Id -eq $taskId)
        if ($isSelected) {
            $button.BackColor = [System.Drawing.Color]::FromArgb(47, 108, 188)
            $button.ForeColor = [System.Drawing.Color]::White
        }
        else {
            $button.BackColor = [System.Drawing.Color]::White
            $button.ForeColor = [System.Drawing.Color]::FromArgb(32, 42, 56)
        }
    }
}

function Set-ActiveTask {
    param(
        [Parameter(Mandatory = $true)]$Task,
        [string]$StatusOverride
    )

    $script:SelectedTask = $Task
    $selectedToolLabel.Text = $Task.Name
    $descriptionBox.Text = $Task.Description

    $adminText = 'No'
    if ($Task.RequiresAdmin) {
        $adminText = 'Yes'
    }
    $requiresAdminLabel.Text = "Requires Admin: $adminText"

    if ($Task.Id -eq 'Apps.DebloatHelper') {
        $selectionCount = $script:DebloatSelections.Count
        $inputSummaryLabel.Text = "Input: $selectionCount selected app(s) for review"
        $chooseAppsButton.Visible = $true
        if ($selectionCount -gt 0) {
            $descriptionBox.Text += "`r`n`r`nCurrent debloat selection count: $selectionCount"
        }
    }
    else {
        $inputSummaryLabel.Text = 'Input: None required'
        $chooseAppsButton.Visible = $false
    }

    if ($StatusOverride) {
        $statusLabel.Text = "Status: $StatusOverride"
    }
    elseif (-not $worker.IsBusy) {
        $statusLabel.Text = "Status: Ready to run $($Task.Name)"
    }

    Update-SelectedToolVisuals
}

function Start-ToolkitTask {
    param(
        [Parameter(Mandatory = $true)]$Task
    )

    if ($worker.IsBusy) {
        [System.Windows.Forms.MessageBox]::Show(
            'Another task is already running. Wait for it to finish before starting the next one.',
            'Toolkit',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        return
    }

    Set-ActiveTask -Task $Task -StatusOverride "Queued $($Task.Name)"

    if ($Task.Id -eq 'Apps.DebloatHelper') {
        $selectionSaved = Show-DebloatChooser
        Set-ActiveTask -Task $Task -StatusOverride "Queued $($Task.Name)"
        if (-not $selectionSaved) {
            $statusLabel.Text = "Status: Cancelled $($Task.Name)"
            return
        }
    }

    if ($Task.RequiresAdmin -and -not (Test-ToolkitIsAdmin)) {
        [System.Windows.Forms.MessageBox]::Show(
            'The toolkit is not currently elevated. Close it and relaunch the root EXE so it can auto-elevate.',
            'Toolkit',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    $progressBar.Value = 0
    $progressBar.Style = 'Marquee'
    $statusLabel.Text = "Status: Starting $($Task.Name)..."
    Add-UiLogLine -Line "Queued task: $($Task.Name)"

    $runButton.Enabled = $false
    $chooseAppsButton.Enabled = $false
    Set-ToolButtonsEnabled -Enabled $false
    $tabControl.Enabled = $false

    $taskPayload = [pscustomobject]@{
        Task = $Task
        UserInput = @{}
    }

    if ($Task.Id -eq 'Apps.DebloatHelper' -and $script:DebloatSelections.Count -gt 0) {
        $selectionPath = Join-Path (Join-Path $toolkitRoot 'Logs') 'debloat-selection.txt'
        Set-Content -LiteralPath $selectionPath -Value $script:DebloatSelections -Encoding UTF8
        $taskPayload.UserInput.SelectionFile = $selectionPath
    }

    $worker.RunWorkerAsync($taskPayload)
}

foreach ($category in $categoryOrder) {
    $tabPage = New-Object System.Windows.Forms.TabPage
    $tabPage.Text = $category
    $tabPage.BackColor = [System.Drawing.Color]::White
    $tabControl.TabPages.Add($tabPage) | Out-Null

    $flowPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $flowPanel.Dock = 'Fill'
    $flowPanel.Padding = New-Object System.Windows.Forms.Padding(12)
    $flowPanel.AutoScroll = $true
    $flowPanel.FlowDirection = 'TopDown'
    $flowPanel.WrapContents = $false
    $tabPage.Controls.Add($flowPanel)

    foreach ($task in $tasksByCategory[$category]) {
        $button = New-Object System.Windows.Forms.Button
        $button.Tag = $task
        $button.Text = if ($task.RequiresAdmin) { "$($task.Name)  [Admin]" } else { $task.Name }
        $button.Height = 46
        $button.Width = 360
        $button.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
        $button.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)
        $button.FlatStyle = 'Flat'
        $button.BackColor = [System.Drawing.Color]::White
        $button.ForeColor = [System.Drawing.Color]::FromArgb(32, 42, 56)
        $button.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(197, 208, 225)
        $button.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(235, 241, 250)
        $toolTip.SetToolTip($button, $task.Description)

        $button.Add_MouseEnter({
            Set-ActiveTask -Task $this.Tag
        })
        $button.Add_Click({
            Start-ToolkitTask -Task $this.Tag
        })

        $flowPanel.Controls.Add($button)
        $script:TaskButtons[$task.Id] = $button
        [void]$script:TaskButtonList.Add($button)
    }

    $flowPanel.Add_Resize({
        Update-ToolButtonWidths
    })
}

$worker.add_DoWork({
    param($sender, $e)

    $taskPayload = $e.Argument
    $task = $taskPayload.Task
    $context = New-ToolkitTaskContext `
        -ReportProgress {
            param($Percent, $Status)
            $sender.ReportProgress([Math]::Max(0, [Math]::Min(100, [int]$Percent)), [pscustomobject]@{
                Kind = 'Progress'
                Status = $Status
            })
        } `
        -WriteStatus {
            param($Status)
            $sender.ReportProgress(-1, [pscustomobject]@{
                Kind = 'Status'
                Status = $Status
            })
        } `
        -WriteLog {
            param($Message, $Level)
            $line = Write-ToolkitLog -Message $Message -Level $Level
            $sender.ReportProgress(-1, [pscustomobject]@{
                Kind = 'Log'
                Status = $line
            })
        }

    $context | Add-Member -NotePropertyName UserInput -NotePropertyValue $taskPayload.UserInput -Force

    try {
        & $context.WriteLog "Starting task: $($task.Name)" 'INFO'
        Invoke-ToolkitTaskById -TaskId $task.Id -Context $context
        & $context.WriteLog "Task completed successfully: $($task.Name)" 'INFO'
        $e.Result = [pscustomobject]@{ Success = $true; TaskName = $task.Name }
    }
    catch {
        $errorMessage = $_.Exception.Message
        & $context.WriteLog "Task failed: $errorMessage" 'ERROR'
        $e.Result = [pscustomobject]@{ Success = $false; TaskName = $task.Name; Error = $errorMessage }
    }
})

$worker.add_ProgressChanged({
    param($sender, $e)

    $payload = $e.UserState
    if ($payload -and $payload.Kind -eq 'Progress') {
        if ($progressBar.Style -ne 'Continuous') {
            $progressBar.Style = 'Continuous'
        }
        if ($e.ProgressPercentage -ge 0) {
            $progressBar.Value = [Math]::Max(0, [Math]::Min(100, $e.ProgressPercentage))
        }
        $statusLabel.Text = "Status: $($payload.Status)"
        return
    }

    if ($payload -and $payload.Kind -eq 'Status') {
        $statusLabel.Text = "Status: $($payload.Status)"
        return
    }

    if ($payload -and $payload.Kind -eq 'Log') {
        Add-UiLogLine -Line $payload.Status
    }
})

$worker.add_RunWorkerCompleted({
    param($sender, $e)

    $runButton.Enabled = $true
    $chooseAppsButton.Enabled = $true
    $tabControl.Enabled = $true
    Set-ToolButtonsEnabled -Enabled $true

    if ($e.Result -and $e.Result.Success) {
        $progressBar.Style = 'Continuous'
        $progressBar.Value = 100
        $statusLabel.Text = "Status: Completed - $($e.Result.TaskName)"
        [System.Windows.Forms.MessageBox]::Show(
            "Task completed: $($e.Result.TaskName)",
            'Toolkit',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }
    else {
        $progressBar.Style = 'Continuous'
        $progressBar.Value = 0
        $errorText = if ($e.Result) { $e.Result.Error } else { 'Unknown error' }
        $statusLabel.Text = "Status: Failed - $errorText"
        [System.Windows.Forms.MessageBox]::Show(
            "Task failed: $errorText",
            'Toolkit',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }

    if ($script:SelectedTask) {
        Set-ActiveTask -Task $script:SelectedTask -StatusOverride ($statusLabel.Text -replace '^Status:\s*', '')
    }
})

$runButton.Add_Click({
    if (-not $script:SelectedTask) {
        [System.Windows.Forms.MessageBox]::Show(
            'Hover over or click a tool first so the launcher knows which tool to run.',
            'Toolkit',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        return
    }

    Start-ToolkitTask -Task $script:SelectedTask
})

$chooseAppsButton.Add_Click({
    if ($script:SelectedTask -and $script:SelectedTask.Id -eq 'Apps.DebloatHelper') {
        if (Show-DebloatChooser) {
            Set-ActiveTask -Task $script:SelectedTask -StatusOverride 'Debloat selection updated'
        }
    }
})

$openLogsButton.Add_Click({
    $settings = Get-ToolkitSettings
    $logPath = Join-Path (Get-ToolkitRoot) $settings.LogRoot
    if (-not (Test-Path -LiteralPath $logPath)) {
        New-Item -Path $logPath -ItemType Directory -Force | Out-Null
    }
    Start-Process explorer.exe $logPath
})

$openRootButton.Add_Click({
    Start-Process explorer.exe (Get-ToolkitRoot)
})

$tabControl.Add_SelectedIndexChanged({
    if ($worker.IsBusy) {
        return
    }

    $activeTab = $tabControl.SelectedTab
    if (-not $activeTab) {
        return
    }

    $categoryTasks = $tasksByCategory[$activeTab.Text]
    if ($categoryTasks.Count -gt 0) {
        Set-ActiveTask -Task $categoryTasks[0]
    }
})

function Apply-SplitterLayout {
    $leftWidth = 430
    $minimumRightWidth = 500
    $availableWidth = $contentSplit.ClientSize.Width
    if ($availableWidth -le 0) {
        return
    }

    $contentSplit.Panel1MinSize = 360
    $contentSplit.Panel2MinSize = $minimumRightWidth
    $maxLeftWidth = [Math]::Max($contentSplit.Panel1MinSize, ($availableWidth - $minimumRightWidth))
    $contentSplit.SplitterDistance = [Math]::Max($contentSplit.Panel1MinSize, [Math]::Min($leftWidth, $maxLeftWidth))
}

$form.Add_Shown({
    Apply-SplitterLayout
    Update-ToolButtonWidths

    if ($categoryOrder.Count -gt 0) {
        $tabControl.SelectedTab = $tabControl.TabPages[0]
        $firstTask = $tasksByCategory[$categoryOrder[0]] | Select-Object -First 1
        if ($firstTask) {
            Set-ActiveTask -Task $firstTask
        }
    }
})

$form.Add_Resize({
    Apply-SplitterLayout
    Update-ToolButtonWidths
})

$form.Add_KeyDown({
    param($sender, $e)

    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::F5 -and $script:SelectedTask) {
        Start-ToolkitTask -Task $script:SelectedTask
        $e.Handled = $true
    }
})

[void]$form.ShowDialog()
