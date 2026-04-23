[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$toolkitRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$modulePath = Join-Path $toolkitRoot 'Modules\MaintenanceToolkit\MaintenanceToolkit.psm1'
Import-Module $modulePath -Force

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)

$tasks = Get-ToolkitTaskCatalog | Sort-Object Category, Name

$form = New-Object System.Windows.Forms.Form
$form.Text = "T3CHNRD'S Windows Tool Kit"
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object System.Drawing.Size(1100, 700)
$form.MinimumSize = New-Object System.Drawing.Size(1000, 650)
$form.BackColor = [System.Drawing.Color]::FromArgb(245, 248, 252)
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Dock = 'Top'
$headerPanel.Height = 80
$headerPanel.BackColor = [System.Drawing.Color]::FromArgb(26, 45, 78)
$form.Controls.Add($headerPanel)

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "T3CHNRD'S Windows Tool Kit"
$titleLabel.ForeColor = [System.Drawing.Color]::White
$titleLabel.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 16, [System.Drawing.FontStyle]::Bold)
$titleLabel.AutoSize = $true
$titleLabel.Location = New-Object System.Drawing.Point(20, 15)
$headerPanel.Controls.Add($titleLabel)

$subLabel = New-Object System.Windows.Forms.Label
$subLabel.Text = 'Unified repair, cleanup, update, deployment, and imported script runner'
$subLabel.ForeColor = [System.Drawing.Color]::FromArgb(195, 213, 241)
$subLabel.AutoSize = $true
$subLabel.Location = New-Object System.Drawing.Point(22, 48)
$headerPanel.Controls.Add($subLabel)

$leftPanel = New-Object System.Windows.Forms.Panel
$leftPanel.Dock = 'Left'
$leftPanel.Width = 500
$leftPanel.Padding = New-Object System.Windows.Forms.Padding(12)
$leftPanel.BackColor = [System.Drawing.Color]::FromArgb(245, 248, 252)
$form.Controls.Add($leftPanel)

$taskLabel = New-Object System.Windows.Forms.Label
$taskLabel.Text = 'Available Tools'
$taskLabel.AutoSize = $true
$taskLabel.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 11, [System.Drawing.FontStyle]::Bold)
$taskLabel.Location = New-Object System.Drawing.Point(8, 6)
$leftPanel.Controls.Add($taskLabel)

$taskList = New-Object System.Windows.Forms.ListView
$taskList.View = 'Details'
$taskList.FullRowSelect = $true
$taskList.HideSelection = $false
$taskList.MultiSelect = $false
$taskList.Location = New-Object System.Drawing.Point(8, 34)
$taskList.Size = New-Object System.Drawing.Size(470, 500)
$taskList.Anchor = 'Top,Bottom,Left,Right'
[void]$taskList.Columns.Add('Tool', 250)
[void]$taskList.Columns.Add('Category', 110)
[void]$taskList.Columns.Add('Admin', 60)
[void]$taskList.Columns.Add('ID', 0)
$leftPanel.Controls.Add($taskList)

foreach ($task in $tasks) {
    $item = New-Object System.Windows.Forms.ListViewItem($task.Name)
    [void]$item.SubItems.Add($task.Category)
    $adminText = 'No'
    if ($task.RequiresAdmin) { $adminText = 'Yes' }
    [void]$item.SubItems.Add($adminText)
    [void]$item.SubItems.Add($task.Id)
    [void]$taskList.Items.Add($item)
}

$rightPanel = New-Object System.Windows.Forms.Panel
$rightPanel.Dock = 'Fill'
$rightPanel.Padding = New-Object System.Windows.Forms.Padding(12)
$rightPanel.BackColor = [System.Drawing.Color]::White
$form.Controls.Add($rightPanel)

$descriptionTitle = New-Object System.Windows.Forms.Label
$descriptionTitle.Text = 'Task Details'
$descriptionTitle.AutoSize = $true
$descriptionTitle.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 11, [System.Drawing.FontStyle]::Bold)
$descriptionTitle.Location = New-Object System.Drawing.Point(8, 6)
$rightPanel.Controls.Add($descriptionTitle)

$descriptionBox = New-Object System.Windows.Forms.TextBox
$descriptionBox.Multiline = $true
$descriptionBox.ReadOnly = $true
$descriptionBox.BorderStyle = 'FixedSingle'
$descriptionBox.BackColor = [System.Drawing.Color]::FromArgb(251, 252, 255)
$descriptionBox.Location = New-Object System.Drawing.Point(8, 34)
$descriptionBox.Size = New-Object System.Drawing.Size(540, 90)
$descriptionBox.Anchor = 'Top,Left,Right'
$rightPanel.Controls.Add($descriptionBox)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = 'Status: Ready'
$statusLabel.AutoSize = $true
$statusLabel.Location = New-Object System.Drawing.Point(8, 138)
$statusLabel.Anchor = 'Top,Left,Right'
$rightPanel.Controls.Add($statusLabel)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(8, 160)
$progressBar.Size = New-Object System.Drawing.Size(540, 24)
$progressBar.Anchor = 'Top,Left,Right'
$progressBar.Minimum = 0
$progressBar.Maximum = 100
$progressBar.Value = 0
$rightPanel.Controls.Add($progressBar)

$requiresAdminLabel = New-Object System.Windows.Forms.Label
$requiresAdminLabel.Text = 'Requires Admin: No'
$requiresAdminLabel.AutoSize = $true
$requiresAdminLabel.Location = New-Object System.Drawing.Point(8, 194)
$rightPanel.Controls.Add($requiresAdminLabel)

$runButton = New-Object System.Windows.Forms.Button
$runButton.Text = 'Run Selected Task'
$runButton.Location = New-Object System.Drawing.Point(8, 222)
$runButton.Size = New-Object System.Drawing.Size(180, 34)
$runButton.BackColor = [System.Drawing.Color]::FromArgb(47, 108, 188)
$runButton.ForeColor = [System.Drawing.Color]::White
$runButton.FlatStyle = 'Flat'
$rightPanel.Controls.Add($runButton)

$openLogsButton = New-Object System.Windows.Forms.Button
$openLogsButton.Text = 'Open Logs Folder'
$openLogsButton.Location = New-Object System.Drawing.Point(198, 222)
$openLogsButton.Size = New-Object System.Drawing.Size(150, 34)
$openLogsButton.FlatStyle = 'Flat'
$rightPanel.Controls.Add($openLogsButton)

$openRootButton = New-Object System.Windows.Forms.Button
$openRootButton.Text = 'Open Project Root'
$openRootButton.Location = New-Object System.Drawing.Point(358, 222)
$openRootButton.Size = New-Object System.Drawing.Size(150, 34)
$openRootButton.FlatStyle = 'Flat'
$rightPanel.Controls.Add($openRootButton)

$logTitle = New-Object System.Windows.Forms.Label
$logTitle.Text = 'Execution Log'
$logTitle.AutoSize = $true
$logTitle.Location = New-Object System.Drawing.Point(8, 270)
$rightPanel.Controls.Add($logTitle)

$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Multiline = $true
$logBox.ScrollBars = 'Vertical'
$logBox.ReadOnly = $true
$logBox.BorderStyle = 'FixedSingle'
$logBox.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 253)
$logBox.Location = New-Object System.Drawing.Point(8, 292)
$logBox.Size = New-Object System.Drawing.Size(540, 275)
$logBox.Anchor = 'Top,Bottom,Left,Right'
$rightPanel.Controls.Add($logBox)

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

function Get-SelectedTaskObject {
    if ($taskList.SelectedItems.Count -eq 0) { return $null }
    $selectedId = $taskList.SelectedItems[0].SubItems[3].Text
    return $tasks | Where-Object { $_.Id -eq $selectedId } | Select-Object -First 1
}

$taskList.add_SelectedIndexChanged({
    $selectedTask = Get-SelectedTaskObject
    if (-not $selectedTask) {
        $descriptionBox.Text = ''
        $requiresAdminLabel.Text = 'Requires Admin: No'
        return
    }

    $descriptionBox.Text = $selectedTask.Description
    $selectedAdminText = 'No'
    if ($selectedTask.RequiresAdmin) { $selectedAdminText = 'Yes' }
    $requiresAdminLabel.Text = "Requires Admin: $selectedAdminText"
})

$openLogsButton.Add_Click({
    $settings = Get-ToolkitSettings
    $logPath = Join-Path (Get-ToolkitRoot) $settings.LogRoot
    if (-not (Test-Path $logPath)) {
        New-Item -Path $logPath -ItemType Directory -Force | Out-Null
    }
    Start-Process explorer.exe $logPath
})

$openRootButton.Add_Click({
    Start-Process explorer.exe (Get-ToolkitRoot)
})

$worker.add_DoWork({
    param($sender, $e)

    $task = $e.Argument
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
    $taskList.Enabled = $true

    if ($e.Result -and $e.Result.Success) {
        $progressBar.Value = 100
        $statusLabel.Text = "Status: Completed - $($e.Result.TaskName)"
        [System.Windows.Forms.MessageBox]::Show("Task completed: $($e.Result.TaskName)", 'Toolkit', 'OK', 'Information') | Out-Null
    }
    else {
        $progressBar.Value = 0
        $err = if ($e.Result) { $e.Result.Error } else { 'Unknown error' }
        $statusLabel.Text = 'Status: Failed'
        [System.Windows.Forms.MessageBox]::Show("Task failed: $err", 'Toolkit', 'OK', 'Error') | Out-Null
    }
})

$runButton.Add_Click({
    if ($worker.IsBusy) {
        [System.Windows.Forms.MessageBox]::Show('Another task is already running.', 'Toolkit', 'OK', 'Warning') | Out-Null
        return
    }

    $selectedTask = Get-SelectedTaskObject
    if (-not $selectedTask) {
        [System.Windows.Forms.MessageBox]::Show('Select a task first.', 'Toolkit', 'OK', 'Warning') | Out-Null
        return
    }

    if ($selectedTask.RequiresAdmin -and -not (Test-ToolkitIsAdmin)) {
        [System.Windows.Forms.MessageBox]::Show(
            "This task requires elevation. Restart PowerShell as Administrator and relaunch the toolkit.",
            'Toolkit',
            'OK',
            'Warning'
        ) | Out-Null
        return
    }

    $progressBar.Value = 0
    $statusLabel.Text = "Status: Starting $($selectedTask.Name)..."
    Add-UiLogLine -Line "Queued task: $($selectedTask.Name)"
    $runButton.Enabled = $false
    $taskList.Enabled = $false
    $worker.RunWorkerAsync($selectedTask)
})

if ($taskList.Items.Count -gt 0) {
    $taskList.Items[0].Selected = $true
}

[void]$form.ShowDialog()
