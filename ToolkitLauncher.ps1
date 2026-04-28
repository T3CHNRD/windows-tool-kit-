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
$hardwareModulePath = Join-Path $toolkitRoot 'Modules\HardwareDiagnostics\HardwareDiagnostics.psm1'
if (Test-Path -LiteralPath $hardwareModulePath) {
    Import-Module $hardwareModulePath -Force
}

[System.Windows.Forms.Application]::EnableVisualStyles()

$script:Tasks = @()
$script:CategoryOrder = @()
$script:TasksByCategory = @{}

$script:SelectedTask = $null
$script:DebloatSelections = @()
$script:DebloatInventory = @()
$script:WindowsUpdateSkipSelections = @()
$script:WindowsUpdateInventory = @()
$script:TaskListsByCategory = @{}
$script:CurrentTaskProcess = $null
$script:CurrentTaskStdoutPath = $null
$script:CurrentTaskStderrPath = $null
$script:CurrentTaskOutputLineCount = 0
$script:CurrentTaskErrorLineCount = 0
$script:CurrentTaskResult = $null
$script:CancellationRequested = $false
$script:IsDarkMode = $false

$form = New-Object System.Windows.Forms.Form
$form.Text = "T3CHNRD'S Windows Tool Kit"
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object System.Drawing.Size(1380, 860)
$form.MinimumSize = New-Object System.Drawing.Size(1220, 740)
$form.BackColor = [System.Drawing.Color]::FromArgb(244, 247, 252)
$form.Font = New-Object System.Drawing.Font('Segoe UI', 10)
$form.KeyPreview = $true

$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.AutoPopDelay = 12000
$toolTip.InitialDelay = 300
$toolTip.ReshowDelay = 150
$toolTip.ShowAlways = $true

$menuStrip = New-Object System.Windows.Forms.MenuStrip
$menuStrip.Dock = 'Fill'
$menuStrip.BackColor = [System.Drawing.Color]::White

$fileMenu = New-Object System.Windows.Forms.ToolStripMenuItem('&File')
$importScriptMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem('Import Script...')
$openLogsMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem('Open Logs Folder')
$darkModeMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem('Dark Mode')
$darkModeMenuItem.CheckOnClick = $true
$aboutMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem('About / Credits')
$exitMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem('Exit')
[void]$fileMenu.DropDownItems.Add($importScriptMenuItem)
[void]$fileMenu.DropDownItems.Add($openLogsMenuItem)
[void]$fileMenu.DropDownItems.Add($darkModeMenuItem)
[void]$fileMenu.DropDownItems.Add($aboutMenuItem)
[void]$fileMenu.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator))
[void]$fileMenu.DropDownItems.Add($exitMenuItem)

$editMenu = New-Object System.Windows.Forms.ToolStripMenuItem('&Edit')
$editScriptMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem('Edit Selected Script...')
[void]$editMenu.DropDownItems.Add($editScriptMenuItem)

[void]$menuStrip.Items.Add($fileMenu)
[void]$menuStrip.Items.Add($editMenu)
$form.MainMenuStrip = $menuStrip

$rootLayout = New-Object System.Windows.Forms.TableLayoutPanel
$rootLayout.Dock = 'Fill'
$rootLayout.ColumnCount = 1
$rootLayout.RowCount = 3
[void]$rootLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$rootLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 108)))
[void]$rootLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$form.Controls.Add($rootLayout)
[void]$rootLayout.Controls.Add($menuStrip, 0, 0)

$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Dock = 'Fill'
$headerPanel.BackColor = [System.Drawing.Color]::FromArgb(26, 45, 78)
$headerPanel.Padding = New-Object System.Windows.Forms.Padding(24, 16, 24, 16)
[void]$rootLayout.Controls.Add($headerPanel, 0, 1)

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "T3CHNRD'S Windows Tool Kit"
$titleLabel.ForeColor = [System.Drawing.Color]::White
$titleLabel.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 20, [System.Drawing.FontStyle]::Bold)
$titleLabel.AutoSize = $true
$titleLabel.Location = New-Object System.Drawing.Point(0, 8)
[void]$headerPanel.Controls.Add($titleLabel)

$subLabel = New-Object System.Windows.Forms.Label
$subLabel.Text = 'Unified repair, cleanup, update, deployment, diagnostics, and imported script runner'
$subLabel.ForeColor = [System.Drawing.Color]::FromArgb(202, 216, 240)
$subLabel.Font = New-Object System.Drawing.Font('Segoe UI', 11)
$subLabel.AutoSize = $true
$subLabel.Location = New-Object System.Drawing.Point(2, 52)
[void]$headerPanel.Controls.Add($subLabel)

$contentSplit = New-Object System.Windows.Forms.SplitContainer
$contentSplit.Dock = 'Fill'
$contentSplit.BackColor = [System.Drawing.Color]::FromArgb(232, 237, 245)
$contentSplit.FixedPanel = 'Panel1'
[void]$rootLayout.Controls.Add($contentSplit, 0, 2)

$leftLayout = New-Object System.Windows.Forms.TableLayoutPanel
$leftLayout.Dock = 'Fill'
$leftLayout.Padding = New-Object System.Windows.Forms.Padding(14, 14, 10, 14)
$leftLayout.ColumnCount = 1
$leftLayout.RowCount = 2
[void]$leftLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$leftLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$leftLayout.BackColor = [System.Drawing.Color]::FromArgb(244, 247, 252)
[void]$contentSplit.Panel1.Controls.Add($leftLayout)

$leftTitle = New-Object System.Windows.Forms.Label
$leftTitle.Text = 'Available Tools'
$leftTitle.AutoSize = $true
$leftTitle.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 14, [System.Drawing.FontStyle]::Bold)
$leftTitle.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 8)
[void]$leftLayout.Controls.Add($leftTitle, 0, 0)

$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = 'Fill'
$tabControl.Padding = New-Object System.Drawing.Point(16, 6)
$tabControl.Multiline = $true
[void]$leftLayout.Controls.Add($tabControl, 0, 1)

$rightLayout = New-Object System.Windows.Forms.TableLayoutPanel
$rightLayout.Dock = 'Fill'
$rightLayout.Padding = New-Object System.Windows.Forms.Padding(16, 14, 16, 14)
$rightLayout.ColumnCount = 1
$rightLayout.RowCount = 10
[void]$rightLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$rightLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$rightLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 140)))
[void]$rightLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$rightLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 58)))
[void]$rightLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$rightLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 52)))
[void]$rightLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$rightLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$rightLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$rightLayout.BackColor = [System.Drawing.Color]::White
[void]$contentSplit.Panel2.Controls.Add($rightLayout)

$selectedToolLabel = New-Object System.Windows.Forms.Label
$selectedToolLabel.Text = 'Tool Details'
$selectedToolLabel.AutoSize = $true
$selectedToolLabel.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 15, [System.Drawing.FontStyle]::Bold)
$selectedToolLabel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 4)
[void]$rightLayout.Controls.Add($selectedToolLabel, 0, 0)

$descriptionHintLabel = New-Object System.Windows.Forms.Label
$descriptionHintLabel.Text = 'Select a tool from the list on the left. Double-click it or use Run Selected Tool.'
$descriptionHintLabel.AutoSize = $true
$descriptionHintLabel.ForeColor = [System.Drawing.Color]::FromArgb(84, 96, 118)
$descriptionHintLabel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 8)
[void]$rightLayout.Controls.Add($descriptionHintLabel, 0, 1)

$descriptionBox = New-Object System.Windows.Forms.TextBox
$descriptionBox.Multiline = $true
$descriptionBox.ReadOnly = $true
$descriptionBox.Dock = 'Fill'
$descriptionBox.BorderStyle = 'FixedSingle'
$descriptionBox.BackColor = [System.Drawing.Color]::FromArgb(249, 251, 255)
$descriptionBox.ScrollBars = 'Vertical'
$descriptionBox.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)
[void]$rightLayout.Controls.Add($descriptionBox, 0, 2)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = 'Status: Ready'
$statusLabel.AutoSize = $true
$statusLabel.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 11)
$statusLabel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 6)
[void]$rightLayout.Controls.Add($statusLabel, 0, 3)

$progressPanel = New-Object System.Windows.Forms.Panel
$progressPanel.Dock = 'Fill'
$progressPanel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 12)

$script:ProgressValue = 0
$script:ProgressAnimationFrame = 0

$progressRoadPanel = New-Object System.Windows.Forms.Panel
$progressRoadPanel.Dock = 'Fill'
$progressRoadPanel.Height = 46
$progressRoadPanel.BackColor = [System.Drawing.Color]::FromArgb(249, 251, 255)
$progressRoadPanel.Add_Paint({
    param($sender, $paintEvent)

    $graphics = $paintEvent.Graphics
    $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $bounds = $sender.ClientRectangle
    if ($bounds.Width -lt 80 -or $bounds.Height -lt 30) {
        return
    }

    $trackLeft = 12
    $trackTop = [Math]::Max(14, [Math]::Floor($bounds.Height / 2) - 5)
    $trackWidth = [Math]::Max(20, $bounds.Width - 74)
    $trackHeight = 10
    $percent = [Math]::Max(0, [Math]::Min(100, [int]$script:ProgressValue))
    $fillWidth = [Math]::Floor($trackWidth * ($percent / 100))
    $bikeX = $trackLeft + [Math]::Max(0, $fillWidth) - 10
    if ($bikeX -lt $trackLeft) { $bikeX = $trackLeft }
    if ($bikeX -gt ($trackLeft + $trackWidth - 24)) { $bikeX = $trackLeft + $trackWidth - 24 }
    $bikeY = $trackTop - 17 + (($script:ProgressAnimationFrame % 2) * 1)

    $trackBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(218, 225, 235))
    $fillBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(44, 157, 87))
    $stripePen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(255, 255, 255), 2)
    $darkPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(42, 52, 66), 2)
    $bikeBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(38, 125, 205))
    $riderBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(248, 184, 65))
    $wheelBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(20, 27, 36))
    $flagBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(230, 57, 70))
    $textBrush = New-Object System.Drawing.SolidBrush($sender.ForeColor)
    $font = New-Object System.Drawing.Font('Segoe UI Semibold', 8)

    try {
        $graphics.FillRectangle($trackBrush, $trackLeft, $trackTop, $trackWidth, $trackHeight)
        if ($fillWidth -gt 0) {
            $graphics.FillRectangle($fillBrush, $trackLeft, $trackTop, $fillWidth, $trackHeight)
        }

        for ($stripeX = $trackLeft + 18; $stripeX -lt ($trackLeft + $trackWidth); $stripeX += 36) {
            $graphics.DrawLine($stripePen, $stripeX, $trackTop + 5, $stripeX + 14, $trackTop + 5)
        }

        $flagX = $trackLeft + $trackWidth + 10
        $graphics.DrawLine($darkPen, $flagX, $trackTop - 15, $flagX, $trackTop + 16)
        $graphics.FillPolygon($flagBrush, @(
            (New-Object System.Drawing.Point($flagX, $trackTop - 15)),
            (New-Object System.Drawing.Point($flagX + 22, $trackTop - 9)),
            (New-Object System.Drawing.Point($flagX, $trackTop - 3))
        ))

        $graphics.FillEllipse($wheelBrush, $bikeX, $bikeY + 20, 10, 10)
        $graphics.FillEllipse($wheelBrush, $bikeX + 24, $bikeY + 20, 10, 10)
        $graphics.DrawLine($darkPen, $bikeX + 5, $bikeY + 24, $bikeX + 18, $bikeY + 12)
        $graphics.DrawLine($darkPen, $bikeX + 18, $bikeY + 12, $bikeX + 29, $bikeY + 24)
        $graphics.DrawLine($darkPen, $bikeX + 10, $bikeY + 24, $bikeX + 29, $bikeY + 24)
        $graphics.FillRectangle($bikeBrush, $bikeX + 13, $bikeY + 14, 14, 7)
        $graphics.FillEllipse($riderBrush, $bikeX + 15, $bikeY + 3, 9, 9)
        $graphics.DrawLine($darkPen, $bikeX + 19, $bikeY + 12, $bikeX + 15, $bikeY + 21)
        $graphics.DrawLine($darkPen, $bikeX + 20, $bikeY + 12, $bikeX + 27, $bikeY + 19)

        $graphics.DrawString(('{0}%' -f $percent), $font, $textBrush, ($trackLeft + $trackWidth - 34), 0)
    }
    finally {
        foreach ($resource in @($trackBrush, $fillBrush, $stripePen, $darkPen, $bikeBrush, $riderBrush, $wheelBrush, $flagBrush, $textBrush, $font)) {
            if ($resource) { $resource.Dispose() }
        }
    }
})
[void]$progressPanel.Controls.Add($progressRoadPanel)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Dock = 'Bottom'
$progressBar.Height = 8
$progressBar.Minimum = 0
$progressBar.Maximum = 100
$progressBar.Value = 0
$progressBar.Style = 'Continuous'
[void]$progressPanel.Controls.Add($progressBar)
[void]$rightLayout.Controls.Add($progressPanel, 0, 4)

$metaLabel = New-Object System.Windows.Forms.Label
$metaLabel.Text = 'Requires Admin: Yes | Input: None required'
$metaLabel.AutoSize = $true
$metaLabel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)
[void]$rightLayout.Controls.Add($metaLabel, 0, 5)

$buttonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$buttonPanel.Dock = 'Fill'
$buttonPanel.FlowDirection = 'LeftToRight'
$buttonPanel.WrapContents = $true
$buttonPanel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 12)
[void]$rightLayout.Controls.Add($buttonPanel, 0, 6)

$runButton = New-Object System.Windows.Forms.Button
$runButton.Text = 'Run Selected Tool'
$runButton.Size = New-Object System.Drawing.Size(170, 38)
$runButton.BackColor = [System.Drawing.Color]::FromArgb(47, 108, 188)
$runButton.ForeColor = [System.Drawing.Color]::White
$runButton.FlatStyle = 'Flat'
$toolTip.SetToolTip($runButton, 'Runs the tool currently selected in the active tab.')
[void]$buttonPanel.Controls.Add($runButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Text = 'Cancel Running Task'
$cancelButton.Size = New-Object System.Drawing.Size(170, 38)
$cancelButton.FlatStyle = 'Flat'
$cancelButton.Enabled = $false
$toolTip.SetToolTip($cancelButton, 'Stops the currently running background task.')
[void]$buttonPanel.Controls.Add($cancelButton)

$chooseAppsButton = New-Object System.Windows.Forms.Button
$chooseAppsButton.Text = 'Choose Apps'
$chooseAppsButton.Size = New-Object System.Drawing.Size(126, 38)
$chooseAppsButton.FlatStyle = 'Flat'
$chooseAppsButton.Visible = $false
$toolTip.SetToolTip($chooseAppsButton, 'Choose applications to include in the debloat review list.')
[void]$buttonPanel.Controls.Add($chooseAppsButton)

$openLogsButton = New-Object System.Windows.Forms.Button
$openLogsButton.Text = 'Open Logs Folder'
$openLogsButton.Size = New-Object System.Drawing.Size(150, 38)
$openLogsButton.FlatStyle = 'Flat'
[void]$buttonPanel.Controls.Add($openLogsButton)

$logTitle = New-Object System.Windows.Forms.Label
$logTitle.Text = 'Execution Log'
$logTitle.AutoSize = $true
$logTitle.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 11, [System.Drawing.FontStyle]::Bold)
$logTitle.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 6)
[void]$rightLayout.Controls.Add($logTitle, 0, 7)

$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Multiline = $true
$logBox.ReadOnly = $true
$logBox.ScrollBars = 'Vertical'
$logBox.Dock = 'Fill'
$logBox.BorderStyle = 'FixedSingle'
$logBox.BackColor = [System.Drawing.Color]::FromArgb(247, 249, 252)
[void]$rightLayout.Controls.Add($logBox, 0, 8)

$footerLabel = New-Object System.Windows.Forms.Label
$footerLabel.Text = 'Tip: single-click selects, double-click runs, and F5 reruns the selected tool.'
$footerLabel.AutoSize = $true
$footerLabel.ForeColor = [System.Drawing.Color]::FromArgb(98, 108, 124)
$footerLabel.Margin = New-Object System.Windows.Forms.Padding(0, 8, 0, 0)
[void]$rightLayout.Controls.Add($footerLabel, 0, 9)

$taskMonitorTimer = New-Object System.Windows.Forms.Timer
$taskMonitorTimer.Interval = 450

$progressAnimationTimer = New-Object System.Windows.Forms.Timer
$progressAnimationTimer.Interval = 120
$progressAnimationTimer.Add_Tick({
    if (Test-TaskIsRunning) {
        $script:ProgressAnimationFrame += 1
        $progressRoadPanel.Invalidate()
    }
})
$progressAnimationTimer.Start()

function Add-UiLogLine {
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Line
    )

    if ($null -eq $Line -or [string]::IsNullOrWhiteSpace($Line)) {
        return
    }

    $timestamped = "[{0:HH:mm:ss}] {1}" -f (Get-Date), $Line
    $logBox.AppendText($timestamped + [Environment]::NewLine)
}

function Set-ToolkitProgress {
    param(
        [Parameter(Mandatory = $true)][int]$Value
    )

    $bounded = [Math]::Max(0, [Math]::Min(100, $Value))
    if ($progressBar.Style -ne 'Continuous') {
        $progressBar.Style = 'Continuous'
    }
    $progressBar.Value = $bounded

    $script:ProgressValue = $bounded
    $script:ProgressAnimationFrame += 1
    $progressRoadPanel.Invalidate()
}

function Get-ToolkitThemePalette {
    if ($script:IsDarkMode) {
        return @{
            FormBack = [System.Drawing.Color]::FromArgb(17, 24, 39)
            PanelBack = [System.Drawing.Color]::FromArgb(24, 34, 52)
            ContentBack = [System.Drawing.Color]::FromArgb(14, 20, 32)
            Surface = [System.Drawing.Color]::FromArgb(31, 42, 62)
            SurfaceAlt = [System.Drawing.Color]::FromArgb(38, 51, 74)
            Text = [System.Drawing.Color]::FromArgb(235, 241, 250)
            MutedText = [System.Drawing.Color]::FromArgb(169, 183, 204)
            Border = [System.Drawing.Color]::FromArgb(75, 91, 116)
            Header = [System.Drawing.Color]::FromArgb(11, 22, 42)
            HeaderSub = [System.Drawing.Color]::FromArgb(174, 194, 224)
            Primary = [System.Drawing.Color]::FromArgb(62, 125, 207)
            ButtonBack = [System.Drawing.Color]::FromArgb(42, 56, 80)
        }
    }

    return @{
        FormBack = [System.Drawing.Color]::FromArgb(244, 247, 252)
        PanelBack = [System.Drawing.Color]::FromArgb(244, 247, 252)
        ContentBack = [System.Drawing.Color]::FromArgb(232, 237, 245)
        Surface = [System.Drawing.Color]::White
        SurfaceAlt = [System.Drawing.Color]::FromArgb(247, 249, 252)
        Text = [System.Drawing.Color]::FromArgb(15, 23, 42)
        MutedText = [System.Drawing.Color]::FromArgb(84, 96, 118)
        Border = [System.Drawing.Color]::FromArgb(190, 198, 211)
        Header = [System.Drawing.Color]::FromArgb(26, 45, 78)
        HeaderSub = [System.Drawing.Color]::FromArgb(202, 216, 240)
        Primary = [System.Drawing.Color]::FromArgb(47, 108, 188)
        ButtonBack = [System.Drawing.Color]::White
    }
}

function Set-ToolkitControlTheme {
    param(
        [Parameter(Mandatory = $true)]$Control,
        [Parameter(Mandatory = $true)][hashtable]$Palette
    )

    if ($Control -is [System.Windows.Forms.ProgressBar]) {
        return
    }

    if ($Control -is [System.Windows.Forms.ListView]) {
        $Control.BackColor = [System.Drawing.Color]::White
        $Control.ForeColor = [System.Drawing.Color]::Black
    }
    elseif ($Control -is [System.Windows.Forms.TextBox] -or $Control -is [System.Windows.Forms.CheckedListBox]) {
        $Control.BackColor = $Palette.SurfaceAlt
        $Control.ForeColor = $Palette.Text
    }
    elseif ($Control -is [System.Windows.Forms.Button]) {
        if ($Control -eq $runButton) {
            $Control.BackColor = $Palette.Primary
            $Control.ForeColor = [System.Drawing.Color]::White
        }
        else {
            $Control.BackColor = $Palette.ButtonBack
            $Control.ForeColor = $Palette.Text
        }
    }
    elseif ($Control -is [System.Windows.Forms.Label]) {
        $Control.ForeColor = $Palette.Text
    }
    elseif ($Control -is [System.Windows.Forms.TabPage] -or $Control -is [System.Windows.Forms.Panel] -or $Control -is [System.Windows.Forms.TableLayoutPanel] -or $Control -is [System.Windows.Forms.FlowLayoutPanel]) {
        $Control.BackColor = $Palette.PanelBack
        $Control.ForeColor = $Palette.Text
    }

    foreach ($child in $Control.Controls) {
        Set-ToolkitControlTheme -Control $child -Palette $Palette
    }
}

function Set-ToolkitMenuTheme {
    param(
        [Parameter(Mandatory = $true)][hashtable]$Palette
    )

    $menuStrip.BackColor = $Palette.Surface
    $menuStrip.ForeColor = $Palette.Text
    foreach ($menuItem in @($fileMenu, $editMenu, $importScriptMenuItem, $openLogsMenuItem, $darkModeMenuItem, $aboutMenuItem, $exitMenuItem, $editScriptMenuItem)) {
        $menuItem.BackColor = $Palette.Surface
        $menuItem.ForeColor = $Palette.Text
        if ($menuItem.DropDown) {
            $menuItem.DropDown.BackColor = $Palette.Surface
            $menuItem.DropDown.ForeColor = $Palette.Text
        }
    }
}

function Apply-ToolkitTheme {
    $palette = Get-ToolkitThemePalette

    $form.BackColor = $palette.FormBack
    $contentSplit.BackColor = $palette.ContentBack
    $contentSplit.Panel1.BackColor = $palette.PanelBack
    $contentSplit.Panel2.BackColor = $palette.Surface
    $headerPanel.BackColor = $palette.Header
    $titleLabel.ForeColor = [System.Drawing.Color]::White
    $subLabel.ForeColor = $palette.HeaderSub
    $leftLayout.BackColor = $palette.PanelBack
    $rightLayout.BackColor = $palette.Surface

    Set-ToolkitMenuTheme -Palette $palette
    Set-ToolkitControlTheme -Control $leftLayout -Palette $palette
    Set-ToolkitControlTheme -Control $rightLayout -Palette $palette

    $descriptionHintLabel.ForeColor = $palette.MutedText
    $footerLabel.ForeColor = $palette.MutedText
    $descriptionBox.BackColor = $palette.SurfaceAlt
    $descriptionBox.ForeColor = $palette.Text
    $logBox.BackColor = $palette.SurfaceAlt
    $logBox.ForeColor = $palette.Text
    $progressPanel.BackColor = $palette.Surface
    $progressRoadPanel.BackColor = $palette.Surface
    $progressRoadPanel.ForeColor = $palette.MutedText
    $progressRoadPanel.Invalidate()

    foreach ($tabPage in $tabControl.TabPages) {
        $tabPage.BackColor = [System.Drawing.Color]::White
        $tabPage.ForeColor = [System.Drawing.Color]::Black
    }

    foreach ($category in $script:TaskListsByCategory.Keys) {
        $taskList = $script:TaskListsByCategory[$category]
        $taskList.BackColor = [System.Drawing.Color]::White
        $taskList.ForeColor = [System.Drawing.Color]::Black
    }

    $runButton.BackColor = $palette.Primary
    $runButton.ForeColor = [System.Drawing.Color]::White
}

function Test-TaskIsRunning {
    return ($script:CurrentTaskProcess -and -not $script:CurrentTaskProcess.HasExited)
}

function Stop-RunningTaskProcessTree {
    if (-not $script:CurrentTaskProcess) {
        return
    }

    try {
        & taskkill.exe /PID $script:CurrentTaskProcess.Id /T /F | Out-Null
    }
    catch {
    }
}

function Get-CategoryTabLabel {
    param(
        [Parameter(Mandatory = $true)][string]$Category
    )

    switch ($Category) {
        'Applications' { 'Apps' }
        'Cleanup' { 'Cleanup' }
        'Deployment' { 'Deploy' }
        'Hardware Diagnostics' { 'Hardware' }
        'Misc/Utility' { 'Misc' }
        'Network Tools' { 'Network' }
        'OneNote & Documents' { 'OneNote' }
        'Repair' { 'Repair' }
        'Storage / Setup' { 'Storage' }
        'Update Tools' { 'Updates' }
        default { $Category }
    }
}

function Refresh-TaskCatalogState {
    $script:Tasks = @(Get-ToolkitTaskCatalog | Sort-Object Category, Name)
    $script:CategoryOrder = @(Get-ToolkitCategoryOrder | Where-Object { $script:Tasks.Category -contains $_ })
    $script:TasksByCategory = @{}

    foreach ($category in $script:CategoryOrder) {
        $script:TasksByCategory[$category] = @($script:Tasks | Where-Object { $_.Category -eq $category })
    }
}

function Get-TaskSourcePath {
    param(
        [Parameter(Mandatory = $true)]$Task
    )

    if ($Task.Id -like 'Legacy.*') {
        $scriptName = ($Task.Name -replace '^Legacy: ', '') + '.ps1'
        return (Join-Path $toolkitRoot "LegacyScripts\$scriptName")
    }

    $map = @{
        'Apps.DebloatHelper'            = 'Scripts\Tasks\Invoke-DebloatInventory.ps1'
        'Cleanup.TempFiles'             = 'Scripts\Tasks\Invoke-ClearTempJunk.ps1'
        'Cleanup.FreeSpace'             = 'Scripts\Tasks\Invoke-FreeCDriveSpace.ps1'
        'Disk.Monitor'                  = 'Scripts\Tasks\Invoke-DiskSpaceMonitor.ps1'
        'Integration.MediaCreationTool' = 'Integrations\Invoke-MediaCreationWorkflow.ps1'
        'Integration.Microsoft365'      = 'Integrations\Invoke-InstallMicrosoft365.ps1'
        'Network.Diagnostics'           = 'Scripts\Tasks\Invoke-NetworkMaintenance.ps1'
        'Network.DhcpRenew'             = 'Scripts\Tasks\Invoke-DhcpRenew.ps1'
        'Network.ResetStack'            = 'Scripts\Tasks\Invoke-ResetNetworkStack.ps1'
        'Repair.WindowsHealth'          = 'Scripts\Tasks\Invoke-WindowsRepairChecks.ps1'
        'Update.AllApps'                = 'Scripts\Tasks\Invoke-UpdateAllApps.ps1'
        'Update.VendorBIOS'             = 'Scripts\Tasks\Invoke-BiosUpdate.ps1'
        'Update.VendorDrivers'          = 'Scripts\Tasks\Invoke-DriverUpdate.ps1'
        'Update.VendorFirmware'         = 'Scripts\Tasks\Invoke-FirmwareUpdate.ps1'
        'Update.WindowsOS'              = 'Scripts\Tasks\Invoke-WindowsUpdateTool.ps1'
        'Hardware.MouseKeyboardTest'    = 'Modules\HardwareDiagnostics\HardwareDiagnostics.psm1'
        'Hardware.MonitorPixelTest'     = 'Modules\HardwareDiagnostics\HardwareDiagnostics.psm1'
        'Storage.CloneGuide'            = 'Modules\StorageTools\StorageTools.psm1'
        'Storage.NewComputerSetup'      = 'Modules\StorageTools\StorageTools.psm1'
    }

    if ($map.ContainsKey($Task.Id)) {
        return (Join-Path $toolkitRoot $map[$Task.Id])
    }

    return $null
}

function Show-ScriptEditor {
    param(
        [Parameter(Mandatory = $true)]$Task
    )

    $sourcePath = Get-TaskSourcePath -Task $Task
    if (-not $sourcePath -or -not (Test-Path -LiteralPath $sourcePath)) {
        [System.Windows.Forms.MessageBox]::Show(
            "No editable source file is mapped for $($Task.Name) yet.",
            'Toolkit',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        return
    }

    $editorForm = New-Object System.Windows.Forms.Form
    $editorForm.Text = "Edit Script - $($Task.Name)"
    $editorForm.StartPosition = 'CenterParent'
    $editorForm.Size = New-Object System.Drawing.Size(980, 760)
    $editorForm.MinimumSize = New-Object System.Drawing.Size(760, 560)
    $editorForm.Font = $form.Font

    $editorLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $editorLayout.Dock = 'Fill'
    $editorLayout.Padding = New-Object System.Windows.Forms.Padding(12)
    $editorLayout.ColumnCount = 1
    $editorLayout.RowCount = 4
    [void]$editorLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$editorLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$editorLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$editorLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$editorForm.Controls.Add($editorLayout)

    $pathLabel = New-Object System.Windows.Forms.Label
    $pathLabel.Text = $sourcePath
    $pathLabel.AutoSize = $true
    $pathLabel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 6)
    [void]$editorLayout.Controls.Add($pathLabel, 0, 0)

    $hintLabel = New-Object System.Windows.Forms.Label
    $hintLabel.Text = 'Edit the script here and save changes directly back into the toolkit.'
    $hintLabel.AutoSize = $true
    $hintLabel.ForeColor = [System.Drawing.Color]::FromArgb(84, 96, 118)
    $hintLabel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)
    [void]$editorLayout.Controls.Add($hintLabel, 0, 1)

    $scriptEditorBox = New-Object System.Windows.Forms.TextBox
    $scriptEditorBox.Multiline = $true
    $scriptEditorBox.AcceptsReturn = $true
    $scriptEditorBox.AcceptsTab = $true
    $scriptEditorBox.ScrollBars = 'Both'
    $scriptEditorBox.WordWrap = $false
    $scriptEditorBox.Dock = 'Fill'
    $scriptEditorBox.Font = New-Object System.Drawing.Font('Consolas', 10)
    $scriptEditorBox.Text = [IO.File]::ReadAllText($sourcePath)
    [void]$editorLayout.Controls.Add($scriptEditorBox, 0, 2)

    $editorButtons = New-Object System.Windows.Forms.FlowLayoutPanel
    $editorButtons.Dock = 'Fill'
    $editorButtons.FlowDirection = 'RightToLeft'
    $editorButtons.WrapContents = $false
    $editorButtons.Margin = New-Object System.Windows.Forms.Padding(0, 10, 0, 0)
    [void]$editorLayout.Controls.Add($editorButtons, 0, 3)

    $closeEditorButton = New-Object System.Windows.Forms.Button
    $closeEditorButton.Text = 'Close'
    $closeEditorButton.Size = New-Object System.Drawing.Size(120, 36)
    [void]$editorButtons.Controls.Add($closeEditorButton)

    $saveEditorButton = New-Object System.Windows.Forms.Button
    $saveEditorButton.Text = 'Save'
    $saveEditorButton.Size = New-Object System.Drawing.Size(120, 36)
    [void]$editorButtons.Controls.Add($saveEditorButton)

    $reloadEditorButton = New-Object System.Windows.Forms.Button
    $reloadEditorButton.Text = 'Reload'
    $reloadEditorButton.Size = New-Object System.Drawing.Size(120, 36)
    [void]$editorButtons.Controls.Add($reloadEditorButton)

    $openContainingFolderButton = New-Object System.Windows.Forms.Button
    $openContainingFolderButton.Text = 'Open Folder'
    $openContainingFolderButton.Size = New-Object System.Drawing.Size(120, 36)
    [void]$editorButtons.Controls.Add($openContainingFolderButton)

    $reloadEditorButton.Add_Click({
        $scriptEditorBox.Text = [IO.File]::ReadAllText($sourcePath)
    })

    $openContainingFolderButton.Add_Click({
        Start-Process explorer.exe "/select,`"$sourcePath`""
    })

    $saveEditorButton.Add_Click({
        $tempSavePath = Join-Path $env:TEMP ("toolkit-parse-{0}.ps1" -f ([guid]::NewGuid().ToString('N')))
        try {
            Set-Content -LiteralPath $tempSavePath -Value $scriptEditorBox.Text -Encoding UTF8
            $null = $parseErrors = $parseTokens = $parseAst = $null
            [System.Management.Automation.Language.Parser]::ParseFile($tempSavePath, [ref]$parseTokens, [ref]$parseErrors) | Out-Null
            if ($parseErrors.Count -gt 0) {
                $message = ($parseErrors | ForEach-Object { $_.Message } | Select-Object -First 5) -join [Environment]::NewLine
                [System.Windows.Forms.MessageBox]::Show(
                    "Save blocked because the script has parse errors:`r`n`r`n$message",
                    'Toolkit',
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                ) | Out-Null
                return
            }

            Set-Content -LiteralPath $sourcePath -Value $scriptEditorBox.Text -Encoding UTF8
            Add-UiLogLine -Line "Saved script changes: $sourcePath"
            [System.Windows.Forms.MessageBox]::Show(
                "Saved changes to:`r`n$sourcePath",
                'Toolkit',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
        }
        finally {
            Remove-Item -LiteralPath $tempSavePath -Force -ErrorAction SilentlyContinue
        }
    })

    $closeEditorButton.Add_Click({
        $editorForm.Close()
    })

    [void]$editorForm.ShowDialog($form)
}

function Open-ToolkitLogsFolder {
    $settings = Get-ToolkitSettings
    $logPath = Join-Path (Get-ToolkitRoot) $settings.LogRoot
    if (-not (Test-Path -LiteralPath $logPath)) {
        New-Item -Path $logPath -ItemType Directory -Force | Out-Null
    }
    Start-Process explorer.exe $logPath
}

function Show-ToolkitAbout {
    $aboutText = @"
T3CHNRD'S Windows Tool Kit

Created for T3CHNRD / GitHub: https://github.com/T3CHNRD
Project repository: https://github.com/T3CHNRD/windows-tool-kit-

External project credits:
- mallockey/Install-Microsoft365
  https://github.com/mallockey/Install-Microsoft365
  Used as the legitimate Microsoft 365 / Office Deployment Tool workflow integration.

- LottieFiles loading bar animations
  https://lottiefiles.com/free-animations/loading-bar
  Used as visual inspiration for the toolkit's motorcycle progress indicator. No third-party animation file is bundled.

Notes:
- Microsoft 365 installation still requires valid Microsoft licensing.
- The toolkit avoids activation-bypass or piracy workflows.
"@

    [System.Windows.Forms.MessageBox]::Show(
        $aboutText,
        "About T3CHNRD'S Windows Tool Kit",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
}

function Import-ScriptIntoToolkit {
    if (Test-TaskIsRunning) {
        [System.Windows.Forms.MessageBox]::Show(
            'Wait for the current task to finish before importing another script.',
            'Toolkit',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        return
    }

    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = 'PowerShell scripts (*.ps1)|*.ps1|All files (*.*)|*.*'
    $dialog.Multiselect = $false
    $dialog.Title = 'Import PowerShell Script Into Toolkit'

    if ($dialog.ShowDialog($form) -ne [System.Windows.Forms.DialogResult]::OK) {
        return
    }

    $destinationDir = Join-Path $toolkitRoot 'LegacyScripts'
    if (-not (Test-Path -LiteralPath $destinationDir)) {
        New-Item -Path $destinationDir -ItemType Directory -Force | Out-Null
    }

    $destinationPath = Join-Path $destinationDir ([IO.Path]::GetFileName($dialog.FileName))
    if ((Test-Path -LiteralPath $destinationPath) -and ((Resolve-Path -LiteralPath $dialog.FileName).Path -ne (Resolve-Path -LiteralPath $destinationPath).Path)) {
        $overwrite = [System.Windows.Forms.MessageBox]::Show(
            "A script with this name already exists in the toolkit.`r`n`r`nOverwrite it?",
            'Toolkit',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($overwrite -ne [System.Windows.Forms.DialogResult]::Yes) {
            return
        }
    }

    Copy-Item -LiteralPath $dialog.FileName -Destination $destinationPath -Force
    Add-UiLogLine -Line "Imported script into toolkit: $destinationPath"
    Refresh-TaskTabs
    [System.Windows.Forms.MessageBox]::Show(
        "Imported script:`r`n$destinationPath",
        'Toolkit',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
}

function Edit-SelectedScript {
    $selectedTask = Get-SelectedTaskFromActiveTab
    if (-not $selectedTask) {
        [System.Windows.Forms.MessageBox]::Show(
            'Select a tool first, then open its script in the editor.',
            'Toolkit',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        return
    }

    Show-ScriptEditor -Task $selectedTask
}

function Refresh-TaskTabs {
    $selectedTaskId = if ($script:SelectedTask) { $script:SelectedTask.Id } else { $null }

    Refresh-TaskCatalogState
    $tabControl.TabPages.Clear()
    $script:TaskListsByCategory = @{}

    foreach ($category in $script:CategoryOrder) {
        $tabPage = New-Object System.Windows.Forms.TabPage
        $tabPage.Text = Get-CategoryTabLabel -Category $category
        $tabPage.Tag = $category
        $tabPage.BackColor = [System.Drawing.Color]::White
        [void]$tabControl.TabPages.Add($tabPage)

        $taskList = New-Object System.Windows.Forms.ListView
        $taskList.Dock = 'Fill'
        $taskList.View = 'Details'
        $taskList.FullRowSelect = $true
        $taskList.HideSelection = $false
        $taskList.MultiSelect = $false
        $taskList.ShowItemToolTips = $true
        $taskList.GridLines = $true
        $taskList.Activation = 'OneClick'
        [void]$taskList.Columns.Add('Tool', 320)
        [void]$taskList.Columns.Add('Admin', 70)
        [void]$tabPage.Controls.Add($taskList)
        $script:TaskListsByCategory[$category] = $taskList

        foreach ($task in $script:TasksByCategory[$category]) {
            $adminText = if ($task.RequiresAdmin) { 'Yes' } else { 'No' }
            $item = New-Object System.Windows.Forms.ListViewItem($task.Name)
            [void]$item.SubItems.Add($adminText)
            $item.Tag = $task.Id
            $item.ToolTipText = $task.Description
            [void]$taskList.Items.Add($item)
        }

        $taskList.Add_SelectedIndexChanged({
            $selectedTask = Get-SelectedTaskFromActiveTab
            if ($selectedTask) {
                Update-SelectedTaskDisplay -Task $selectedTask
            }
        })

        $taskList.Add_DoubleClick({
            $selectedTask = Get-SelectedTaskFromActiveTab
            if ($selectedTask) {
                Start-ToolkitTask -Task $selectedTask
            }
        })
    }

    if ($tabControl.TabPages.Count -eq 0) {
        Update-SelectedTaskDisplay -Task $null
        return
    }

    $selectedPageIndex = 0
    if ($selectedTaskId) {
        $taskToRestore = $script:Tasks | Where-Object { $_.Id -eq $selectedTaskId } | Select-Object -First 1
        if ($taskToRestore) {
            $selectedPage = $tabControl.TabPages | Where-Object { $_.Tag -eq $taskToRestore.Category } | Select-Object -First 1
            if ($selectedPage) {
                $selectedPageIndex = $tabControl.TabPages.IndexOf($selectedPage)
            }
        }
    }

    $tabControl.SelectedIndex = [Math]::Max(0, $selectedPageIndex)
    Select-FirstTaskInTab

    if ($selectedTaskId) {
        foreach ($taskList in $script:TaskListsByCategory.Values) {
            foreach ($item in $taskList.Items) {
                if ($item.Tag -eq $selectedTaskId) {
                    $item.Selected = $true
                    $item.Focused = $true
                    break
                }
            }
        }
    }

    $selectedTask = Get-SelectedTaskFromActiveTab
    if ($selectedTask) {
        Update-SelectedTaskDisplay -Task $selectedTask
    }
    else {
        Update-SelectedTaskDisplay -Task $null
    }

    Apply-ToolkitTheme
}

Refresh-TaskCatalogState

function Get-ToolkitLogDirectory {
    $settings = Get-ToolkitSettings
    $logPath = Join-Path (Get-ToolkitRoot) $settings.LogRoot
    if (-not (Test-Path -LiteralPath $logPath)) {
        New-Item -Path $logPath -ItemType Directory -Force | Out-Null
    }

    return $logPath
}

function Get-NewFileLines {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][ref]$ProcessedCount
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return @()
    }

    $allLines = @(Get-Content -LiteralPath $Path -ErrorAction SilentlyContinue)
    if ($ProcessedCount.Value -ge $allLines.Count) {
        return @()
    }

    $newLines = $allLines[$ProcessedCount.Value..($allLines.Count - 1)]
    $ProcessedCount.Value = $allLines.Count
    return @($newLines)
}

function Process-TaskOutputLine {
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Line,
        [switch]$IsErrorStream
    )

    if ($null -eq $Line -or $Line.Length -eq 0) {
        return
    }

    if ($Line -like '__TTK__|*') {
        $parts = $Line -split '\|', 5
        $kind = if ($parts.Count -gt 1) { $parts[1] } else { '' }

        switch ($kind) {
            'PROGRESS' {
                $percent = 0
                if ($parts.Count -gt 2) {
                    [void][int]::TryParse($parts[2], [ref]$percent)
                }
                $status = if ($parts.Count -gt 3) { $parts[3] } else { 'Running task' }
                Set-ToolkitProgress -Value $percent
                $statusLabel.Text = "Status: $status"
                return
            }
            'STATUS' {
                $status = if ($parts.Count -gt 2) { $parts[2] } else { 'Running task' }
                $statusLabel.Text = "Status: $status"
                return
            }
            'LOG' {
                if ($parts.Count -gt 2) {
                    Add-UiLogLine -Line $parts[2]
                }
                return
            }
            'RESULT' {
                $script:CurrentTaskResult = [pscustomobject]@{
                    Outcome = if ($parts.Count -gt 2) { $parts[2] } else { 'UNKNOWN' }
                    TaskName = if ($parts.Count -gt 3) { $parts[3] } else { '' }
                    Message = if ($parts.Count -gt 4) { $parts[4] } else { '' }
                }
                return
            }
        }
    }

    if ($Line) {
        $prefix = if ($IsErrorStream) { 'STDERR: ' } else { '' }
        Add-UiLogLine -Line ($prefix + $Line)
    }
}

function Complete-RunningTask {
    if (-not $script:CurrentTaskProcess) {
        return
    }

    try {
        $script:CurrentTaskProcess.WaitForExit()
        $script:CurrentTaskProcess.Refresh()
    }
    catch {
    }

    $outputCount = $script:CurrentTaskOutputLineCount
    foreach ($line in (Get-NewFileLines -Path $script:CurrentTaskStdoutPath -ProcessedCount ([ref]$outputCount))) {
        Process-TaskOutputLine -Line ([string]$line)
    }
    $script:CurrentTaskOutputLineCount = $outputCount

    $errorCount = $script:CurrentTaskErrorLineCount
    foreach ($line in (Get-NewFileLines -Path $script:CurrentTaskStderrPath -ProcessedCount ([ref]$errorCount))) {
        Process-TaskOutputLine -Line ([string]$line) -IsErrorStream
    }
    $script:CurrentTaskErrorLineCount = $errorCount

    $taskName = if ($script:CurrentTaskResult -and $script:CurrentTaskResult.TaskName) {
        $script:CurrentTaskResult.TaskName
    }
    elseif ($script:SelectedTask) {
        $script:SelectedTask.Name
    }
    else {
        'Task'
    }

    $exitCode = 0
    $hasExitCode = $false
    try {
        $exitCode = [int]$script:CurrentTaskProcess.ExitCode
        $hasExitCode = $true
    }
    catch {
    }

    $failedMessage = $null
    if ($script:CurrentTaskResult -and $script:CurrentTaskResult.Outcome -eq 'FAIL') {
        $failedMessage = if ($script:CurrentTaskResult.Message) { $script:CurrentTaskResult.Message } else { 'Unknown error' }
    }
    elseif ($script:CurrentTaskResult -and $script:CurrentTaskResult.Outcome -eq 'SUCCESS') {
        $failedMessage = $null
    }
    elseif ($hasExitCode -and $exitCode -ne 0) {
        $failedMessage = "Process exited with code $exitCode."
    }

    $taskMonitorTimer.Stop()
    Set-InteractiveState -Enabled $true
    Select-FirstTaskInTab
    $progressBar.Style = 'Continuous'

    if ($script:CancellationRequested) {
        Set-ToolkitProgress -Value 0
        $statusLabel.Text = "Status: Cancelled - $taskName"
        [System.Windows.Forms.MessageBox]::Show(
            "Task cancelled: $taskName",
            'Toolkit',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }
    elseif ($failedMessage) {
        Set-ToolkitProgress -Value 0
        $statusLabel.Text = "Status: Failed - $failedMessage"
        [System.Windows.Forms.MessageBox]::Show(
            "Task failed: $failedMessage",
            'Toolkit',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
    else {
        Set-ToolkitProgress -Value 100
        $statusLabel.Text = "Status: Completed - $taskName"
        [System.Windows.Forms.MessageBox]::Show(
            "Task completed: $taskName",
            'Toolkit',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }

    $script:CurrentTaskProcess = $null
    $script:CurrentTaskStdoutPath = $null
    $script:CurrentTaskStderrPath = $null
    $script:CurrentTaskOutputLineCount = 0
    $script:CurrentTaskErrorLineCount = 0
    $script:CurrentTaskResult = $null
    $script:CancellationRequested = $false
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

            [void]$apps.Add([pscustomobject]@{
                Name = $name
                Source = 'Desktop app'
            })
        }
    }
    catch {
    }

    try {
        Get-AppxPackage | ForEach-Object {
            [void]$apps.Add([pscustomobject]@{
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

    $pickerLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $pickerLayout.Dock = 'Fill'
    $pickerLayout.Padding = New-Object System.Windows.Forms.Padding(14)
    $pickerLayout.ColumnCount = 1
    $pickerLayout.RowCount = 4
    [void]$pickerLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$pickerLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$pickerLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$pickerLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$picker.Controls.Add($pickerLayout)

    $infoLabel = New-Object System.Windows.Forms.Label
    $infoLabel.Text = 'Check the apps you want included in the review list, then click Save.'
    $infoLabel.AutoSize = $true
    $infoLabel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 8)
    [void]$pickerLayout.Controls.Add($infoLabel, 0, 0)

    $searchBox = New-Object System.Windows.Forms.TextBox
    $searchBox.Dock = 'Fill'
    $searchBox.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)
    [void]$pickerLayout.Controls.Add($searchBox, 0, 1)

    $appChecklist = New-Object System.Windows.Forms.CheckedListBox
    $appChecklist.Dock = 'Fill'
    $appChecklist.CheckOnClick = $true
    [void]$pickerLayout.Controls.Add($appChecklist, 0, 2)

    $pickerButtons = New-Object System.Windows.Forms.FlowLayoutPanel
    $pickerButtons.Dock = 'Fill'
    $pickerButtons.FlowDirection = 'RightToLeft'
    $pickerButtons.WrapContents = $false
    $pickerButtons.Margin = New-Object System.Windows.Forms.Padding(0, 10, 0, 0)
    [void]$pickerLayout.Controls.Add($pickerButtons, 0, 3)

    $saveButton = New-Object System.Windows.Forms.Button
    $saveButton.Text = 'Save'
    $saveButton.Size = New-Object System.Drawing.Size(120, 36)
    [void]$pickerButtons.Controls.Add($saveButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = 'Cancel'
    $cancelButton.Size = New-Object System.Drawing.Size(120, 36)
    [void]$pickerButtons.Controls.Add($cancelButton)

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
                [void]$selected.Add($name)
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

function Get-WindowsUpdateInventory {
    $session = New-Object -ComObject Microsoft.Update.Session
    $searcher = $session.CreateUpdateSearcher()
    $searchResult = $searcher.Search('IsInstalled=0')
    $inventory = @()

    for ($index = 0; $index -lt $searchResult.Updates.Count; $index++) {
        $update = $searchResult.Updates.Item($index)
        $kbIds = @($update.KBArticleIDs | ForEach-Object { 'KB{0}' -f $_ })
        $inventory += [pscustomobject]@{
            Title = [string]$update.Title
            Label = if ($kbIds.Count -gt 0) { "{0} [{1}]" -f $update.Title, ($kbIds -join ', ') } else { [string]$update.Title }
            MatchKeys = @([string]$update.Title) + @($kbIds)
        }
    }

    return $inventory | Sort-Object Title
}

function Show-WindowsUpdateSkipChooser {
    $script:WindowsUpdateInventory = @(Get-WindowsUpdateInventory)

    $picker = New-Object System.Windows.Forms.Form
    $picker.Text = 'Choose Windows Updates To Skip'
    $picker.StartPosition = 'CenterParent'
    $picker.Size = New-Object System.Drawing.Size(760, 560)
    $picker.MinimumSize = New-Object System.Drawing.Size(700, 500)
    $picker.Font = $form.Font

    $pickerLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $pickerLayout.Dock = 'Fill'
    $pickerLayout.Padding = New-Object System.Windows.Forms.Padding(14)
    $pickerLayout.ColumnCount = 1
    $pickerLayout.RowCount = 4
    [void]$pickerLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$pickerLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$pickerLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$pickerLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$picker.Controls.Add($pickerLayout)

    $infoLabel = New-Object System.Windows.Forms.Label
    $infoLabel.Text = 'Check any pending Windows updates you want skipped and hidden for this run.'
    $infoLabel.AutoSize = $true
    $infoLabel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 8)
    [void]$pickerLayout.Controls.Add($infoLabel, 0, 0)

    $searchBox = New-Object System.Windows.Forms.TextBox
    $searchBox.Dock = 'Fill'
    $searchBox.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)
    [void]$pickerLayout.Controls.Add($searchBox, 0, 1)

    $updateChecklist = New-Object System.Windows.Forms.CheckedListBox
    $updateChecklist.Dock = 'Fill'
    $updateChecklist.CheckOnClick = $true
    [void]$pickerLayout.Controls.Add($updateChecklist, 0, 2)

    $pickerButtons = New-Object System.Windows.Forms.FlowLayoutPanel
    $pickerButtons.Dock = 'Fill'
    $pickerButtons.FlowDirection = 'RightToLeft'
    $pickerButtons.WrapContents = $false
    $pickerButtons.Margin = New-Object System.Windows.Forms.Padding(0, 10, 0, 0)
    [void]$pickerLayout.Controls.Add($pickerButtons, 0, 3)

    $saveButton = New-Object System.Windows.Forms.Button
    $saveButton.Text = 'Save'
    $saveButton.Size = New-Object System.Drawing.Size(120, 36)
    [void]$pickerButtons.Controls.Add($saveButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = 'Cancel'
    $cancelButton.Size = New-Object System.Drawing.Size(120, 36)
    [void]$pickerButtons.Controls.Add($cancelButton)

    function Populate-WindowsUpdateList {
        param([string]$FilterText)

        $updateChecklist.Items.Clear()
        foreach ($update in $script:WindowsUpdateInventory) {
            if ($FilterText -and $update.Label -notlike "*$FilterText*") {
                continue
            }

            $isChecked = $false
            foreach ($key in $update.MatchKeys) {
                if ($script:WindowsUpdateSkipSelections -contains $key) {
                    $isChecked = $true
                    break
                }
            }

            [void]$updateChecklist.Items.Add($update.Label, $isChecked)
        }
    }

    Populate-WindowsUpdateList -FilterText ''

    $searchBox.Add_TextChanged({
        Populate-WindowsUpdateList -FilterText $searchBox.Text.Trim()
    })

    $saveButton.Add_Click({
        $selected = New-Object System.Collections.Generic.List[string]
        foreach ($item in $updateChecklist.CheckedItems) {
            $selectedLabel = [string]$item
            $selectedUpdate = $script:WindowsUpdateInventory | Where-Object { $_.Label -eq $selectedLabel } | Select-Object -First 1
            if ($selectedUpdate) {
                foreach ($key in $selectedUpdate.MatchKeys) {
                    if ($key) {
                        [void]$selected.Add($key)
                    }
                }
            }
        }

        $script:WindowsUpdateSkipSelections = @($selected | Sort-Object -Unique)
        $picker.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $picker.Close()
    })

    $cancelButton.Add_Click({
        $picker.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $picker.Close()
    })

    return ($picker.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK)
}

function Get-SelectedTaskFromActiveTab {
    $selectedTab = $tabControl.SelectedTab
    if (-not $selectedTab) {
        return $null
    }

    $selectedCategory = if ($selectedTab.Tag) { [string]$selectedTab.Tag } else { [string]$selectedTab.Text }
    $taskList = $script:TaskListsByCategory[$selectedCategory]
    if (-not $taskList) {
        return $null
    }

    if ($taskList.SelectedItems.Count -eq 0) {
        return $null
    }

    $taskId = $taskList.SelectedItems[0].Tag
    return ($script:Tasks | Where-Object { $_.Id -eq $taskId } | Select-Object -First 1)
}

function Update-SelectedTaskDisplay {
    param(
        $Task
    )

    if (-not $Task) {
        $script:SelectedTask = $null
        $selectedToolLabel.Text = 'Tool Details'
        $descriptionBox.Text = ''
        $statusLabel.Text = 'Status: Ready'
        $metaLabel.Text = 'Requires Admin: Yes | Input: None required'
        $chooseAppsButton.Visible = $false
        $chooseAppsButton.Text = 'Configure'
        return
    }

    $script:SelectedTask = $Task
    $selectedToolLabel.Text = $Task.Name
    $descriptionBox.Text = $Task.Description

    $adminText = if ($Task.RequiresAdmin) { 'Yes' } else { 'No' }
    $inputText = 'None required'
    if ($Task.Id -eq 'Apps.DebloatHelper') {
        $inputText = '{0} selected app(s) for review' -f $script:DebloatSelections.Count
        $chooseAppsButton.Text = 'Choose Apps'
        $toolTip.SetToolTip($chooseAppsButton, 'Choose applications to include in the debloat review list.')
        $chooseAppsButton.Visible = $true
        if ($script:DebloatSelections.Count -gt 0) {
            $descriptionBox.Text += "`r`n`r`nCurrent debloat selection count: $($script:DebloatSelections.Count)"
        }
    }
    elseif ($Task.Id -eq 'Update.WindowsOS') {
        $selectionCount = 0
        foreach ($update in $script:WindowsUpdateInventory) {
            foreach ($key in $update.MatchKeys) {
                if ($script:WindowsUpdateSkipSelections -contains $key) {
                    $selectionCount += 1
                    break
                }
            }
        }
        if ($selectionCount -eq 0 -and $script:WindowsUpdateSkipSelections.Count -gt 0) {
            $selectionCount = $script:WindowsUpdateSkipSelections.Count
        }
        $inputText = '{0} update(s) selected to skip' -f $selectionCount
        $chooseAppsButton.Text = 'Skip Updates'
        $toolTip.SetToolTip($chooseAppsButton, 'Choose pending Windows updates to skip and hide before installation.')
        $chooseAppsButton.Visible = $true
    }
    else {
        $chooseAppsButton.Visible = $false
        $chooseAppsButton.Text = 'Configure'
    }

    $metaLabel.Text = "Requires Admin: $adminText | Input: $inputText"
    if (-not (Test-TaskIsRunning)) {
        $statusLabel.Text = "Status: Ready to run $($Task.Name)"
    }
}

function Select-FirstTaskInTab {
    $selectedTab = $tabControl.SelectedTab
    if (-not $selectedTab) {
        return
    }

    $selectedCategory = if ($selectedTab.Tag) { [string]$selectedTab.Tag } else { [string]$selectedTab.Text }
    $taskList = $script:TaskListsByCategory[$selectedCategory]
    if ($taskList -and $taskList.Items.Count -gt 0 -and $taskList.SelectedItems.Count -eq 0) {
        $taskList.Items[0].Selected = $true
        $taskList.Items[0].Focused = $true
    }
}

function Set-InteractiveState {
    param(
        [Parameter(Mandatory = $true)][bool]$Enabled
    )

    $tabControl.Enabled = $Enabled
    $runButton.Enabled = $Enabled
    $chooseAppsButton.Enabled = $Enabled
    $openLogsButton.Enabled = $Enabled
    $importScriptMenuItem.Enabled = $Enabled
    $editScriptMenuItem.Enabled = $Enabled
    $openLogsMenuItem.Enabled = $Enabled
    $cancelButton.Enabled = (-not $Enabled)
    foreach ($category in $script:TaskListsByCategory.Keys) {
        $script:TaskListsByCategory[$category].Enabled = $Enabled
    }
}

function Start-InteractiveHardwareTool {
    param(
        [Parameter(Mandatory = $true)]$Task
    )

    $progressBar.Style = 'Continuous'
    Set-ToolkitProgress -Value 10
    $statusLabel.Text = "Status: Opening $($Task.Name)..."
    Add-UiLogLine -Line "Opening interactive hardware tool: $($Task.Name)"
    Set-InteractiveState -Enabled $false

    try {
        if ($Task.Id -eq 'Hardware.MouseKeyboardTest') {
            Set-ToolkitProgress -Value 25
            $result = Test-TtkMouseKeyboardActivity
        }
        elseif ($Task.Id -eq 'Hardware.MonitorPixelTest') {
            Set-ToolkitProgress -Value 25
            $result = Start-TtkMonitorPixelTest
        }
        else {
            throw "Unsupported interactive hardware tool: $($Task.Id)"
        }

        if ($result) {
            Add-UiLogLine -Line ([string]$result)
        }
        Set-ToolkitProgress -Value 100
        $statusLabel.Text = "Status: Completed - $($Task.Name)"
        Add-UiLogLine -Line "Interactive hardware tool closed: $($Task.Name)"
        [System.Windows.Forms.MessageBox]::Show(
            "Task completed: $($Task.Name)",
            'Toolkit',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }
    catch {
        Set-ToolkitProgress -Value 0
        $statusLabel.Text = "Status: Failed - $($_.Exception.Message)"
        Add-UiLogLine -Line "Interactive hardware tool failed: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show(
            "Task failed: $($_.Exception.Message)",
            'Toolkit',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
    finally {
        Set-InteractiveState -Enabled $true
    }
}

function Start-ToolkitTask {
    param(
        [Parameter(Mandatory = $true)]$Task
    )

    if (Test-TaskIsRunning) {
        [System.Windows.Forms.MessageBox]::Show(
            'Another task is already running. Wait for it to finish before starting the next one.',
            'Toolkit',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        return
    }

    if ($Task.Id -in @('Hardware.MouseKeyboardTest', 'Hardware.MonitorPixelTest')) {
        Start-InteractiveHardwareTool -Task $Task
        return
    }

    if ($Task.Id -eq 'Apps.DebloatHelper') {
        if (-not (Show-DebloatChooser)) {
            $statusLabel.Text = "Status: Cancelled $($Task.Name)"
            Update-SelectedTaskDisplay -Task $Task
            return
        }
        Update-SelectedTaskDisplay -Task $Task
    }
    elseif ($Task.Id -eq 'Update.WindowsOS') {
        if (-not (Show-WindowsUpdateSkipChooser)) {
            $statusLabel.Text = "Status: Cancelled $($Task.Name)"
            Update-SelectedTaskDisplay -Task $Task
            return
        }
        Update-SelectedTaskDisplay -Task $Task
    }

    Set-ToolkitProgress -Value 0
    $progressBar.Style = 'Marquee'
    $statusLabel.Text = "Status: Starting $($Task.Name)..."
    Add-UiLogLine -Line "Queued task: $($Task.Name)"
    $script:CancellationRequested = $false
    Set-InteractiveState -Enabled $false

    $selectionPath = $null
    $logDir = Get-ToolkitLogDirectory

    if ($Task.Id -eq 'Apps.DebloatHelper' -and $script:DebloatSelections.Count -gt 0) {
        $selectionPath = Join-Path $logDir 'debloat-selection.txt'
        Set-Content -LiteralPath $selectionPath -Value $script:DebloatSelections -Encoding UTF8
    }
    elseif ($Task.Id -eq 'Update.WindowsOS') {
        $selectionPath = Join-Path $logDir 'windowsupdate-skip-selection.txt'
        Set-Content -LiteralPath $selectionPath -Value $script:WindowsUpdateSkipSelections -Encoding UTF8
    }

    $hostScript = Join-Path $toolkitRoot 'Scripts\Invoke-ToolkitTaskHost.ps1'
    if (-not (Test-Path -LiteralPath $hostScript)) {
        Set-InteractiveState -Enabled $true
        throw "Task host script not found: $hostScript"
    }

    $safeTaskName = ($Task.Id -replace '[^a-zA-Z0-9\-]', '-')
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $script:CurrentTaskStdoutPath = Join-Path $logDir ("task-{0}-{1}-stdout.log" -f $safeTaskName, $stamp)
    $script:CurrentTaskStderrPath = Join-Path $logDir ("task-{0}-{1}-stderr.log" -f $safeTaskName, $stamp)
    $script:CurrentTaskOutputLineCount = 0
    $script:CurrentTaskErrorLineCount = 0
    $script:CurrentTaskResult = $null

    $argumentParts = @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', "`"$hostScript`"",
        '-ToolkitRoot', "`"$toolkitRoot`"",
        '-TaskId', "`"$($Task.Id)`""
    )

    if ($selectionPath) {
        $argumentParts += @('-SelectionFile', "`"$selectionPath`"")
    }

    $script:CurrentTaskProcess = Start-Process -FilePath 'powershell.exe' `
        -ArgumentList ($argumentParts -join ' ') `
        -PassThru `
        -WindowStyle Hidden `
        -RedirectStandardOutput $script:CurrentTaskStdoutPath `
        -RedirectStandardError $script:CurrentTaskStderrPath

    $taskMonitorTimer.Start()
}

Refresh-TaskTabs

$taskMonitorTimer.Add_Tick({
    if (-not $script:CurrentTaskProcess) {
        $taskMonitorTimer.Stop()
        return
    }

    $outputCount = $script:CurrentTaskOutputLineCount
    foreach ($line in (Get-NewFileLines -Path $script:CurrentTaskStdoutPath -ProcessedCount ([ref]$outputCount))) {
        Process-TaskOutputLine -Line ([string]$line)
    }
    $script:CurrentTaskOutputLineCount = $outputCount

    $errorCount = $script:CurrentTaskErrorLineCount
    foreach ($line in (Get-NewFileLines -Path $script:CurrentTaskStderrPath -ProcessedCount ([ref]$errorCount))) {
        Process-TaskOutputLine -Line ([string]$line) -IsErrorStream
    }
    $script:CurrentTaskErrorLineCount = $errorCount

    if ($script:CurrentTaskProcess.HasExited) {
        Complete-RunningTask
    }
})

$runButton.Add_Click({
    $selectedTask = Get-SelectedTaskFromActiveTab
    if (-not $selectedTask) {
        [System.Windows.Forms.MessageBox]::Show(
            'Select a tool from the active tab first.',
            'Toolkit',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        return
    }

    Start-ToolkitTask -Task $selectedTask
})

$cancelButton.Add_Click({
    if (-not (Test-TaskIsRunning)) {
        return
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        'Cancel the running task?',
        'Toolkit',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
        return
    }

    $script:CancellationRequested = $true
    $statusLabel.Text = 'Status: Cancelling running task...'
    Add-UiLogLine -Line 'Cancellation requested for the running task.'
    Stop-RunningTaskProcessTree
})

$chooseAppsButton.Add_Click({
    if ($script:SelectedTask -and $script:SelectedTask.Id -eq 'Apps.DebloatHelper') {
        if (Show-DebloatChooser) {
            Update-SelectedTaskDisplay -Task $script:SelectedTask
            $statusLabel.Text = 'Status: Debloat selection updated'
        }
    }
    elseif ($script:SelectedTask -and $script:SelectedTask.Id -eq 'Update.WindowsOS') {
        if (Show-WindowsUpdateSkipChooser) {
            Update-SelectedTaskDisplay -Task $script:SelectedTask
            $statusLabel.Text = 'Status: Windows Update skip list updated'
        }
    }
})

$openLogsButton.Add_Click({
    Open-ToolkitLogsFolder
})

$importScriptMenuItem.Add_Click({ Import-ScriptIntoToolkit })
$openLogsMenuItem.Add_Click({ Open-ToolkitLogsFolder })
$darkModeMenuItem.Add_Click({
    $script:IsDarkMode = [bool]$darkModeMenuItem.Checked
    Apply-ToolkitTheme
    Add-UiLogLine -Line ("Dark mode {0}." -f $(if ($script:IsDarkMode) { 'enabled' } else { 'disabled' }))
})
$exitMenuItem.Add_Click({ $form.Close() })
$editScriptMenuItem.Add_Click({ Edit-SelectedScript })
$aboutMenuItem.Add_Click({ Show-ToolkitAbout })

$tabControl.Add_SelectedIndexChanged({
    Select-FirstTaskInTab
    $selectedTask = Get-SelectedTaskFromActiveTab
    if ($selectedTask) {
        Update-SelectedTaskDisplay -Task $selectedTask
    }
})

function Apply-SplitterLayout {
    $availableWidth = $contentSplit.Width
    if ($availableWidth -le 0) {
        return
    }

    $contentSplit.Panel1MinSize = 0
    $contentSplit.Panel2MinSize = 0

    $desiredLeftWidth = 560
    $minimumLeftWidth = 320
    $minimumRightWidth = 520
    $maxLeftWidth = $availableWidth - $minimumRightWidth
    $effectiveLeftMin = [Math]::Min($minimumLeftWidth, [Math]::Max(160, ($availableWidth - 180)))
    $effectiveRightMin = [Math]::Min($minimumRightWidth, [Math]::Max(160, ($availableWidth - $effectiveLeftMin - 1)))

    if ($maxLeftWidth -lt $minimumLeftWidth) {
        $splitterDistance = [Math]::Max(140, [Math]::Floor($availableWidth * 0.42))
    }
    else {
        $splitterDistance = [Math]::Max($minimumLeftWidth, [Math]::Min($desiredLeftWidth, $maxLeftWidth))
    }

    $splitterDistance = [Math]::Max(1, [Math]::Min($splitterDistance, ($availableWidth - 1)))
    $splitterDistance = [Math]::Max($effectiveLeftMin, [Math]::Min($splitterDistance, ($availableWidth - $effectiveRightMin - 1)))
    $contentSplit.SplitterDistance = [int]$splitterDistance
    $contentSplit.Panel1MinSize = $effectiveLeftMin
    $contentSplit.Panel2MinSize = [Math]::Max(120, [Math]::Min($effectiveRightMin, ($availableWidth - $splitterDistance - 1)))
}

$form.Add_Shown({
    Apply-ToolkitTheme
    Apply-SplitterLayout
    if ($tabControl.TabPages.Count -gt 0) {
        $tabControl.SelectedIndex = 0
        Select-FirstTaskInTab
        $selectedTask = Get-SelectedTaskFromActiveTab
        if ($selectedTask) {
            Update-SelectedTaskDisplay -Task $selectedTask
        }
    }
})

$form.Add_Resize({
    Apply-SplitterLayout
})

$form.Add_KeyDown({
    param($sender, $e)

    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::F5) {
        $selectedTask = Get-SelectedTaskFromActiveTab
        if ($selectedTask) {
            Start-ToolkitTask -Task $selectedTask
            $e.Handled = $true
        }
    }
})

[void]$form.ShowDialog()
