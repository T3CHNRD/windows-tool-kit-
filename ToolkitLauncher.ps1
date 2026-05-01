#Requires -Version 5.1
<#
.SYNOPSIS
    T3CHNRD's Windows Tool Kit — v2 Launcher
    Dual-mode UI: Simple Mode (mom-friendly) + IT Pro Mode (full power)

.DESCRIPTION
    Drop-in replacement for ToolkitLauncher.ps1.
    - Keeps ALL existing modules, scripts, tasks, and categories unchanged.
    - Adds a mode toggle between Simple Mode and IT Pro Mode.
    - Simple Mode: plain-English names, limited safe tools, no log, friendly completion screen.
    - IT Pro Mode: all tools, full log, PRO badges, technical descriptions.
    - Both modes share the same BackgroundWorker execution engine underneath.

.NOTES
    Place this file in the same root folder as your existing ToolkitLauncher.ps1.
    Run with: powershell.exe -NoProfile -ExecutionPolicy Bypass -File ToolkitLauncher_v2.ps1
#>

[CmdletBinding()]
param(
    [switch]$ElevatedRelaunch
)

Set-StrictMode -Version Latest
# FIX: CRITICAL-03 - keep WinForms event/paint/timer warnings from becoming app-killing exceptions.
$ErrorActionPreference = 'Continue'

# FIX: CRITICAL-01 - capture the script invocation at script scope before helper functions run.
$script:ScriptInvocation = $MyInvocation

# ============================================================
#  SECTION 1 — ELEVATION (unchanged from your original)
# ============================================================

function Test-LauncherIsAdmin {
    $identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-LauncherCommandPath {
    # FIX: CRITICAL-01 - use the launcher invocation, not this function's invocation scope.
    $commandInfo = $script:ScriptInvocation.MyCommand
    if ($commandInfo -and $commandInfo.PSObject.Properties['Path']) {
        $commandPath = $commandInfo.Path
        if ($commandPath) { return $commandPath }
    }

    $commandLineArgs = [Environment]::GetCommandLineArgs()
    if ($commandLineArgs.Count -gt 0) {
        foreach ($arg in $commandLineArgs) {
            if (-not $arg) { continue }
            $candidate = $arg.Trim('"')
            if (-not (Test-Path -LiteralPath $candidate)) { continue }
            $item = Get-Item -LiteralPath $candidate -ErrorAction SilentlyContinue
            if (-not $item) { continue }

            $extension = [IO.Path]::GetExtension($item.FullName)
            $fileName = [IO.Path]::GetFileNameWithoutExtension($item.FullName)
            if ($extension -ieq '.ps1' -or ($extension -ieq '.exe' -and $fileName -notmatch '^(powershell|pwsh)$')) {
                return $item.FullName
            }
        }
    }

    # FIX: MED-19 - PS2EXE can run from the extracted EXE path while $PSScriptRoot points elsewhere.
    try {
        $exePath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
        $exeName = [IO.Path]::GetFileNameWithoutExtension($exePath)
        if ($exePath -and (Test-Path -LiteralPath $exePath) -and $exeName -notmatch '^(powershell|pwsh)$') {
            return (Get-Item -LiteralPath $exePath).FullName
        }
    }
    catch {}

    return $null
}

function Get-LauncherRootPath {
    $commandPath = Get-LauncherCommandPath
    if ($commandPath) { return (Split-Path -Parent $commandPath) }
    return [System.AppDomain]::CurrentDomain.BaseDirectory.TrimEnd('\')
}

function Restart-LauncherElevated {
    param([Parameter(Mandatory)][string]$ToolkitRoot)
    $commandPath = Get-LauncherCommandPath
    if (-not $commandPath) { throw 'Could not determine the launcher path for elevation.' }
    $startInfo                  = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.Verb             = 'runas'
    $startInfo.WorkingDirectory = $ToolkitRoot
    $startInfo.UseShellExecute  = $true
    if ([IO.Path]::GetExtension($commandPath) -eq '.exe') {
        $startInfo.FileName = $commandPath
    } else {
        # FIX: CRITICAL-02 - relaunch through -Command with single-quote escaping for paths with spaces/special chars.
        $escapedPath = $commandPath -replace "'", "''"
        $escapedRoot = $ToolkitRoot -replace "'", "''"
        $startInfo.FileName  = 'powershell.exe'
        $startInfo.Arguments = "-NoProfile -ExecutionPolicy Bypass -Command `"Set-Location -LiteralPath '$escapedRoot'; & '$escapedPath' -ElevatedRelaunch`""
    }
    [void][System.Diagnostics.Process]::Start($startInfo)
}

$toolkitRoot = Get-LauncherRootPath

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

if (-not (Test-LauncherIsAdmin)) {
    try { Restart-LauncherElevated -ToolkitRoot $toolkitRoot }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Administrator rights required.`r`n`r`n$($_.Exception.Message)",
            "T3CHNRD's Windows Tool Kit",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
    }
    return
}

# ============================================================
#  SECTION 2 — LOAD YOUR EXISTING MODULES (unchanged)
# ============================================================

$modulePath = Join-Path $toolkitRoot 'Modules\MaintenanceToolkit\MaintenanceToolkit.psm1'
if (-not (Test-Path -LiteralPath $modulePath)) {
    [System.Windows.Forms.MessageBox]::Show(
        "The toolkit app files were not found next to the launcher.`r`n`r`nExpected:`r`n$modulePath`r`n`r`nIf you opened the portable ZIP directly, extract the ZIP to a folder or USB drive first, then run the EXE from the extracted folder.",
        "T3CHNRD's Windows Tool Kit",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    return
}

try {
    Import-Module $modulePath -Force -ErrorAction Stop
}
catch {
    [System.Windows.Forms.MessageBox]::Show(
        "The toolkit module could not be loaded.`r`n`r`n$($_.Exception.Message)`r`n`r`nMake sure the full app folder is extracted and contains the Modules folder.",
        "T3CHNRD's Windows Tool Kit",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    return
}

$hardwareModulePath = Join-Path $toolkitRoot 'Modules\HardwareDiagnostics\HardwareDiagnostics.psm1'
if (Test-Path -LiteralPath $hardwareModulePath) {
    Import-Module $hardwareModulePath -Force -ErrorAction Stop
}

[System.Windows.Forms.Application]::EnableVisualStyles()

# ============================================================
#  SECTION 3 — SCRIPT STATE
# ============================================================

$script:IsSimpleMode            = $false   # <— THE MODE FLAG
$script:SelectedTask            = $null
$script:Tasks                   = @()
$script:CategoryOrder           = @()
$script:TasksByCategory         = @{}
$script:CurrentTaskProcess      = $null
$script:CancellationRequested   = $false
$script:ProgressValue           = 0
$script:ProgressAnimFrame       = 0
$script:DebloatSelections       = @()
$script:DebloatInventory        = @()
$script:WindowsUpdateSkipSel    = @()
$script:WindowsUpdateInventory  = @()
$script:TaskListsByCategory     = @{}
$script:CurrentTaskStdoutPath   = $null
$script:CurrentTaskStderrPath   = $null
$script:CurrentTaskStdoutCount  = 0
$script:CurrentTaskStderrCount  = 0
$script:CurrentTaskResult       = $null

# FIX: CRITICAL-04 - pre-create reusable GDI resources and dispose them on form close.
$script:CardBorderPen           = New-Object System.Drawing.Pen(([System.Drawing.Color]::FromArgb(30, 36, 52)), 1)

# FIX: HIGH-04 - mirror launcher UI log messages to disk so crashes do not lose context.
$script:UiLogPath               = Join-Path (Join-Path $toolkitRoot 'Logs') ("Toolkit-UI-{0:yyyyMMdd}.log" -f (Get-Date))

# ============================================================
#  SECTION 4 — SIMPLE MODE TASK MAP
#  Maps friendly names → your existing task IDs
# ============================================================

$script:SimpleTasks = @(
    [PSCustomObject]@{
        FriendlyName = 'Clean Up My Computer'
        FriendlyDesc = 'Removes junk files that slow your computer down. Nothing important gets deleted — completely safe.'
        TaskId       = 'Cleanup.TempFiles'
        Icon         = '🧹'
        SafetyNote   = 'Safe · No personal files touched'
    },
    [PSCustomObject]@{
        FriendlyName = 'Free Up Storage Space'
        FriendlyDesc = "Shows what's taking up space and helps you clear it safely. Will ask before deleting anything."
        TaskId       = 'Cleanup.FreeSpace'
        Icon         = '💾'
        SafetyNote   = 'Safe · Shows preview first'
    },
    [PSCustomObject]@{
        FriendlyName = 'Fix Windows Problems'
        FriendlyDesc = 'Scans and repairs Windows system files automatically. Great when things feel slow or broken. Takes about 10 minutes.'
        TaskId       = 'Repair.WindowsHealth'
        Icon         = '🔧'
        SafetyNote   = 'Safe · Microsoft built-in repair tools'
    },
    [PSCustomObject]@{
        FriendlyName = 'Check My Internet'
        FriendlyDesc = 'Tests your internet connection and explains any problems in plain English — with suggestions on how to fix them.'
        TaskId       = 'Network.Diagnostics'
        Icon         = '🌐'
        SafetyNote   = 'Safe · Read-only scan'
    },
    [PSCustomObject]@{
        FriendlyName = 'Update All My Apps'
        FriendlyDesc = 'Updates all your installed apps at once. Safer and faster than doing each one manually.'
        TaskId       = 'Update.AllApps'
        Icon         = '🔄'
        SafetyNote   = 'Safe · Uses Windows built-in winget'
    },
    [PSCustomObject]@{
        FriendlyName = 'Install Microsoft 365'
        FriendlyDesc = 'Walks you through installing Microsoft 365 the official way. You will need your Microsoft account.'
        TaskId       = 'Integration.Microsoft365'
        Icon         = '📦'
        SafetyNote   = 'Official Microsoft installer only'
    }
)

# ============================================================
#  SECTION 5 — COLOR PALETTES
# ============================================================

function Get-Palette {
    if ($script:IsSimpleMode) {
        return @{
            # Light / Simple
            FormBack     = [System.Drawing.Color]::FromArgb(240, 244, 255)
            Header       = [System.Drawing.Color]::FromArgb(255, 255, 255)
            HeaderBorder = [System.Drawing.Color]::FromArgb(220, 228, 245)
            Sidebar      = [System.Drawing.Color]::FromArgb(255, 255, 255)
            SidebarBord  = [System.Drawing.Color]::FromArgb(220, 228, 245)
            CardBack     = [System.Drawing.Color]::FromArgb(255, 255, 255)
            CardHover    = [System.Drawing.Color]::FromArgb(235, 242, 255)
            CardActive   = [System.Drawing.Color]::FromArgb(219, 234, 254)
            CardBorder   = [System.Drawing.Color]::FromArgb(210, 222, 245)
            CardActBord  = [System.Drawing.Color]::FromArgb(147, 197, 253)
            PanelRight   = [System.Drawing.Color]::FromArgb(255, 255, 255)
            PanelBorder  = [System.Drawing.Color]::FromArgb(220, 228, 245)
            LogBack      = [System.Drawing.Color]::FromArgb(247, 250, 255)
            RunBtn       = [System.Drawing.Color]::FromArgb(37, 99, 235)
            RunBtnText   = [System.Drawing.Color]::White
            AccentBar    = [System.Drawing.Color]::FromArgb(37, 99, 235)
            Text         = [System.Drawing.Color]::FromArgb(15, 23, 65)
            TextMid      = [System.Drawing.Color]::FromArgb(71, 85, 135)
            TextDim      = [System.Drawing.Color]::FromArgb(148, 163, 200)
            BarTrack     = [System.Drawing.Color]::FromArgb(219, 228, 245)
            BarFill      = [System.Drawing.Color]::FromArgb(37, 99, 235)
            Green        = [System.Drawing.Color]::FromArgb(22, 163, 74)
            Yellow       = [System.Drawing.Color]::FromArgb(161, 98, 7)
            Red          = [System.Drawing.Color]::FromArgb(185, 28, 28)
        }
    }
    # Dark / Pro
    return @{
        FormBack     = [System.Drawing.Color]::FromArgb(11, 13, 18)
        Header       = [System.Drawing.Color]::FromArgb(18, 21, 29)
        HeaderBorder = [System.Drawing.Color]::FromArgb(30, 36, 52)
        Sidebar      = [System.Drawing.Color]::FromArgb(18, 21, 29)
        SidebarBord  = [System.Drawing.Color]::FromArgb(30, 36, 52)
        CardBack     = [System.Drawing.Color]::FromArgb(18, 21, 29)
        CardHover    = [System.Drawing.Color]::FromArgb(24, 28, 40)
        CardActive   = [System.Drawing.Color]::FromArgb(20, 30, 55)
        CardBorder   = [System.Drawing.Color]::FromArgb(30, 36, 52)
        CardActBord  = [System.Drawing.Color]::FromArgb(79, 142, 247)
        PanelRight   = [System.Drawing.Color]::FromArgb(18, 21, 29)
        PanelBorder  = [System.Drawing.Color]::FromArgb(30, 36, 52)
        LogBack      = [System.Drawing.Color]::FromArgb(7, 10, 15)
        RunBtn       = [System.Drawing.Color]::FromArgb(79, 142, 247)
        RunBtnText   = [System.Drawing.Color]::White
        AccentBar    = [System.Drawing.Color]::FromArgb(79, 142, 247)
        Text         = [System.Drawing.Color]::FromArgb(226, 232, 244)
        TextMid      = [System.Drawing.Color]::FromArgb(139, 149, 173)
        TextDim      = [System.Drawing.Color]::FromArgb(69, 79, 102)
        BarTrack     = [System.Drawing.Color]::FromArgb(30, 35, 51)
        BarFill      = [System.Drawing.Color]::FromArgb(79, 142, 247)
        Green        = [System.Drawing.Color]::FromArgb(52, 211, 153)
        Yellow       = [System.Drawing.Color]::FromArgb(251, 191, 36)
        Red          = [System.Drawing.Color]::FromArgb(248, 113, 113)
    }
}

# ============================================================
#  SECTION 6 — FORM & LAYOUT CONSTRUCTION
# ============================================================

$form                  = New-Object System.Windows.Forms.Form
$form.Text             = "T3CHNRD's Windows Tool Kit"
$form.StartPosition    = 'CenterScreen'
$form.Size             = New-Object System.Drawing.Size(1400, 880)
$form.MinimumSize      = New-Object System.Drawing.Size(1100, 700)
$form.BackColor        = [System.Drawing.Color]::FromArgb(11, 13, 18)
$form.Font             = New-Object System.Drawing.Font('Segoe UI', 10)
$form.KeyPreview       = $true

# ── Root layout ──────────────────────────────────────────────
$rootPanel             = New-Object System.Windows.Forms.TableLayoutPanel
$rootPanel.Dock        = 'Fill'
$rootPanel.ColumnCount = 1
$rootPanel.RowCount    = 2
[void]$rootPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 58)))
[void]$rootPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$form.Controls.Add($rootPanel)

# ── Title bar ────────────────────────────────────────────────
$titleBar              = New-Object System.Windows.Forms.Panel
$titleBar.Dock         = 'Fill'
$titleBar.BackColor    = [System.Drawing.Color]::FromArgb(18, 21, 29)
$titleBar.Padding      = New-Object System.Windows.Forms.Padding(16, 0, 16, 0)
[void]$rootPanel.Controls.Add($titleBar, 0, 0)

# Logo label
$logoLabel             = New-Object System.Windows.Forms.Label
$logoLabel.Text        = 'T3'
$logoLabel.Font        = New-Object System.Drawing.Font('Segoe UI Black', 9, [System.Drawing.FontStyle]::Bold)
$logoLabel.ForeColor   = [System.Drawing.Color]::White
$logoLabel.BackColor   = [System.Drawing.Color]::FromArgb(79, 142, 247)
$logoLabel.Size        = New-Object System.Drawing.Size(32, 32)
$logoLabel.Location    = New-Object System.Drawing.Point(16, 13)
$logoLabel.TextAlign   = 'MiddleCenter'
[void]$titleBar.Controls.Add($logoLabel)

# Title
$titleLabel            = New-Object System.Windows.Forms.Label
$titleLabel.Text       = "T3CHNRD's Windows Tool Kit"
$titleLabel.Font       = New-Object System.Drawing.Font('Segoe UI Semibold', 13, [System.Drawing.FontStyle]::Bold)
$titleLabel.ForeColor  = [System.Drawing.Color]::FromArgb(226, 232, 244)
$titleLabel.AutoSize   = $true
$titleLabel.Location   = New-Object System.Drawing.Point(56, 10)
[void]$titleBar.Controls.Add($titleLabel)

# Mode sub-label
$modeSubLabel          = New-Object System.Windows.Forms.Label
$modeSubLabel.Text     = 'IT Pro Mode'
$modeSubLabel.Font     = New-Object System.Drawing.Font('Segoe UI', 9)
$modeSubLabel.ForeColor= [System.Drawing.Color]::FromArgb(69, 79, 102)
$modeSubLabel.AutoSize = $true
$modeSubLabel.Location = New-Object System.Drawing.Point(57, 33)
[void]$titleBar.Controls.Add($modeSubLabel)

# ── Mode Toggle Button ────────────────────────────────────────
$modeToggleBtn         = New-Object System.Windows.Forms.Button
$modeToggleBtn.Text    = '🙂  Switch to Simple Mode'
$modeToggleBtn.Font    = New-Object System.Drawing.Font('Segoe UI Emoji', 10, [System.Drawing.FontStyle]::Bold)
$modeToggleBtn.ForeColor = [System.Drawing.Color]::White
$modeToggleBtn.BackColor = [System.Drawing.Color]::FromArgb(79, 142, 247)
$modeToggleBtn.FlatStyle = 'Flat'
$modeToggleBtn.FlatAppearance.BorderSize = 0
$modeToggleBtn.Size    = New-Object System.Drawing.Size(220, 36)
$modeToggleBtn.Cursor  = [System.Windows.Forms.Cursors]::Hand
[void]$titleBar.Controls.Add($modeToggleBtn)

# Clock label (right side of titlebar)
$clockLabel            = New-Object System.Windows.Forms.Label
$clockLabel.Font       = New-Object System.Drawing.Font('Consolas', 11)
$clockLabel.ForeColor  = [System.Drawing.Color]::FromArgb(69, 79, 102)
$clockLabel.AutoSize   = $true
$clockLabel.TextAlign  = 'MiddleRight'
[void]$titleBar.Controls.Add($clockLabel)

# Position toggle & clock on resize
$titleBar.Add_Resize({
    $modeToggleBtn.Location = New-Object System.Drawing.Point(($titleBar.Width - $modeToggleBtn.Width - 120), 11)
    $clockLabel.Location    = New-Object System.Drawing.Point(($titleBar.Width - 95), 20)
})

# ── Body: sidebar + main ─────────────────────────────────────
$bodySplit                    = New-Object System.Windows.Forms.SplitContainer
$bodySplit.Dock               = 'Fill'
$bodySplit.FixedPanel         = 'Panel1'
# FIX: MARKET-01 - keep startup min sizes tiny until the container has a real width.
# WinForms validates splitter constraints immediately, even before layout has completed.
$bodySplit.Panel1MinSize      = 1
$bodySplit.Panel2MinSize      = 1
$bodySplit.BackColor          = [System.Drawing.Color]::FromArgb(30, 36, 52)
[void]$rootPanel.Controls.Add($bodySplit, 0, 1)

# ── Sidebar ───────────────────────────────────────────────────
$sidebar               = New-Object System.Windows.Forms.Panel
$sidebar.Dock          = 'Fill'
$sidebar.BackColor     = [System.Drawing.Color]::FromArgb(18, 21, 29)
[void]$bodySplit.Panel1.Controls.Add($sidebar)

$sidebarLayout         = New-Object System.Windows.Forms.TableLayoutPanel
$sidebarLayout.Dock    = 'Fill'
$sidebarLayout.ColumnCount = 1
$sidebarLayout.RowCount    = 2
[void]$sidebarLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 44)))
[void]$sidebarLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$sidebar.Controls.Add($sidebarLayout)

# Search box
$searchBox             = New-Object System.Windows.Forms.TextBox
$searchBox.Dock        = 'Fill'
$searchBox.Font        = New-Object System.Drawing.Font('Segoe UI Emoji', 10)
$searchBox.BackColor   = [System.Drawing.Color]::FromArgb(24, 28, 40)
$searchBox.ForeColor   = [System.Drawing.Color]::FromArgb(139, 149, 173)
$searchBox.BorderStyle = 'FixedSingle'
$searchBox.Text        = '🔍  Search tools…'
$searchBox.Margin      = New-Object System.Windows.Forms.Padding(8, 6, 8, 6)
[void]$sidebarLayout.Controls.Add($searchBox, 0, 0)

$searchBox.Add_Enter({ if ($searchBox.Text -eq '🔍  Search tools…') { $searchBox.Text = ''; $searchBox.ForeColor = (Get-Palette).Text } })
$searchBox.Add_Leave({ if ($searchBox.Text -eq '') { $searchBox.Text = '🔍  Search tools…'; $searchBox.ForeColor = (Get-Palette).TextDim } })

# Nav list
$navList               = New-Object System.Windows.Forms.ListBox
$navList.Dock          = 'Fill'
$navList.BorderStyle   = 'None'
$navList.BackColor     = [System.Drawing.Color]::FromArgb(18, 21, 29)
$navList.ForeColor     = [System.Drawing.Color]::FromArgb(139, 149, 173)
$navList.Font          = New-Object System.Drawing.Font('Segoe UI Emoji', 10)
$navList.ItemHeight    = 34
$navList.DrawMode      = 'OwnerDrawFixed'
[void]$sidebarLayout.Controls.Add($navList, 0, 1)

# ── Main split: tool area + right panel ───────────────────────
$mainSplit                  = New-Object System.Windows.Forms.SplitContainer
$mainSplit.Dock             = 'Fill'
$mainSplit.FixedPanel       = 'Panel2'
# FIX: MARKET-01 - set real splitter distances only after layout has completed.
$mainSplit.Panel1MinSize    = 1
$mainSplit.Panel2MinSize    = 1
$mainSplit.BackColor        = [System.Drawing.Color]::FromArgb(30, 36, 52)
[void]$bodySplit.Panel2.Controls.Add($mainSplit)

# ── Left of main: tool header + tab/card area ────────────────
$leftMain              = New-Object System.Windows.Forms.TableLayoutPanel
$leftMain.Dock         = 'Fill'
$leftMain.ColumnCount  = 1
$leftMain.RowCount     = 3
[void]$leftMain.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 80)))
[void]$leftMain.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 36)))
[void]$leftMain.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$mainSplit.Panel1.Controls.Add($leftMain)

# Tool header
$toolHeader            = New-Object System.Windows.Forms.Panel
$toolHeader.Dock       = 'Fill'
$toolHeader.BackColor  = [System.Drawing.Color]::FromArgb(18, 21, 29)
$toolHeader.Padding    = New-Object System.Windows.Forms.Padding(16, 10, 16, 10)
[void]$leftMain.Controls.Add($toolHeader, 0, 0)

$toolNameLabel         = New-Object System.Windows.Forms.Label
$toolNameLabel.Font    = New-Object System.Drawing.Font('Segoe UI Semibold', 14, [System.Drawing.FontStyle]::Bold)
$toolNameLabel.ForeColor = [System.Drawing.Color]::FromArgb(226, 232, 244)
$toolNameLabel.AutoSize  = $true
$toolNameLabel.Location  = New-Object System.Drawing.Point(0, 4)
[void]$toolHeader.Controls.Add($toolNameLabel)

$toolDescLabel         = New-Object System.Windows.Forms.Label
$toolDescLabel.Font    = New-Object System.Drawing.Font('Segoe UI', 10)
$toolDescLabel.ForeColor = [System.Drawing.Color]::FromArgb(139, 149, 173)
$toolDescLabel.AutoSize  = $true
$toolDescLabel.MaximumSize = New-Object System.Drawing.Size(700, 0)
$toolDescLabel.Location  = New-Object System.Drawing.Point(1, 36)
[void]$toolHeader.Controls.Add($toolDescLabel)

# Tab strip (Pro Mode) / hidden in Simple
$tabStrip              = New-Object System.Windows.Forms.TabControl
$tabStrip.Dock         = 'Fill'
$tabStrip.Padding      = New-Object System.Drawing.Point(12, 4)
$tabStrip.Font         = New-Object System.Drawing.Font('Segoe UI', 9)
[void]$leftMain.Controls.Add($tabStrip, 0, 1)

# Card panel (scrollable flow) — used in BOTH modes
$cardPanel             = New-Object System.Windows.Forms.FlowLayoutPanel
$cardPanel.Dock        = 'Fill'
$cardPanel.AutoScroll  = $true
$cardPanel.BackColor   = [System.Drawing.Color]::FromArgb(11, 13, 18)
$cardPanel.Padding     = New-Object System.Windows.Forms.Padding(12)
$cardPanel.WrapContents = $true
[void]$leftMain.Controls.Add($cardPanel, 0, 2)

# ── Right panel ───────────────────────────────────────────────
$rightLayout           = New-Object System.Windows.Forms.TableLayoutPanel
$rightLayout.Dock      = 'Fill'
$rightLayout.ColumnCount = 1
$rightLayout.RowCount  = 8
$rightLayout.BackColor = [System.Drawing.Color]::FromArgb(18, 21, 29)
$rightLayout.Padding   = New-Object System.Windows.Forms.Padding(14, 12, 14, 12)
[void]$mainSplit.Panel2.Controls.Add($rightLayout)

[void]$rightLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))   # 0 selected label
[void]$rightLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 130))) # 1 desc
[void]$rightLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))   # 2 meta
[void]$rightLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 52)))  # 3 run btn
[void]$rightLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 60)))  # 4 progress
[void]$rightLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))   # 5 health header
[void]$rightLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 90)))  # 6 health grid
[void]$rightLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) # 7 log

# Row 0 — selected label
$rpSelectedLabel       = New-Object System.Windows.Forms.Label
$rpSelectedLabel.Text  = 'SELECTED TOOL'
$rpSelectedLabel.Font  = New-Object System.Drawing.Font('Consolas', 8, [System.Drawing.FontStyle]::Bold)
$rpSelectedLabel.ForeColor = [System.Drawing.Color]::FromArgb(69, 79, 102)
$rpSelectedLabel.AutoSize  = $true
$rpSelectedLabel.Margin    = New-Object System.Windows.Forms.Padding(0, 0, 0, 4)
[void]$rightLayout.Controls.Add($rpSelectedLabel, 0, 0)

# Row 1 — description box
$rpDescBox             = New-Object System.Windows.Forms.TextBox
$rpDescBox.Multiline   = $true
$rpDescBox.ReadOnly    = $true
$rpDescBox.Dock        = 'Fill'
$rpDescBox.BackColor   = [System.Drawing.Color]::FromArgb(24, 28, 40)
$rpDescBox.ForeColor   = [System.Drawing.Color]::FromArgb(139, 149, 173)
$rpDescBox.BorderStyle = 'FixedSingle'
$rpDescBox.Font        = New-Object System.Drawing.Font('Segoe UI', 10)
$rpDescBox.ScrollBars  = 'Vertical'
$rpDescBox.Text        = 'Select a tool to see details.'
$rpDescBox.Margin      = New-Object System.Windows.Forms.Padding(0, 0, 0, 8)
[void]$rightLayout.Controls.Add($rpDescBox, 0, 1)

# Row 2 — meta strip
$rpMetaLabel           = New-Object System.Windows.Forms.Label
$rpMetaLabel.Text      = ''
$rpMetaLabel.Font      = New-Object System.Drawing.Font('Consolas', 9)
$rpMetaLabel.ForeColor = [System.Drawing.Color]::FromArgb(251, 191, 36)
$rpMetaLabel.AutoSize  = $true
$rpMetaLabel.Margin    = New-Object System.Windows.Forms.Padding(0, 0, 0, 8)
[void]$rightLayout.Controls.Add($rpMetaLabel, 0, 2)

# Row 3 — Run button
$runBtn                = New-Object System.Windows.Forms.Button
$runBtn.Dock           = 'Fill'
$runBtn.Text           = '▶  Run Selected Tool'
$runBtn.Font           = New-Object System.Drawing.Font('Segoe UI Semibold', 11, [System.Drawing.FontStyle]::Bold)
$runBtn.BackColor      = [System.Drawing.Color]::FromArgb(79, 142, 247)
$runBtn.ForeColor      = [System.Drawing.Color]::White
$runBtn.FlatStyle      = 'Flat'
$runBtn.FlatAppearance.BorderSize = 0
$runBtn.Cursor         = [System.Windows.Forms.Cursors]::Hand
$runBtn.Margin         = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)
[void]$rightLayout.Controls.Add($runBtn, 0, 3)

# Row 4 — Progress
$progressPanel         = New-Object System.Windows.Forms.Panel
$progressPanel.Dock    = 'Fill'
$progressPanel.BackColor = [System.Drawing.Color]::FromArgb(18, 21, 29)
[void]$rightLayout.Controls.Add($progressPanel, 0, 4)

$statusLabel           = New-Object System.Windows.Forms.Label
$statusLabel.Text      = 'Status: Ready'
$statusLabel.Font      = New-Object System.Drawing.Font('Segoe UI', 9)
$statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(139, 149, 173)
$statusLabel.AutoSize  = $true
$statusLabel.Location  = New-Object System.Drawing.Point(0, 0)
[void]$progressPanel.Controls.Add($statusLabel)

$progressBar           = New-Object System.Windows.Forms.ProgressBar
$progressBar.Minimum   = 0
$progressBar.Maximum   = 100
$progressBar.Value     = 0
$progressBar.Style     = 'Continuous'
$progressBar.Size      = New-Object System.Drawing.Size(270, 10)
$progressBar.Location  = New-Object System.Drawing.Point(0, 22)
[void]$progressPanel.Controls.Add($progressBar)

$cancelBtn             = New-Object System.Windows.Forms.Button
$cancelBtn.Text        = 'Cancel'
$cancelBtn.Font        = New-Object System.Drawing.Font('Segoe UI', 9)
$cancelBtn.Size        = New-Object System.Drawing.Size(80, 26)
$cancelBtn.Location    = New-Object System.Drawing.Point(0, 36)
$cancelBtn.FlatStyle   = 'Flat'
$cancelBtn.BackColor   = [System.Drawing.Color]::FromArgb(248, 113, 113)
$cancelBtn.ForeColor   = [System.Drawing.Color]::White
$cancelBtn.FlatAppearance.BorderSize = 0
$cancelBtn.Enabled     = $false
[void]$progressPanel.Controls.Add($cancelBtn)

$progressPanel.Add_Resize({
    $progressBar.Width = [Math]::Max(50, $progressPanel.Width - 10)
})

# Row 5 — Health header (Pro only)
$healthHeaderLabel     = New-Object System.Windows.Forms.Label
$healthHeaderLabel.Text= 'SYSTEM HEALTH'
$healthHeaderLabel.Font= New-Object System.Drawing.Font('Consolas', 8, [System.Drawing.FontStyle]::Bold)
$healthHeaderLabel.ForeColor = [System.Drawing.Color]::FromArgb(69, 79, 102)
$healthHeaderLabel.AutoSize  = $true
$healthHeaderLabel.Margin    = New-Object System.Windows.Forms.Padding(0, 6, 0, 4)
[void]$rightLayout.Controls.Add($healthHeaderLabel, 0, 5)

# Row 6 — Health grid
$healthPanel           = New-Object System.Windows.Forms.TableLayoutPanel
$healthPanel.Dock      = 'Fill'
$healthPanel.ColumnCount = 4
$healthPanel.RowCount  = 1
$healthPanel.BackColor = [System.Drawing.Color]::FromArgb(18, 21, 29)
for ($i = 0; $i -lt 4; $i++) {
    [void]$healthPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25)))
}
[void]$rightLayout.Controls.Add($healthPanel, 0, 6)

function New-HealthTile {
    param([string]$Name, [string]$Value, [System.Drawing.Color]$ValueColor)
    $tile = New-Object System.Windows.Forms.Panel
    $tile.Dock      = 'Fill'
    $tile.BackColor = [System.Drawing.Color]::FromArgb(24, 28, 40)
    $tile.Margin    = New-Object System.Windows.Forms.Padding(2)
    $n = New-Object System.Windows.Forms.Label
    $n.Text      = $Name
    $n.Font      = New-Object System.Drawing.Font('Consolas', 7, [System.Drawing.FontStyle]::Bold)
    $n.ForeColor = [System.Drawing.Color]::FromArgb(69, 79, 102)
    $n.AutoSize  = $true
    $n.Location  = New-Object System.Drawing.Point(6, 6)
    $v = New-Object System.Windows.Forms.Label
    $v.Text      = $Value
    $v.Font      = New-Object System.Drawing.Font('Consolas', 11, [System.Drawing.FontStyle]::Bold)
    $v.ForeColor = $ValueColor
    $v.AutoSize  = $true
    $v.Location  = New-Object System.Drawing.Point(6, 22)
    [void]$tile.Controls.Add($n)
    [void]$tile.Controls.Add($v)
    return $tile
}

[void]$healthPanel.Controls.Add((New-HealthTile 'CPU'     '12%'    ([System.Drawing.Color]::FromArgb(79,142,247))), 0, 0)
[void]$healthPanel.Controls.Add((New-HealthTile 'RAM'     '4.1 GB' ([System.Drawing.Color]::FromArgb(52,211,153))), 1, 0)
[void]$healthPanel.Controls.Add((New-HealthTile 'C: DISK' '67%'    ([System.Drawing.Color]::FromArgb(251,191,36))), 2, 0)
[void]$healthPanel.Controls.Add((New-HealthTile 'UPTIME'  '6h 14m' ([System.Drawing.Color]::FromArgb(139,149,173))), 3, 0)

# Row 7 — Log (Pro) / Friendly completion (Simple)
$logPanel              = New-Object System.Windows.Forms.Panel
$logPanel.Dock         = 'Fill'
[void]$rightLayout.Controls.Add($logPanel, 0, 7)

# Pro log box
$logHeaderLabel        = New-Object System.Windows.Forms.Label
$logHeaderLabel.Text   = 'EXECUTION LOG'
$logHeaderLabel.Font   = New-Object System.Drawing.Font('Consolas', 8, [System.Drawing.FontStyle]::Bold)
$logHeaderLabel.ForeColor = [System.Drawing.Color]::FromArgb(69, 79, 102)
$logHeaderLabel.AutoSize  = $true
$logHeaderLabel.Location  = New-Object System.Drawing.Point(0, 0)
[void]$logPanel.Controls.Add($logHeaderLabel)

$logBox                = New-Object System.Windows.Forms.TextBox
$logBox.Multiline      = $true
$logBox.ReadOnly       = $true
$logBox.ScrollBars     = 'Vertical'
$logBox.BackColor      = [System.Drawing.Color]::FromArgb(7, 10, 15)
$logBox.ForeColor      = [System.Drawing.Color]::FromArgb(139, 149, 173)
$logBox.BorderStyle    = 'FixedSingle'
$logBox.Font           = New-Object System.Drawing.Font('Consolas', 9)
$logBox.Location       = New-Object System.Drawing.Point(0, 18)
[void]$logPanel.Controls.Add($logBox)

# Simple completion panel
$simpleStatusPanel     = New-Object System.Windows.Forms.Panel
$simpleStatusPanel.Dock= 'Fill'
$simpleStatusPanel.BackColor = [System.Drawing.Color]::FromArgb(240, 253, 244)
$simpleStatusPanel.Visible   = $false

$simpleIconLabel       = New-Object System.Windows.Forms.Label
$simpleIconLabel.Text  = '✅'
$simpleIconLabel.Font  = New-Object System.Drawing.Font('Segoe UI Emoji', 28)
$simpleIconLabel.AutoSize   = $true
$simpleIconLabel.TextAlign  = 'MiddleCenter'
[void]$simpleStatusPanel.Controls.Add($simpleIconLabel)

$simpleTitleLabel      = New-Object System.Windows.Forms.Label
$simpleTitleLabel.Text = 'All Done!'
$simpleTitleLabel.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 14, [System.Drawing.FontStyle]::Bold)
$simpleTitleLabel.ForeColor = [System.Drawing.Color]::FromArgb(22, 163, 74)
$simpleTitleLabel.AutoSize  = $true
[void]$simpleStatusPanel.Controls.Add($simpleTitleLabel)

$simpleMsgLabel        = New-Object System.Windows.Forms.Label
$simpleMsgLabel.Text   = ''
$simpleMsgLabel.Font   = New-Object System.Drawing.Font('Segoe UI', 11)
$simpleMsgLabel.ForeColor = [System.Drawing.Color]::FromArgb(71, 85, 135)
$simpleMsgLabel.MaximumSize = New-Object System.Drawing.Size(260, 0)
$simpleMsgLabel.AutoSize    = $true
[void]$simpleStatusPanel.Controls.Add($simpleMsgLabel)

[void]$logPanel.Controls.Add($simpleStatusPanel)

$logPanel.Add_Resize({
    $logBox.Size     = New-Object System.Drawing.Size($logPanel.Width, [Math]::Max(20, $logPanel.Height - 22))
    $w = $simpleStatusPanel.Width
    $simpleIconLabel.Location  = New-Object System.Drawing.Point(([Math]::Max(0,($w - $simpleIconLabel.Width)/2)), 20)
    $simpleTitleLabel.Location = New-Object System.Drawing.Point(([Math]::Max(0,($w - $simpleTitleLabel.Width)/2)), 72)
    $simpleMsgLabel.Location   = New-Object System.Drawing.Point(([Math]::Max(0,($w - 260)/2)), 108)
})

# ============================================================
#  SECTION 7 — NAV LIST OWNER DRAW
# ============================================================

# Pro nav items: category names + section headers
$script:ProNavItems = @(
    @{ Text='─── Maintenance';   IsHeader=$true  },
    @{ Text='🧹  Cleanup';       IsHeader=$false; Category='Cleanup'            },
    @{ Text='🔧  Repair';        IsHeader=$false; Category='Repair'             },
    @{ Text='🔄  Updates';       IsHeader=$false; Category='Update Tools'       },
    @{ Text='─── Network';       IsHeader=$true  },
    @{ Text='🌐  Network Tools'; IsHeader=$false; Category='Network Tools'      },
    @{ Text='─── Storage';       IsHeader=$true  },
    @{ Text='💾  Storage/Setup'; IsHeader=$false; Category='Storage / Setup'    },
    @{ Text='─── Security [PRO]';IsHeader=$true  },
    @{ Text='🛡️  Security';      IsHeader=$false; Category='Security'           },
    @{ Text='─── Apps & HW';     IsHeader=$true  },
    @{ Text='📦  Applications';  IsHeader=$false; Category='Applications'       },
    @{ Text='🖥️  Hardware Diag'; IsHeader=$false; Category='Hardware Diagnostics'},
    @{ Text='💻  Legacy Scripts';IsHeader=$false; Category='Misc/Utility'       }
)

# Simple nav items
$script:SimpleNavItems = @(
    @{ Text='─── What do you need?'; IsHeader=$true },
    @{ Text='🧹  Clean Up';           IsHeader=$false; SimpleIdx=0 },
    @{ Text='💾  Free Up Space';      IsHeader=$false; SimpleIdx=1 },
    @{ Text='🔧  Fix Problems';       IsHeader=$false; SimpleIdx=2 },
    @{ Text='🌐  Internet Help';      IsHeader=$false; SimpleIdx=3 },
    @{ Text='🔄  Update Apps';        IsHeader=$false; SimpleIdx=4 },
    @{ Text='📦  Install Office';     IsHeader=$false; SimpleIdx=5 }
)

$navList.Add_DrawItem({
    param($sender, $e)
    $pal = Get-Palette
    if ($e.Index -lt 0 -or $e.Index -ge $navList.Items.Count) { return }

    $itemData = $navList.Items[$e.Index]
    $isHeader = $itemData.IsHeader
    $isSelected = ($e.State -band [System.Windows.Forms.DrawItemState]::Selected) -ne 0

    if ($isHeader) {
        $e.Graphics.FillRectangle((New-Object System.Drawing.SolidBrush($pal.Sidebar)), $e.Bounds)
        $brush = New-Object System.Drawing.SolidBrush($pal.TextDim)
        $font  = New-Object System.Drawing.Font('Consolas', 8, [System.Drawing.FontStyle]::Bold)
        $sf    = New-Object System.Drawing.StringFormat
        $sf.LineAlignment = 'Center'
        $e.Graphics.DrawString($itemData.Text, $font, $brush,
            [System.Drawing.RectangleF]::new($e.Bounds.X + 10, $e.Bounds.Y, $e.Bounds.Width - 10, $e.Bounds.Height), $sf)
        $brush.Dispose(); $font.Dispose(); $sf.Dispose()
    } else {
        $bgColor = if ($isSelected) { $pal.CardActive } else { $pal.Sidebar }
        $e.Graphics.FillRectangle((New-Object System.Drawing.SolidBrush($bgColor)), $e.Bounds)
        if ($isSelected) {
            # Accent bar on left
            $e.Graphics.FillRectangle(
                (New-Object System.Drawing.SolidBrush($pal.AccentBar)),
                [System.Drawing.Rectangle]::new($e.Bounds.X, $e.Bounds.Y + 5, 3, $e.Bounds.Height - 10)
            )
        }
        $fgColor = if ($isSelected) { $pal.Text } else { $pal.TextMid }
        $brush   = New-Object System.Drawing.SolidBrush($fgColor)
        $navFontStyle = if ($isSelected) { [System.Drawing.FontStyle]::Bold } else { [System.Drawing.FontStyle]::Regular }
        $font    = New-Object System.Drawing.Font('Segoe UI Emoji', 10, $navFontStyle)
        $sf      = New-Object System.Drawing.StringFormat
        $sf.LineAlignment = 'Center'
        $e.Graphics.DrawString($itemData.Text, $font, $brush,
            [System.Drawing.RectangleF]::new($e.Bounds.X + 14, $e.Bounds.Y, $e.Bounds.Width - 14, $e.Bounds.Height), $sf)
        $brush.Dispose(); $font.Dispose(); $sf.Dispose()
    }
})

# ============================================================
#  SECTION 8 — TASK CATALOG & CARD BUILDING
# ============================================================

function Refresh-TaskCatalog {
    # FIX: HIGH-01 - do not turn a null catalog into an array containing a null task.
    $rawTasks = Get-ToolkitTaskCatalog
    $script:Tasks         = if ($rawTasks) { @($rawTasks | Where-Object { $_ } | Sort-Object Category, Name) } else { @() }
    $script:CategoryOrder = @(Get-ToolkitCategoryOrder | Where-Object { $script:Tasks.Category -contains $_ })
    $script:TasksByCategory = @{}
    foreach ($cat in $script:CategoryOrder) {
        $script:TasksByCategory[$cat] = @($script:Tasks | Where-Object { $_.Category -eq $cat })
    }
}

function New-ToolCard {
    param(
        [string]$Name,
        [string]$Description,
        [string]$Icon,
        [bool]$IsProOnly = $false,
        [string]$TaskId = '',
        [bool]$RequiresAdmin = $false
    )

    $pal        = Get-Palette
    $isSimple   = $script:IsSimpleMode
    $cardW      = if ($isSimple) { 280 } else { 220 }
    $cardH      = if ($isSimple) { 120 } else { 110 }

    $card                  = New-Object System.Windows.Forms.Panel
    $card.Size             = New-Object System.Drawing.Size($cardW, $cardH)
    $card.BackColor        = $pal.CardBack
    $card.Cursor           = [System.Windows.Forms.Cursors]::Hand
    $card.Margin           = New-Object System.Windows.Forms.Padding(6)
    $card.Tag              = $TaskId

    # Border effect via Paint
    $card.Add_Paint({
        param($s, $pe)
        try {
            # FIX: CRITICAL-04 - use a reusable pen and cast dimensions before arithmetic.
            $paintControl = if ($s -is [array]) { $s[0] } else { $s }
            $paintWidth = [int]$paintControl.ClientSize.Width
            $paintHeight = [int]$paintControl.ClientSize.Height
            if ($paintWidth -le 1 -or $paintHeight -le 1) { return }
            $pe.Graphics.DrawRectangle($script:CardBorderPen, 0, 0, ($paintWidth - 1), ($paintHeight - 1))
        }
        catch {
            Add-LogLine "Card paint skipped: $($_.Exception.Message)"
        }
    })

    # Icon
    $iconLbl               = New-Object System.Windows.Forms.Label
    $iconLbl.Text          = $Icon
    $iconFontSize          = if ($isSimple) { 22 } else { 16 }
    $iconLbl.Font          = New-Object System.Drawing.Font('Segoe UI Emoji', $iconFontSize)
    $iconLbl.AutoSize      = $true
    $iconLbl.Location      = New-Object System.Drawing.Point(10, 8)
    [void]$card.Controls.Add($iconLbl)

    # Pro badge
    if ($IsProOnly -and -not $isSimple) {
        $badge             = New-Object System.Windows.Forms.Label
        $badge.Text        = 'PRO'
        $badge.Font        = New-Object System.Drawing.Font('Consolas', 7, [System.Drawing.FontStyle]::Bold)
        $badge.ForeColor   = [System.Drawing.Color]::FromArgb(56, 217, 245)
        $badge.BackColor   = [System.Drawing.Color]::FromArgb(14, 50, 65)
        $badge.AutoSize    = $true
        $badge.Padding     = New-Object System.Windows.Forms.Padding(3, 1, 3, 1)
        $badge.Location    = New-Object System.Drawing.Point(($cardW - 36), 6)
        [void]$card.Controls.Add($badge)
    }

    # Name
    $nameLbl               = New-Object System.Windows.Forms.Label
    $nameLbl.Text          = $Name
    $nameFontSize          = if ($isSimple) { 12 } else { 10 }
    $nameLbl.Font          = New-Object System.Drawing.Font('Segoe UI Semibold', $nameFontSize, [System.Drawing.FontStyle]::Bold)
    $nameLbl.ForeColor     = $pal.Text
    $nameLbl.MaximumSize   = New-Object System.Drawing.Size(($cardW - 18), 0)
    $nameLbl.AutoSize      = $true
    $nameY                 = if ($isSimple) { 46 } else { 40 }
    $nameLbl.Location      = New-Object System.Drawing.Point(10, $nameY)
    [void]$card.Controls.Add($nameLbl)

    # Desc
    $descLbl               = New-Object System.Windows.Forms.Label
    $descLbl.Text          = $Description
    $descFontSize          = if ($isSimple) { 9 } else { 8 }
    $descLbl.Font          = New-Object System.Drawing.Font('Segoe UI', $descFontSize)
    $descLbl.ForeColor     = $pal.TextMid
    $descLbl.MaximumSize   = New-Object System.Drawing.Size(($cardW - 18), 0)
    $descLbl.AutoSize      = $true
    $descY                 = if ($isSimple) { 72 } else { 62 }
    $descLbl.Location      = New-Object System.Drawing.Point(10, $descY)
    [void]$card.Controls.Add($descLbl)

    # Hover
    $hoverEnter = {
        $this.BackColor = (Get-Palette).CardHover
    }
    $hoverLeave = {
        if ($script:SelectedTask -and $script:SelectedTask.Id -eq $this.Tag) {
            $this.BackColor = (Get-Palette).CardActive
        } else {
            $this.BackColor = (Get-Palette).CardBack
        }
    }
    $card.Add_MouseEnter($hoverEnter)
    $card.Add_MouseLeave($hoverLeave)
    foreach ($ctrl in $card.Controls) {
        $ctrl.Add_MouseEnter($hoverEnter)
        $ctrl.Add_MouseLeave($hoverLeave)
    }

    return $card
}

function Populate-ProCards {
    param([string]$Category)
    $cardPanel.Controls.Clear()
    $tasks = $script:TasksByCategory[$Category]
    if (-not $tasks) { return }
    foreach ($task in $tasks) {
        $card = New-ToolCard `
            -Name         $task.Name `
            -Description  $task.Description `
            -Icon         (Get-TaskIcon $task.Id) `
            -IsProOnly    ($task.Id -like 'Security.*' -or $task.Id -like 'Hardware.*') `
            -TaskId       $task.Id `
            -RequiresAdmin $task.RequiresAdmin

        $card.Add_Click({
            Select-Task -TaskId $this.Tag
        })
        foreach ($ctrl in $card.Controls) {
            $ctrl.Add_Click({ Select-Task -TaskId $this.Parent.Tag })
        }
        [void]$cardPanel.Controls.Add($card)
    }
}

function Populate-SimpleCards {
    $cardPanel.Controls.Clear()
    foreach ($st in $script:SimpleTasks) {
        $card = New-ToolCard `
            -Name        $st.FriendlyName `
            -Description $st.FriendlyDesc `
            -Icon        $st.Icon `
            -TaskId      $st.TaskId

        $card.Add_Click({
            Select-SimpleTask -TaskId $this.Tag
        })
        foreach ($ctrl in $card.Controls) {
            $ctrl.Add_Click({ Select-SimpleTask -TaskId $this.Parent.Tag })
        }
        [void]$cardPanel.Controls.Add($card)
    }
}

function Get-TaskIcon {
    param([string]$Id)
    $icons = @{
        'Cleanup.TempFiles'          = '🧹'
        'Cleanup.FreeSpace'          = '💾'
        'Disk.Monitor'               = '📊'
        'Apps.DebloatHelper'         = '🗑️'
        'Repair.WindowsHealth'       = '🔧'
        'Network.Diagnostics'        = '🌐'
        'Network.DhcpRenew'          = '🔁'
        'Network.ResetStack'         = '♻️'
        'Update.AllApps'             = '🔄'
        'Update.WindowsOS'           = '🪟'
        'Update.VendorDrivers'       = '🖥️'
        'Update.VendorBIOS'          = '⚡'
        'Update.VendorFirmware'      = '⚙️'
        'Integration.Microsoft365'   = '📦'
        'Integration.MediaCreationTool' = '💿'
        'Security.BaselineAudit'     = '🛡️'
        'Security.DefenderQuickScan' = '🔍'
        'Security.OpenPortsAudit'    = '🔒'
        'Security.PowerShellRiskScan'= '⚠️'
        'Hardware.MouseKeyboardTest' = '🖱️'
        'Hardware.MonitorPixelTest'  = '🖥️'
        'Storage.CloneGuide'         = '📀'
        'Storage.NewComputerSetup'   = '🆕'
    }
    if ($icons.ContainsKey($Id)) { return $icons[$Id] }
    if ($Id -like 'Legacy.*')    { return '💻' }
    return '🔧'
}

# ============================================================
#  SECTION 9 — TASK SELECTION
# ============================================================

function Select-Task {
    param([string]$TaskId)
    $task = $script:Tasks | Where-Object { $_.Id -eq $TaskId } | Select-Object -First 1
    if (-not $task) { return }
    $script:SelectedTask = $task

    # Update cards highlight
    foreach ($ctrl in $cardPanel.Controls) {
        $ctrl.BackColor = if ($ctrl.Tag -eq $TaskId) { (Get-Palette).CardActive } else { (Get-Palette).CardBack }
    }

    # Update right panel (Pro mode)
    $rpDescBox.Text    = $task.Description
    $rpMetaLabel.Text  = $(if ($task.RequiresAdmin) { '⚡ Requires Admin  ' } else { '' }) + "Category: $($task.Category)"
    $toolNameLabel.Text = (Get-TaskIcon $task.Id) + '  ' + $task.Name
    $toolDescLabel.Text = $task.Description
    $simpleStatusPanel.Visible = $false
}

function Select-SimpleTask {
    param([string]$TaskId)
    $st = $script:SimpleTasks | Where-Object { $_.TaskId -eq $TaskId } | Select-Object -First 1
    if (-not $st) { return }
    $realTask = $script:Tasks | Where-Object { $_.Id -eq $TaskId } | Select-Object -First 1
    $script:SelectedTask = $realTask

    foreach ($ctrl in $cardPanel.Controls) {
        $ctrl.BackColor = if ($ctrl.Tag -eq $TaskId) { (Get-Palette).CardActive } else { (Get-Palette).CardBack }
    }

    $rpDescBox.Text    = $st.FriendlyDesc
    $rpMetaLabel.Text  = $st.SafetyNote
    $toolNameLabel.Text = $st.Icon + '  ' + $st.FriendlyName
    $toolDescLabel.Text = $st.FriendlyDesc
    $simpleStatusPanel.Visible = $false
    $simpleMsgLabel.Text = "$($st.FriendlyName) finished successfully."
}

# ============================================================
#  SECTION 10 — LOGGING (shared)
# ============================================================

function Add-LogLine {
    param([string]$Line)
    if ([string]::IsNullOrWhiteSpace($Line)) { return }
    $ts = "[{0:HH:mm:ss}] {1}" -f (Get-Date), $Line
    try {
        $logBox.AppendText($ts + [Environment]::NewLine)
    }
    catch {}

    # FIX: HIGH-04 - persist launcher messages to Logs\Toolkit-UI-yyyyMMdd.log.
    try {
        $logDir = Split-Path -Parent $script:UiLogPath
        if (-not (Test-Path -LiteralPath $logDir)) {
            New-Item -Path $logDir -ItemType Directory -Force | Out-Null
        }
        Add-Content -LiteralPath $script:UiLogPath -Value $ts -Encoding UTF8
    }
    catch {}
}

# ============================================================
#  SECTION 11 — TASK EXECUTION ENGINE
#  Background process host so long-running tools do not freeze the UI.
# ============================================================

function Test-TaskRunning {
    return ($script:CurrentTaskProcess -and -not $script:CurrentTaskProcess.HasExited)
}

function ConvertTo-LauncherArgument {
    param([Parameter(Mandatory = $true)][string]$Value)
    return '"' + ($Value -replace '"', '`"') + '"'
}

function Clear-TaskProcessFiles {
    foreach ($path in @($script:CurrentTaskStdoutPath, $script:CurrentTaskStderrPath)) {
        if ($path -and (Test-Path -LiteralPath $path)) {
            Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue
        }
    }
    $script:CurrentTaskStdoutPath = $null
    $script:CurrentTaskStderrPath = $null
    $script:CurrentTaskStdoutCount = 0
    $script:CurrentTaskStderrCount = 0
}

function Receive-TaskOutputFile {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][ref]$ProcessedCount,
        [switch]$IsError
    )

    if (-not (Test-Path -LiteralPath $Path)) { return }
    $allLines = @(Get-Content -LiteralPath $Path -ErrorAction SilentlyContinue)
    if ($ProcessedCount.Value -ge $allLines.Count) { return }

    for ($idx = [int]$ProcessedCount.Value; $idx -lt $allLines.Count; $idx++) {
        $line = [string]$allLines[$idx]
        if ([string]::IsNullOrWhiteSpace($line)) { continue }

        if ($line -like '__TTK__|*') {
            $parts = @($line -split '\|')
            switch ($parts[1]) {
                'PROGRESS' {
                    $pct = 0
                    [int]::TryParse($parts[2], [ref]$pct) | Out-Null
                    Set-Progress $pct
                    if ($parts.Count -gt 3 -and -not [string]::IsNullOrWhiteSpace($parts[3])) {
                        $statusLabel.Text = $parts[3]
                        Add-LogLine $parts[3]
                    }
                }
                'STATUS' {
                    if ($parts.Count -gt 2 -and -not [string]::IsNullOrWhiteSpace($parts[2])) {
                        $statusLabel.Text = $parts[2]
                        Add-LogLine $parts[2]
                    }
                }
                'LOG' {
                    if ($parts.Count -gt 2) { Add-LogLine ($parts[2..($parts.Count - 1)] -join '|') }
                }
                'RESULT' {
                    $script:CurrentTaskResult = [pscustomobject]@{
                        Status = if ($parts.Count -gt 2) { $parts[2] } else { 'UNKNOWN' }
                        Name = if ($parts.Count -gt 3) { $parts[3] } else { $script:SelectedTask.Name }
                        Message = if ($parts.Count -gt 4) { ($parts[4..($parts.Count - 1)] -join '|') } else { '' }
                    }
                }
                default {
                    Add-LogLine $line
                }
            }
        }
        else {
            Add-LogLine $(if ($IsError) { "ERROR: $line" } else { $line })
        }
    }

    $ProcessedCount.Value = $allLines.Count
}

function Complete-CurrentTask {
    param(
        [Parameter(Mandatory = $true)][int]$ExitCode
    )

    $task = $script:SelectedTask
    $success = ($ExitCode -eq 0 -and $script:CurrentTaskResult -and $script:CurrentTaskResult.Status -eq 'SUCCESS')
    $message = if ($script:CurrentTaskResult -and $script:CurrentTaskResult.Message) {
        $script:CurrentTaskResult.Message
    }
    elseif ($success) {
        'Task completed successfully.'
    }
    else {
        "Task exited with code $ExitCode."
    }

    $runBtn.Enabled = $true
    $cancelBtn.Enabled = $false
    Set-Progress $(if ($success) { 100 } else { 0 })
    $statusLabel.Text = if ($success) { "Completed: $($task.Name)" } else { "Failed: $($task.Name)" }
    Add-LogLine $(if ($success) { "$($task.Name) completed successfully." } else { "$($task.Name) failed: $message" })

    if ($script:IsSimpleMode) {
        $simpleStatusPanel.Visible = $true
        if ($success) {
            $simpleIconLabel.Text = 'OK'
            $simpleTitleLabel.Text = 'All Done!'
            $simpleTitleLabel.ForeColor = [System.Drawing.Color]::FromArgb(22, 163, 74)
            $simpleStatusPanel.BackColor = [System.Drawing.Color]::FromArgb(240, 253, 244)
            $simpleMsgLabel.Text = "$($task.Name) finished successfully."
        }
        else {
            $simpleIconLabel.Text = '!'
            $simpleTitleLabel.Text = 'Something went wrong'
            $simpleTitleLabel.ForeColor = [System.Drawing.Color]::FromArgb(185, 28, 28)
            $simpleStatusPanel.BackColor = [System.Drawing.Color]::FromArgb(254, 242, 242)
            $simpleMsgLabel.Text = "The task did not finish. Details were saved in the log."
        }
    }

    if ($script:CurrentTaskProcess) {
        try { $script:CurrentTaskProcess.Dispose() } catch {}
    }
    $script:CurrentTaskProcess = $null
    Clear-TaskProcessFiles
}

$taskMonitorTimer = New-Object System.Windows.Forms.Timer
$taskMonitorTimer.Interval = 200
$taskMonitorTimer.Add_Tick({
    try {
        if ($script:CurrentTaskStdoutPath) {
            Receive-TaskOutputFile -Path $script:CurrentTaskStdoutPath -ProcessedCount ([ref]$script:CurrentTaskStdoutCount)
        }
        if ($script:CurrentTaskStderrPath) {
            Receive-TaskOutputFile -Path $script:CurrentTaskStderrPath -ProcessedCount ([ref]$script:CurrentTaskStderrCount) -IsError
        }

        if ($script:CurrentTaskProcess -and $script:CurrentTaskProcess.HasExited) {
            $exitCode = $script:CurrentTaskProcess.ExitCode
            $taskMonitorTimer.Stop()
            Complete-CurrentTask -ExitCode $exitCode
        }
    }
    catch {
        Add-LogLine "Task monitor skipped: $($_.Exception.Message)"
    }
})

function Stop-TaskProcessTree {
    if (-not $script:CurrentTaskProcess) { return }
    # FIX: CRITICAL-05 - log taskkill success/failure instead of swallowing cancellation errors.
    try {
        $result = & taskkill.exe /PID $script:CurrentTaskProcess.Id /T /F 2>&1
        Add-LogLine "Cancel: taskkill result: $($result -join ' ')"
    }
    catch {
        Add-LogLine "Cancel: taskkill failed - $($_.Exception.Message)"
    }
}

function Set-Progress {
    param([int]$Value)
    $v = [Math]::Max(0, [Math]::Min(100, $Value))
    $progressBar.Value  = $v
    $script:ProgressValue = $v
}

function Start-ToolkitTaskProcess {
    param([Parameter(Mandatory = $true)]$Task)

    Clear-TaskProcessFiles
    $script:CurrentTaskResult = $null
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss-fff'
    $script:CurrentTaskStdoutPath = Join-Path $env:TEMP ("T3CHNRD-Launcher-$stamp-stdout.log")
    $script:CurrentTaskStderrPath = Join-Path $env:TEMP ("T3CHNRD-Launcher-$stamp-stderr.log")

    $hostScript = Join-Path $toolkitRoot 'Scripts\Invoke-ToolkitTaskHost.ps1'
    if (-not (Test-Path -LiteralPath $hostScript)) {
        throw "Task host script not found: $hostScript"
    }

    $args = @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', (ConvertTo-LauncherArgument $hostScript),
        '-ToolkitRoot', (ConvertTo-LauncherArgument $toolkitRoot),
        '-TaskId', (ConvertTo-LauncherArgument $Task.Id)
    ) -join ' '

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    # FIX: MARKET-04 - cmd.exe owns redirection; powershell.exe -File does not reliably
    # apply trailing redirection when launched directly through ProcessStartInfo.
    $command = '"powershell.exe" ' + $args + " 1> $(ConvertTo-LauncherArgument $script:CurrentTaskStdoutPath) 2> $(ConvertTo-LauncherArgument $script:CurrentTaskStderrPath)"
    $psi.FileName = $env:ComSpec
    $psi.Arguments = '/d /s /c "' + $command + '"'
    $psi.WorkingDirectory = $toolkitRoot
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $psi
    if (-not $process.Start()) {
        throw "Could not start task host for $($Task.Name)."
    }

    $script:CurrentTaskProcess = $process
    $taskMonitorTimer.Start()
}

$runBtn.Add_Click({
    if (-not $script:SelectedTask) {
        [System.Windows.Forms.MessageBox]::Show(
            'Please select a tool first.',
            "T3CHNRD's Windows Tool Kit",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        return
    }
    if (Test-TaskRunning) {
        [System.Windows.Forms.MessageBox]::Show(
            'A task is already running. Please wait or cancel it.',
            "T3CHNRD's Windows Tool Kit",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    $task = $script:SelectedTask
    $simpleStatusPanel.Visible = $false
    $statusLabel.Text  = "Running: $($task.Name)…"
    Set-Progress 0
    $runBtn.Enabled    = $false
    $cancelBtn.Enabled = $true

    Add-LogLine "Starting: $($task.Name)"

    try {
        # FIX: MARKET-02 - run task scripts in a child host so the WinForms UI remains responsive.
        Start-ToolkitTaskProcess -Task $task
    } catch {
        $runBtn.Enabled    = $true
        $cancelBtn.Enabled = $false
        Add-LogLine "ERROR: $_"
        $statusLabel.Text  = "Error — see log for details"
    }
})

$cancelBtn.Add_Click({
    $script:CancellationRequested = $true
    Stop-TaskProcessTree
    $taskMonitorTimer.Stop()
    $statusLabel.Text  = 'Cancelled.'
    $cancelBtn.Enabled = $false
    $runBtn.Enabled    = $true
    Add-LogLine 'Task cancelled by user.'
    Clear-TaskProcessFiles
})

# ============================================================
#  SECTION 12 — MODE SWITCH ENGINE
# ============================================================

function Apply-Mode {
    $pal    = Get-Palette
    $simple = $script:IsSimpleMode

    # Form + panels
    $form.BackColor         = $pal.FormBack
    $titleBar.BackColor     = $pal.Header
    $sidebar.BackColor      = $pal.Sidebar
    $cardPanel.BackColor    = $pal.FormBack
    $rightLayout.BackColor  = $pal.PanelRight
    $progressPanel.BackColor= $pal.PanelRight
    $logBox.BackColor       = $pal.LogBack
    $toolHeader.BackColor   = $pal.Header
    $rpDescBox.BackColor    = $pal.CardBack
    $searchBox.BackColor    = $pal.CardHover
    $navList.BackColor      = $pal.Sidebar
    $bodySplit.BackColor    = $pal.HeaderBorder
    $mainSplit.BackColor    = $pal.HeaderBorder

    $toolNameLabel.ForeColor  = $pal.Text
    $toolDescLabel.ForeColor  = $pal.TextMid
    $statusLabel.ForeColor    = $pal.TextMid
    $rpDescBox.ForeColor      = $pal.TextMid
    $rpMetaLabel.ForeColor    = $pal.Yellow
    $logBox.ForeColor         = $pal.TextMid
    $rpSelectedLabel.ForeColor= $pal.TextDim
    $healthHeaderLabel.ForeColor = $pal.TextDim
    $logHeaderLabel.ForeColor    = $pal.TextDim
    $titleLabel.ForeColor     = $pal.Text
    $clockLabel.ForeColor     = $pal.TextMid
    $searchBox.ForeColor      = if ($searchBox.Text -eq '🔍  Search tools…') { $pal.TextDim } else { $pal.Text }

    $runBtn.BackColor   = $pal.RunBtn
    $runBtn.ForeColor   = $pal.RunBtnText

    # Simple-mode specific changes
    $tabStrip.Visible              = -not $simple
    $healthHeaderLabel.Visible     = -not $simple
    $healthPanel.Visible           = -not $simple
    $logHeaderLabel.Visible        = -not $simple
    $logBox.Visible                = -not $simple

    if ($simple) {
        $modeToggleBtn.Text        = '⚙️  Switch to IT Pro Mode'
        $modeToggleBtn.BackColor   = [System.Drawing.Color]::FromArgb(22, 163, 74)
        $modeSubLabel.Text         = 'Simple Mode — friendly & safe'
        $modeSubLabel.ForeColor    = [System.Drawing.Color]::FromArgb(22, 163, 74)
        $runBtn.Text               = '▶  Go — Fix It!'
        $runBtn.Font               = New-Object System.Drawing.Font('Segoe UI Semibold', 13, [System.Drawing.FontStyle]::Bold)
        $logoLabel.BackColor       = [System.Drawing.Color]::FromArgb(22, 163, 74)
        $simpleStatusPanel.Visible = $false

        # Rebuild nav
        $navList.Items.Clear()
        foreach ($item in $script:SimpleNavItems) { [void]$navList.Items.Add($item) }

        $toolNameLabel.Text = '🙂  What would you like to fix today?'
        $toolDescLabel.Text = 'Pick something below — each tool explains what it does before starting.'
        $rpDescBox.BackColor = [System.Drawing.Color]::FromArgb(255, 255, 255)
        $rpDescBox.ForeColor = [System.Drawing.Color]::FromArgb(15, 23, 65)

        Populate-SimpleCards
    } else {
        $modeToggleBtn.Text        = '🙂  Switch to Simple Mode'
        $modeToggleBtn.BackColor   = [System.Drawing.Color]::FromArgb(79, 142, 247)
        $modeSubLabel.Text         = 'IT Pro Mode'
        $modeSubLabel.ForeColor    = [System.Drawing.Color]::FromArgb(69, 79, 102)
        $runBtn.Text               = '▶  Run Selected Tool'
        $runBtn.Font               = New-Object System.Drawing.Font('Segoe UI Semibold', 11, [System.Drawing.FontStyle]::Bold)
        $logoLabel.BackColor       = [System.Drawing.Color]::FromArgb(79, 142, 247)

        # Rebuild nav
        $navList.Items.Clear()
        foreach ($item in $script:ProNavItems) { [void]$navList.Items.Add($item) }

        # Rebuild tabs
        $tabStrip.TabPages.Clear()
        foreach ($cat in $script:CategoryOrder) {
            $tp       = New-Object System.Windows.Forms.TabPage
            $tp.Text  = (Get-CategoryTabLabel $cat)
            $tp.Tag   = $cat
            $tp.BackColor = $pal.CardBack
            [void]$tabStrip.TabPages.Add($tp)
        }

        $toolNameLabel.Text = '⚙️  IT Pro Mode'
        $toolDescLabel.Text = 'Select a category from the tabs or sidebar to browse tools.'
        $rpDescBox.BackColor = [System.Drawing.Color]::FromArgb(24, 28, 40)
        $rpDescBox.ForeColor = $pal.TextMid
        Populate-ProCards -Category ($script:CategoryOrder | Select-Object -First 1)
    }

    $navList.Invalidate()
    $navList.Refresh()
}

$modeToggleBtn.Add_Click({
    if (Test-TaskRunning) {
        [System.Windows.Forms.MessageBox]::Show(
            'Please wait for the current task to finish before switching modes.',
            "T3CHNRD's Windows Tool Kit",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }
    $script:IsSimpleMode = -not $script:IsSimpleMode
    $script:SelectedTask = $null
    $simpleStatusPanel.Visible = $false
    Apply-Mode
    Apply-ResponsiveLayout
})

# ============================================================
#  SECTION 13 — NAV & TAB EVENTS
# ============================================================

$navList.Add_SelectedIndexChanged({
    $item = $navList.SelectedItem
    if (-not $item -or $item.IsHeader) { return }

    if ($script:IsSimpleMode) {
        $st = $script:SimpleTasks[$item.SimpleIdx]
        if ($st) { Select-SimpleTask -TaskId $st.TaskId }
    } else {
        $cat = $item.Category
        if ($cat -and $script:TasksByCategory.ContainsKey($cat)) {
            Populate-ProCards -Category $cat
        }
    }
})

$tabStrip.Add_SelectedIndexChanged({
    $tp = $tabStrip.SelectedTab
    if ($tp -and $tp.Tag) {
        Populate-ProCards -Category $tp.Tag
    }
})

# ============================================================
#  SECTION 14 — HELPER FUNCTIONS (from your original)
# ============================================================

function Get-CategoryTabLabel {
    param([string]$Category)
    switch ($Category) {
        'Applications'          { 'Apps'     }
        'Cleanup'               { 'Cleanup'  }
        'Deployment'            { 'Deploy'   }
        'Hardware Diagnostics'  { 'Hardware' }
        'Misc/Utility'          { 'Misc'     }
        'Network Tools'         { 'Network'  }
        'OneNote & Documents'   { 'OneNote'  }
        'Repair'                { 'Repair'   }
        'Storage / Setup'       { 'Storage'  }
        'Update Tools'          { 'Updates'  }
        'Security'              { 'Security' }
        default                 { $Category  }
    }
}

function Apply-ResponsiveLayout {
    if ($bodySplit.Width -le 0 -or $mainSplit.Width -le 0) { return }

    try {
        $sidebarWidth = if ($script:IsSimpleMode) { 170 } else { 210 }
        $bodyWidth = [int]$bodySplit.Width
        if ($bodyWidth -gt 60) {
            # FIX: MARKET-01 - never increase MinSize before the splitter is known valid for the current width.
            $bodySplit.Panel1MinSize = 1
            $bodySplit.Panel2MinSize = 1
            $maxSidebar = [Math]::Max(40, $bodyWidth - 80)
            $safeSidebar = [int]([Math]::Max(40, [Math]::Min($sidebarWidth, $maxSidebar)))
            if ($safeSidebar -gt 0 -and $safeSidebar -lt ($bodyWidth - 1)) {
                $bodySplit.SplitterDistance = $safeSidebar
            }
        }
    }
    catch {
        Add-LogLine "Responsive body layout skipped: $($_.Exception.Message)"
    }

    try {
        $available = [int]$mainSplit.Width
        if ($available -le 0) { return }

        $desiredRight = if ($script:IsSimpleMode) { 350 } else { 390 }
        if ($available -lt 720) {
            $desiredRight = [Math]::Max(240, [Math]::Floor($available * 0.34))
        }

        # FIX: MARKET-01 - keep SplitContainer MinSize at 1 to avoid DPI/startup crashes.
        $mainSplit.Panel1MinSize = 1
        $mainSplit.Panel2MinSize = 1
        $splitterDistance = $available - $desiredRight
        $maxDistance = $available - 2
        $splitterDistance = [Math]::Max(1, [Math]::Min($splitterDistance, $maxDistance))

        if ($splitterDistance -gt 0 -and $splitterDistance -lt $available) {
            $mainSplit.SplitterDistance = [int]$splitterDistance
        }
    }
    catch {
        Add-LogLine "Responsive main layout skipped: $($_.Exception.Message)"
    }

    $toolDescLabel.MaximumSize = New-Object System.Drawing.Size(([Math]::Max(260, $toolHeader.Width - 32)), 0)
}

# ============================================================
#  SECTION 15 — CLOCK & KEYBOARD
# ============================================================

$clockTimer          = New-Object System.Windows.Forms.Timer
$clockTimer.Interval = 1000
$clockTimer.Add_Tick({
    try {
        # FIX: CRITICAL-03 - timer events must never surface as unhandled WinForms exceptions.
        $clockLabel.Text = (Get-Date).ToString('HH:mm:ss')
    }
    catch {
        Add-LogLine "Clock update skipped: $($_.Exception.Message)"
    }
})
$clockTimer.Start()

$form.Add_KeyDown({
    param($s, $e)
    try {
        # FIX: HIGH-15 - F5 rerun and Esc cancel are explicitly wired.
        if ($e.KeyCode -eq 'F5' -and $script:SelectedTask -and -not (Test-TaskRunning)) {
            $runBtn.PerformClick()
        }
        if ($e.KeyCode -eq 'Escape' -and (Test-TaskRunning)) {
            $cancelBtn.PerformClick()
        }
    }
    catch {
        Add-LogLine "Keyboard shortcut skipped: $($_.Exception.Message)"
    }
})

# ============================================================
#  SECTION 16 — STARTUP
# ============================================================

$form.Add_Shown({
    try {
        # FIX: HIGH-09 - theme/mode is applied at startup, not only after the first toggle.
        Refresh-TaskCatalog
        Apply-Mode   # Start in Pro mode
        Apply-ResponsiveLayout
        Add-LogLine 'T3CHNRD Windows Tool Kit v2 initialized.'
        Add-LogLine "Loaded $($script:Tasks.Count) tasks across $($script:CategoryOrder.Count) categories."
        Add-LogLine 'Running as Administrator: OK'
        Add-LogLine 'Use the mode toggle (top right) to switch to Simple Mode.'
    }
    catch {
        Add-LogLine "Startup UI initialization failed: $($_.Exception.Message)"
    }
})

$form.Add_Resize({
    try {
        # FIX: CRITICAL-03 - resize/layout errors should be logged, not crash the form.
        Apply-ResponsiveLayout
    }
    catch {
        Add-LogLine "Resize layout skipped: $($_.Exception.Message)"
    }
})

$form.Add_FormClosing({
    # FIX: LOW-14 - stop/dispose timers and GDI resources during shutdown.
    try { $clockTimer.Stop(); $clockTimer.Dispose() } catch {}
    try { $taskMonitorTimer.Stop(); $taskMonitorTimer.Dispose() } catch {}
    Stop-TaskProcessTree
    Clear-TaskProcessFiles
    try {
        if ($script:CardBorderPen) {
            $script:CardBorderPen.Dispose()
        }
    }
    catch {}
})

[void]$form.ShowDialog()
