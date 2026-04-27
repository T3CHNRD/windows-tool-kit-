Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function New-TtkStatusLabel {
    param(
        [Parameter(Mandatory = $true)][string]$Text,
        [System.Drawing.Color]$BackColor = [System.Drawing.Color]::FromArgb(126, 30, 30)
    )

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.TextAlign = 'MiddleCenter'
    $label.AutoSize = $false
    $label.Size = New-Object System.Drawing.Size(150, 34)
    $label.Margin = New-Object System.Windows.Forms.Padding(4)
    $label.BackColor = $BackColor
    $label.ForeColor = [System.Drawing.Color]::White
    $label.BorderStyle = 'FixedSingle'
    return $label
}

function Test-TtkMouseKeyboardActivity {
    [CmdletBinding()]
    param()

    $detectedColor = [System.Drawing.Color]::FromArgb(33, 132, 71)
    $missingColor = [System.Drawing.Color]::FromArgb(154, 42, 42)
    $neutralColor = [System.Drawing.Color]::FromArgb(84, 96, 118)
    $lastEvent = 'No input detected yet.'
    $detectedKeys = New-Object 'System.Collections.Generic.HashSet[string]'
    $keyButtons = @{}

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Mouse & Keyboard Test'
    $form.Size = New-Object System.Drawing.Size(1120, 760)
    $form.MinimumSize = New-Object System.Drawing.Size(980, 680)
    $form.StartPosition = 'CenterScreen'
    $form.KeyPreview = $true
    $form.Font = New-Object System.Drawing.Font('Segoe UI', 10)

    $layout = New-Object System.Windows.Forms.TableLayoutPanel
    $layout.Dock = 'Fill'
    $layout.Padding = New-Object System.Windows.Forms.Padding(14)
    $layout.ColumnCount = 1
    $layout.RowCount = 6
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$form.Controls.Add($layout)

    $title = New-Object System.Windows.Forms.Label
    $title.Text = 'Mouse and Keyboard Diagnostic'
    $title.AutoSize = $true
    $title.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 16, [System.Drawing.FontStyle]::Bold)
    [void]$layout.Controls.Add($title, 0, 0)

    $instructions = New-Object System.Windows.Forms.Label
    $instructions.Text = 'Press each key on the physical keyboard. Keys start red and turn green when detected. Move/click the mouse to light up the mouse indicators. Click Summary when finished.'
    $instructions.AutoSize = $true
    $instructions.ForeColor = $neutralColor
    $instructions.Margin = New-Object System.Windows.Forms.Padding(0, 4, 0, 10)
    [void]$layout.Controls.Add($instructions, 0, 1)

    $mousePanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $mousePanel.Dock = 'Fill'
    $mousePanel.AutoSize = $true
    $mousePanel.WrapContents = $true
    [void]$layout.Controls.Add($mousePanel, 0, 2)

    $mouseMoveLabel = New-TtkStatusLabel -Text 'Mouse Move'
    $leftClickLabel = New-TtkStatusLabel -Text 'Left Click'
    $rightClickLabel = New-TtkStatusLabel -Text 'Right Click'
    $middleClickLabel = New-TtkStatusLabel -Text 'Middle Click'
    foreach ($label in @($mouseMoveLabel, $leftClickLabel, $rightClickLabel, $middleClickLabel)) {
        [void]$mousePanel.Controls.Add($label)
    }

    $keyboardPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $keyboardPanel.Dock = 'Fill'
    $keyboardPanel.AutoScroll = $true
    $keyboardPanel.ColumnCount = 1
    $keyboardPanel.RowCount = 7
    [void]$layout.Controls.Add($keyboardPanel, 0, 3)

    function Add-KeyRow {
        param(
            [Parameter(Mandatory = $true)][object[]]$Keys
        )

        $row = New-Object System.Windows.Forms.FlowLayoutPanel
        $row.Dock = 'Top'
        $row.AutoSize = $true
        $row.WrapContents = $false
        $row.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 4)
        [void]$keyboardPanel.Controls.Add($row)

        foreach ($keyDef in $Keys) {
            $label = [string]$keyDef[0]
            $keyCode = [System.Windows.Forms.Keys]$keyDef[1]
            $width = if ($keyDef.Count -gt 2) { [int]$keyDef[2] } else { 56 }

            $button = New-Object System.Windows.Forms.Button
            $button.Text = $label
            $button.Tag = $keyCode
            $button.Width = $width
            $button.Height = 38
            $button.Margin = New-Object System.Windows.Forms.Padding(3)
            $button.FlatStyle = 'Flat'
            $button.BackColor = $missingColor
            $button.ForeColor = [System.Drawing.Color]::White
            [void]$row.Controls.Add($button)
            $keyText = [string]$keyCode
            if (-not $keyButtons.ContainsKey($keyText)) {
                $keyButtons[$keyText] = New-Object 'System.Collections.Generic.List[System.Windows.Forms.Button]'
            }
            $keyButtons[$keyText].Add($button)
        }
    }

    Add-KeyRow -Keys @(
        @('Esc', 'Escape'), @('F1', 'F1'), @('F2', 'F2'), @('F3', 'F3'), @('F4', 'F4'), @('F5', 'F5'),
        @('F6', 'F6'), @('F7', 'F7'), @('F8', 'F8'), @('F9', 'F9'), @('F10', 'F10'), @('F11', 'F11'), @('F12', 'F12')
    )
    Add-KeyRow -Keys @(
        @('`~', 'Oemtilde'), @('1', 'D1'), @('2', 'D2'), @('3', 'D3'), @('4', 'D4'), @('5', 'D5'), @('6', 'D6'),
        @('7', 'D7'), @('8', 'D8'), @('9', 'D9'), @('0', 'D0'), @('-_', 'OemMinus'), @('=+', 'Oemplus'), @('Backspace', 'Back', 110)
    )
    Add-KeyRow -Keys @(
        @('Tab', 'Tab', 82), @('Q', 'Q'), @('W', 'W'), @('E', 'E'), @('R', 'R'), @('T', 'T'), @('Y', 'Y'),
        @('U', 'U'), @('I', 'I'), @('O', 'O'), @('P', 'P'), @('[{', 'OemOpenBrackets'), @(']}', 'OemCloseBrackets'), @('\|', 'OemPipe', 78)
    )
    Add-KeyRow -Keys @(
        @('Caps', 'Capital', 94), @('A', 'A'), @('S', 'S'), @('D', 'D'), @('F', 'F'), @('G', 'G'), @('H', 'H'),
        @('J', 'J'), @('K', 'K'), @('L', 'L'), @(';:', 'OemSemicolon'), @("'`"", 'OemQuotes'), @('Enter', 'Return', 106)
    )
    Add-KeyRow -Keys @(
        @('Shift', 'ShiftKey', 118), @('Z', 'Z'), @('X', 'X'), @('C', 'C'), @('V', 'V'), @('B', 'B'), @('N', 'N'),
        @('M', 'M'), @(',<', 'Oemcomma'), @('.>', 'OemPeriod'), @('/?', 'OemQuestion'), @('Shift', 'ShiftKey', 118)
    )
    Add-KeyRow -Keys @(
        @('Ctrl', 'ControlKey', 82), @('Win', 'LWin', 72), @('Alt', 'Menu', 72), @('Space', 'Space', 300),
        @('Alt', 'Menu', 72), @('Ctrl', 'ControlKey', 82), @('Left', 'Left'), @('Up', 'Up'), @('Down', 'Down'), @('Right', 'Right')
    )

    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Text = "Last input: $lastEvent"
    $statusLabel.AutoSize = $true
    $statusLabel.Margin = New-Object System.Windows.Forms.Padding(0, 8, 0, 8)
    [void]$layout.Controls.Add($statusLabel, 0, 4)

    $buttonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $buttonPanel.Dock = 'Fill'
    $buttonPanel.FlowDirection = 'RightToLeft'
    $buttonPanel.AutoSize = $true
    [void]$layout.Controls.Add($buttonPanel, 0, 5)

    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = 'Close'
    $closeButton.Size = New-Object System.Drawing.Size(120, 36)
    [void]$buttonPanel.Controls.Add($closeButton)

    $summaryButton = New-Object System.Windows.Forms.Button
    $summaryButton.Text = 'Summary'
    $summaryButton.Size = New-Object System.Drawing.Size(120, 36)
    [void]$buttonPanel.Controls.Add($summaryButton)

    $resetButton = New-Object System.Windows.Forms.Button
    $resetButton.Text = 'Reset'
    $resetButton.Size = New-Object System.Drawing.Size(120, 36)
    [void]$buttonPanel.Controls.Add($resetButton)

    function Set-InputEvent {
        param(
            [Parameter(Mandatory = $true)][string]$Message
        )

        $script:ttkLastInputUtc = [DateTime]::UtcNow
        $statusLabel.Text = "Last input: $Message"
        $form.Text = "Mouse & Keyboard Test - $Message"
    }

    function Mark-KeyDetected {
        param(
            [Parameter(Mandatory = $true)][System.Windows.Forms.Keys]$KeyCode
        )

        $keyText = [string]$KeyCode
        [void]$detectedKeys.Add($keyText)
        if ($keyButtons.ContainsKey($keyText)) {
            foreach ($button in $keyButtons[$keyText]) {
                $button.BackColor = $detectedColor
            }
        }
        Set-InputEvent -Message "Key detected: $keyText"
    }

    $form.Add_MouseMove({
        $mouseMoveLabel.BackColor = $detectedColor
        Set-InputEvent -Message 'Mouse movement detected'
    })

    $form.Add_MouseDown({
        param($sender, $e)
        switch ($e.Button) {
            'Left' { $leftClickLabel.BackColor = $detectedColor }
            'Right' { $rightClickLabel.BackColor = $detectedColor }
            'Middle' { $middleClickLabel.BackColor = $detectedColor }
        }
        Set-InputEvent -Message "Mouse $($e.Button) click detected"
    })

    $form.Add_KeyDown({
        param($sender, $e)
        Mark-KeyDetected -KeyCode $e.KeyCode
    })

    $resetButton.Add_Click({
        $detectedKeys.Clear()
        foreach ($buttonList in $keyButtons.Values) {
            foreach ($button in $buttonList) {
                $button.BackColor = $missingColor
            }
        }
        foreach ($label in @($mouseMoveLabel, $leftClickLabel, $rightClickLabel, $middleClickLabel)) {
            $label.BackColor = $missingColor
        }
        Set-InputEvent -Message 'Test reset'
    })

    $summaryButton.Add_Click({
        $total = $keyButtons.Count
        $detected = ($keyButtons.Keys | Where-Object { $detectedKeys.Contains($_) }).Count
        $missing = $total - $detected
        $message = "Keyboard keys detected: $detected of $total`r`nKeys still red/not detected: $missing`r`n`r`nGreen means detected. Red means not detected during this test."
        [System.Windows.Forms.MessageBox]::Show(
            $message,
            'Mouse & Keyboard Test Summary',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    })

    $closeButton.Add_Click({ $form.Close() })
    [void]$form.ShowDialog()

    $detectedCount = ($keyButtons.Keys | Where-Object { $detectedKeys.Contains($_) }).Count
    return "Mouse/keyboard test closed. Keyboard keys detected: $detectedCount of $($keyButtons.Count)."
}

function Start-TtkMonitorPixelTest {
    [CmdletBinding()]
    param()

    $colors = @(
        @{ Name = 'Red'; Value = [System.Drawing.Color]::Red },
        @{ Name = 'Green'; Value = [System.Drawing.Color]::Lime },
        @{ Name = 'Blue'; Value = [System.Drawing.Color]::Blue },
        @{ Name = 'White'; Value = [System.Drawing.Color]::White },
        @{ Name = 'Black'; Value = [System.Drawing.Color]::Black }
    )

    $launcher = New-Object System.Windows.Forms.Form
    $launcher.Text = 'Monitor Dead Pixel Test'
    $launcher.Size = New-Object System.Drawing.Size(720, 430)
    $launcher.StartPosition = 'CenterScreen'
    $launcher.Font = New-Object System.Drawing.Font('Segoe UI', 10)

    $layout = New-Object System.Windows.Forms.TableLayoutPanel
    $layout.Dock = 'Fill'
    $layout.Padding = New-Object System.Windows.Forms.Padding(14)
    $layout.ColumnCount = 1
    $layout.RowCount = 5
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$launcher.Controls.Add($layout)

    $title = New-Object System.Windows.Forms.Label
    $title.Text = 'Monitor Dead Pixel Tester and Stuck Pixel Wake Tool'
    $title.AutoSize = $true
    $title.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 14, [System.Drawing.FontStyle]::Bold)
    [void]$layout.Controls.Add($title, 0, 0)

    $instructions = New-Object System.Windows.Forms.TextBox
    $instructions.Multiline = $true
    $instructions.ReadOnly = $true
    $instructions.Dock = 'Fill'
    $instructions.Height = 150
    $instructions.Text = @(
        'How to use:',
        '1. Pick the monitor to test.',
        '2. Click Full Screen Color Test. Use Space or Right Arrow to change colors.',
        '3. Look for pixels that stay black, white, or the wrong color.',
        '4. Press F during the full-screen test to toggle a flashing stuck-pixel wake attempt.',
        '5. Use arrow keys to move the flashing square, +/- to resize it, and ESC to exit.',
        '',
        'Note: flashing can sometimes wake stuck pixels, but it cannot repair physically dead pixels.'
    ) -join [Environment]::NewLine
    [void]$layout.Controls.Add($instructions, 0, 1)

    $screenPicker = New-Object System.Windows.Forms.ComboBox
    $screenPicker.DropDownStyle = 'DropDownList'
    $screenPicker.Dock = 'Top'
    foreach ($screen in [System.Windows.Forms.Screen]::AllScreens) {
        [void]$screenPicker.Items.Add(('{0}: {1}x{2} at {3},{4}' -f $screen.DeviceName, $screen.Bounds.Width, $screen.Bounds.Height, $screen.Bounds.X, $screen.Bounds.Y))
    }
    if ($screenPicker.Items.Count -gt 0) {
        $screenPicker.SelectedIndex = 0
    }
    [void]$layout.Controls.Add($screenPicker, 0, 2)

    $reportBox = New-Object System.Windows.Forms.TextBox
    $reportBox.Multiline = $true
    $reportBox.Dock = 'Fill'
    $reportBox.ScrollBars = 'Vertical'
    $reportBox.Text = 'Notes: Enter any dead/stuck pixel locations here before saving a report.'
    [void]$layout.Controls.Add($reportBox, 0, 3)

    $buttonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $buttonPanel.Dock = 'Fill'
    $buttonPanel.FlowDirection = 'RightToLeft'
    $buttonPanel.AutoSize = $true
    [void]$layout.Controls.Add($buttonPanel, 0, 4)

    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = 'Close'
    $closeButton.Size = New-Object System.Drawing.Size(120, 36)
    [void]$buttonPanel.Controls.Add($closeButton)

    $saveReportButton = New-Object System.Windows.Forms.Button
    $saveReportButton.Text = 'Save Report'
    $saveReportButton.Size = New-Object System.Drawing.Size(120, 36)
    [void]$buttonPanel.Controls.Add($saveReportButton)

    $fullScreenButton = New-Object System.Windows.Forms.Button
    $fullScreenButton.Text = 'Full Screen Color Test'
    $fullScreenButton.Size = New-Object System.Drawing.Size(170, 36)
    [void]$buttonPanel.Controls.Add($fullScreenButton)

    function Start-FullScreenPixelTest {
        $selectedScreen = [System.Windows.Forms.Screen]::AllScreens[$screenPicker.SelectedIndex]
        $index = 0
        $flashIndex = 0
        $flashEnabled = $false
        $flashSize = 120

        $form = New-Object System.Windows.Forms.Form
        $form.Text = 'Dead Pixel Test - ESC exits'
        $form.FormBorderStyle = 'None'
        $form.StartPosition = 'Manual'
        $form.Bounds = $selectedScreen.Bounds
        $form.TopMost = $true
        $form.KeyPreview = $true
        $form.BackColor = $colors[$index].Value

        $help = New-Object System.Windows.Forms.Label
        $help.AutoSize = $true
        $help.BackColor = [System.Drawing.Color]::FromArgb(180, 0, 0, 0)
        $help.ForeColor = [System.Drawing.Color]::White
        $help.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 12, [System.Drawing.FontStyle]::Bold)
        $help.Text = 'ESC exit | Space/Right next color | F toggle stuck-pixel flash | Arrows move flash box | +/- resize'
        $help.Location = New-Object System.Drawing.Point(18, 18)
        [void]$form.Controls.Add($help)

        $flashPanel = New-Object System.Windows.Forms.Panel
        $flashPanel.Size = New-Object System.Drawing.Size($flashSize, $flashSize)
        $flashPanel.Location = New-Object System.Drawing.Point(
            [Math]::Max(0, [int](($form.Width - $flashSize) / 2)),
            [Math]::Max(0, [int](($form.Height - $flashSize) / 2))
        )
        $flashPanel.Visible = $false
        [void]$form.Controls.Add($flashPanel)
        $flashPanel.BringToFront()
        $help.BringToFront()

        $timer = New-Object System.Windows.Forms.Timer
        $timer.Interval = 80
        $timer.Add_Tick({
            if (-not $flashEnabled) {
                return
            }
            $flashIndex = ($flashIndex + 1) % $colors.Count
            $flashPanel.BackColor = $colors[$flashIndex].Value
        })

        function Move-FlashPanel {
            param([int]$DeltaX, [int]$DeltaY)
            $newX = [Math]::Max(0, [Math]::Min(($form.Width - $flashPanel.Width), ($flashPanel.Left + $DeltaX)))
            $newY = [Math]::Max(0, [Math]::Min(($form.Height - $flashPanel.Height), ($flashPanel.Top + $DeltaY)))
            $flashPanel.Location = New-Object System.Drawing.Point($newX, $newY)
        }

        $form.Add_KeyDown({
            param($sender, $e)
            switch ($e.KeyCode) {
                'Escape' {
                    $timer.Stop()
                    $form.Close()
                }
                { $_ -eq [System.Windows.Forms.Keys]::Space -or $_ -eq [System.Windows.Forms.Keys]::Right } {
                    $index = ($index + 1) % $colors.Count
                    $form.BackColor = $colors[$index].Value
                }
                'F' {
                    $flashEnabled = -not $flashEnabled
                    $flashPanel.Visible = $flashEnabled
                    if ($flashEnabled) { $timer.Start() } else { $timer.Stop() }
                }
                'Left' { Move-FlashPanel -DeltaX -20 -DeltaY 0 }
                'Up' { Move-FlashPanel -DeltaX 0 -DeltaY -20 }
                'Down' { Move-FlashPanel -DeltaX 0 -DeltaY 20 }
                'Oemplus' {
                    $flashSize = [Math]::Min(400, ($flashSize + 20))
                    $flashPanel.Size = New-Object System.Drawing.Size($flashSize, $flashSize)
                }
                'Add' {
                    $flashSize = [Math]::Min(400, ($flashSize + 20))
                    $flashPanel.Size = New-Object System.Drawing.Size($flashSize, $flashSize)
                }
                'OemMinus' {
                    $flashSize = [Math]::Max(40, ($flashSize - 20))
                    $flashPanel.Size = New-Object System.Drawing.Size($flashSize, $flashSize)
                }
                'Subtract' {
                    $flashSize = [Math]::Max(40, ($flashSize - 20))
                    $flashPanel.Size = New-Object System.Drawing.Size($flashSize, $flashSize)
                }
            }
        })

        [void]$form.ShowDialog($launcher)
        $timer.Dispose()
    }

    $fullScreenButton.Add_Click({ Start-FullScreenPixelTest })

    $saveReportButton.Add_Click({
        $root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
        $logDir = Join-Path $root 'Logs'
        if (-not (Test-Path -LiteralPath $logDir)) {
            New-Item -Path $logDir -ItemType Directory -Force | Out-Null
        }
        $reportPath = Join-Path $logDir ("MonitorPixelReport-{0}.txt" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
        $report = @(
            "Monitor dead pixel report - $(Get-Date)",
            "Screen: $($screenPicker.SelectedItem)",
            '',
            $reportBox.Text
        )
        Set-Content -LiteralPath $reportPath -Value $report -Encoding UTF8
        [System.Windows.Forms.MessageBox]::Show(
            "Saved report:`r`n$reportPath",
            'Monitor Dead Pixel Test',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    })

    $closeButton.Add_Click({ $launcher.Close() })
    [void]$launcher.ShowDialog()
    return 'Monitor dead pixel test closed.'
}

Export-ModuleMember -Function @(
    'Test-TtkMouseKeyboardActivity',
    'Start-TtkMonitorPixelTest'
)
