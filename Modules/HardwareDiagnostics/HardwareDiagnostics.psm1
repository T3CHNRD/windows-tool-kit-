Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Test-TtkMouseKeyboardActivity {
    [CmdletBinding()]
    param()

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Mouse & Keyboard Test'
    $form.Size = New-Object System.Drawing.Size(640, 360)
    $form.StartPosition = 'CenterScreen'
    $form.KeyPreview = $true

    $label = New-Object System.Windows.Forms.Label
    $label.Dock = 'Fill'
    $label.TextAlign = 'MiddleCenter'
    $label.Font = New-Object System.Drawing.Font('Segoe UI', 14)
    $label.Text = 'Move the mouse, click buttons, or press keys. Close the window when done.'
    $form.Controls.Add($label)

    $lastEvent = 'No input detected yet.'
    $form.Add_MouseMove({ $script:lastEvent = 'Mouse move detected.'; $label.Text = $script:lastEvent })
    $form.Add_MouseDown({ $script:lastEvent = 'Mouse click detected.'; $label.Text = $script:lastEvent })
    $form.Add_KeyDown({ param($sender, $e) $script:lastEvent = "Key detected: $($e.KeyCode)"; $label.Text = $script:lastEvent })

    [void]$form.ShowDialog()
    return $script:lastEvent
}

function Start-TtkMonitorPixelTest {
    [CmdletBinding()]
    param()

    $colors = @(
        [System.Drawing.Color]::Red,
        [System.Drawing.Color]::Lime,
        [System.Drawing.Color]::Blue,
        [System.Drawing.Color]::White,
        [System.Drawing.Color]::Black
    )

    $index = 0
    $form = New-Object System.Windows.Forms.Form
    $form.WindowState = 'Maximized'
    $form.FormBorderStyle = 'None'
    $form.TopMost = $true
    $form.BackColor = $colors[$index]
    $form.KeyPreview = $true

    $form.Add_KeyDown({
        param($sender, $e)
        if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Escape) {
            $form.Close()
            return
        }

        if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Space -or $e.KeyCode -eq [System.Windows.Forms.Keys]::Right) {
            $script:index = ($script:index + 1) % $colors.Count
            $form.BackColor = $colors[$script:index]
        }
    })

    [void]$form.ShowDialog()
    return 'Monitor test closed.'
}

Export-ModuleMember -Function @(
    'Test-TtkMouseKeyboardActivity',
    'Start-TtkMonitorPixelTest'
)
