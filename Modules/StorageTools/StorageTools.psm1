Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-TtkDriveReport {
    [CmdletBinding()]
    param()

    Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3" |
        ForEach-Object {
            $freeGb = [math]::Round($_.FreeSpace / 1GB, 2)
            $sizeGb = [math]::Round($_.Size / 1GB, 2)
            $pctFree = if ($_.Size -gt 0) { [math]::Round(($_.FreeSpace / $_.Size) * 100, 1) } else { 0 }
            [pscustomobject]@{
                Drive = $_.DeviceID
                VolumeName = $_.VolumeName
                FileSystem = $_.FileSystem
                FreeGB = $freeGb
                SizeGB = $sizeGb
                PercentFree = $pctFree
                Severity = if ($pctFree -lt 15) { 'LOW SPACE' } else { 'OK' }
            }
        }
}

function Start-TtkFileTransfer {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Source,
        [Parameter(Mandatory = $true)][string]$Destination,
        [switch]$Move,
        [string]$LogPath
    )

    if (-not (Test-Path -LiteralPath $Source)) {
        throw "Source path not found: $Source"
    }

    $null = New-Item -ItemType Directory -Force -Path $Destination
    if (-not $LogPath) {
        $LogPath = Join-Path $env:TEMP ("robocopy-{0:yyyyMMdd-HHmmss}.log" -f (Get-Date))
    }

    $copyFlags = @('/E','/R:2','/W:2','/COPY:DAT','/DCOPY:DAT','/MT:8','/TEE','/LOG+:' + $LogPath)
    if ($Move) {
        $copyFlags += '/MOVE'
    }

    $arguments = @("`"$Source`"","`"$Destination`"") + $copyFlags
    $process = Start-Process -FilePath 'robocopy.exe' -ArgumentList $arguments -Wait -PassThru -NoNewWindow
    if ($process.ExitCode -ge 8) {
        throw "Robocopy failed with exit code $($process.ExitCode). See $LogPath"
    }

    return [pscustomobject]@{
        ExitCode = $process.ExitCode
        LogPath = $LogPath
        MoveMode = [bool]$Move
    }
}

function Get-TtkDiskInventory {
    [CmdletBinding()]
    param()

    Get-Disk | Select-Object Number, FriendlyName, PartitionStyle, Size, OperationalStatus, BusType
}

function Invoke-TtkCloneDiskGuide {
    [CmdletBinding()]
    param()

    return @(
        'Disk cloning is orchestrated as a guided workflow.',
        '1. Back up important data before continuing.',
        '2. Connect the destination disk and confirm it is the correct target.',
        '3. Use manufacturer or imaging software for the actual sector copy when required.',
        '4. Validate boot order and disk health after the clone completes.'
    )
}

function Invoke-TtkNewComputerSetupChecklist {
    [CmdletBinding()]
    param()

    Get-TtkNewComputerSetupChecklistItems | ForEach-Object {
        "{0}. {1} - {2}" -f $_.Order, $_.Step, $_.Details
    }
}

function Get-TtkNewComputerSetupChecklistItems {
    [CmdletBinding()]
    param()

    return @(
        [pscustomobject]@{
            Order = 1
            Step = 'Windows updates'
            Details = 'Scan, install available Windows updates, then reboot if Windows asks for it.'
            SuggestedTool = 'Windows Update Tool'
        }
        [pscustomobject]@{
            Order = 2
            Step = 'Manufacturer updates'
            Details = 'Install BIOS, firmware, and driver updates from the official Dell, HP, Lenovo, or Framework workflow.'
            SuggestedTool = 'BIOS / Firmware / Driver Update Tool'
        }
        [pscustomobject]@{
            Order = 3
            Step = 'Core applications'
            Details = 'Install Microsoft 365 and any required support or line-of-business applications.'
            SuggestedTool = 'Install Microsoft 365'
        }
        [pscustomobject]@{
            Order = 4
            Step = 'Security baseline'
            Details = 'Verify Defender, firewall, BitLocker, Secure Boot, open ports, and basic security settings.'
            SuggestedTool = 'Security tools'
        }
        [pscustomobject]@{
            Order = 5
            Step = 'Power and recovery'
            Details = 'Set power preferences, create a restore point, and confirm recovery options are available.'
            SuggestedTool = 'Repair / Storage tools'
        }
        [pscustomobject]@{
            Order = 6
            Step = 'User data transfer'
            Details = 'Move user documents, desktop files, browser exports, and profile data from the old device or drive.'
            SuggestedTool = 'File Transfer Script'
        }
        [pscustomobject]@{
            Order = 7
            Step = 'BitLocker key backup'
            Details = 'Back up BitLocker recovery keys to external media before the device leaves your bench.'
            SuggestedTool = 'Backup BitLocker Keys'
        }
        [pscustomobject]@{
            Order = 8
            Step = 'Final validation'
            Details = 'Confirm storage, network, updates, apps, login, printing, and support tools work before handoff.'
            SuggestedTool = 'Diagnostics / Reports'
        }
    )
}

function Show-TtkNewComputerSetupChecklist {
    [CmdletBinding()]
    param()

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $items = @(Get-TtkNewComputerSetupChecklistItems)
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'New Computer Setup Checklist'
    $form.StartPosition = 'CenterScreen'
    $form.Size = New-Object System.Drawing.Size(820, 560)
    $form.MinimumSize = New-Object System.Drawing.Size(720, 500)
    $form.Font = New-Object System.Drawing.Font('Segoe UI', 10)
    $form.BackColor = [System.Drawing.Color]::FromArgb(245, 248, 255)

    $layout = New-Object System.Windows.Forms.TableLayoutPanel
    $layout.Dock = 'Fill'
    $layout.RowCount = 4
    $layout.ColumnCount = 1
    $layout.Padding = New-Object System.Windows.Forms.Padding(14)
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 120)))
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 44)))
    [void]$form.Controls.Add($layout)

    $header = New-Object System.Windows.Forms.Label
    $header.Text = "Use this checklist while setting up a new PC. Check items off as you complete them, then save a report if you want a handoff note."
    $header.AutoSize = $true
    $header.MaximumSize = New-Object System.Drawing.Size(760, 0)
    $header.ForeColor = [System.Drawing.Color]::FromArgb(31, 41, 80)
    $header.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)
    [void]$layout.Controls.Add($header, 0, 0)

    $checkList = New-Object System.Windows.Forms.CheckedListBox
    $checkList.Dock = 'Fill'
    $checkList.CheckOnClick = $true
    $checkList.BackColor = [System.Drawing.Color]::White
    $checkList.ForeColor = [System.Drawing.Color]::FromArgb(15, 23, 65)
    $checkList.BorderStyle = 'FixedSingle'
    foreach ($item in $items) {
        [void]$checkList.Items.Add(("{0}. {1} ({2})" -f $item.Order, $item.Step, $item.SuggestedTool), $false)
    }
    [void]$layout.Controls.Add($checkList, 0, 1)

    $details = New-Object System.Windows.Forms.TextBox
    $details.Multiline = $true
    $details.ReadOnly = $true
    $details.Dock = 'Fill'
    $details.ScrollBars = 'Vertical'
    $details.BackColor = [System.Drawing.Color]::White
    $details.ForeColor = [System.Drawing.Color]::FromArgb(31, 41, 80)
    $details.BorderStyle = 'FixedSingle'
    [void]$layout.Controls.Add($details, 0, 2)

    $buttonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $buttonPanel.Dock = 'Fill'
    $buttonPanel.FlowDirection = 'RightToLeft'
    [void]$layout.Controls.Add($buttonPanel, 0, 3)

    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = 'Close'
    $closeButton.Size = New-Object System.Drawing.Size(110, 32)
    $closeButton.Add_Click({ $form.Close() })
    [void]$buttonPanel.Controls.Add($closeButton)

    $saveButton = New-Object System.Windows.Forms.Button
    $saveButton.Text = 'Save Report'
    $saveButton.Size = New-Object System.Drawing.Size(120, 32)
    [void]$buttonPanel.Controls.Add($saveButton)

    $markAllButton = New-Object System.Windows.Forms.Button
    $markAllButton.Text = 'Mark All Done'
    $markAllButton.Size = New-Object System.Drawing.Size(130, 32)
    [void]$buttonPanel.Controls.Add($markAllButton)

    $updateDetails = {
        $index = $checkList.SelectedIndex
        if ($index -lt 0) { $index = 0 }
        if ($items.Count -eq 0) { return }
        $item = $items[$index]
        $details.Text = @(
            "Step $($item.Order): $($item.Step)"
            ''
            "What this means:"
            $item.Details
            ''
            "Suggested toolkit area:"
            $item.SuggestedTool
        ) -join [Environment]::NewLine
    }
    $checkList.Add_SelectedIndexChanged($updateDetails)

    $markAllButton.Add_Click({
        for ($i = 0; $i -lt $checkList.Items.Count; $i++) {
            $checkList.SetItemChecked($i, $true)
        }
    })

    $saveButton.Add_Click({
        $root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
        $logDir = Join-Path $root 'Logs'
        if (-not (Test-Path -LiteralPath $logDir)) {
            New-Item -Path $logDir -ItemType Directory -Force | Out-Null
        }
        $reportPath = Join-Path $logDir ("NewComputerSetupChecklist-{0:yyyyMMdd-HHmmss}.txt" -f (Get-Date))
        $lines = @("New Computer Setup Checklist - $(Get-Date)", '')
        for ($i = 0; $i -lt $items.Count; $i++) {
            $prefix = if ($checkList.GetItemChecked($i)) { '[DONE]' } else { '[OPEN]' }
            $lines += "{0} {1}. {2} - {3}" -f $prefix, $items[$i].Order, $items[$i].Step, $items[$i].Details
        }
        Set-Content -LiteralPath $reportPath -Value $lines -Encoding UTF8
        [System.Windows.Forms.MessageBox]::Show("Checklist report saved:`r`n$reportPath", 'Checklist Saved', 'OK', 'Information') | Out-Null
    })

    if ($checkList.Items.Count -gt 0) {
        $checkList.SelectedIndex = 0
        & $updateDetails
    }

    [void]$form.ShowDialog()

    $completed = for ($i = 0; $i -lt $items.Count; $i++) {
        if ($checkList.GetItemChecked($i)) { $items[$i].Step }
    }

    return @(
        'Opened interactive new-computer setup checklist.'
        "Completed items selected in checklist UI: $(@($completed).Count) of $($items.Count)."
    ) + (Invoke-TtkNewComputerSetupChecklist)
}

Export-ModuleMember -Function @(
    'Get-TtkDriveReport',
    'Start-TtkFileTransfer',
    'Get-TtkDiskInventory',
    'Invoke-TtkCloneDiskGuide',
    'Invoke-TtkNewComputerSetupChecklist',
    'Get-TtkNewComputerSetupChecklistItems',
    'Show-TtkNewComputerSetupChecklist'
)
