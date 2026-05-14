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

function Get-TtkExternalTransferDrives {
    [CmdletBinding()]
    param()

    $usbLetters = @{}
    try {
        Get-Disk -ErrorAction SilentlyContinue |
            Where-Object { $_.BusType -eq 'USB' } |
            Get-Partition -ErrorAction SilentlyContinue |
            Where-Object { $_.DriveLetter } |
            ForEach-Object { $usbLetters["$($_.DriveLetter):"] = $true }
    }
    catch {
        # Keep detection best-effort. Win32_LogicalDisk below still catches removable media.
    }

    Get-CimInstance -ClassName Win32_LogicalDisk -ErrorAction SilentlyContinue |
        Where-Object { $_.DeviceID -and (($_.DriveType -eq 2) -or $usbLetters.ContainsKey($_.DeviceID)) } |
        ForEach-Object {
            [pscustomobject]@{
                Drive       = $_.DeviceID
                Label       = $_.VolumeName
                FileSystem  = $_.FileSystem
                SizeGB      = if ($_.Size) { [math]::Round($_.Size / 1GB, 2) } else { 0 }
                FreeGB      = if ($_.FreeSpace) { [math]::Round($_.FreeSpace / 1GB, 2) } else { 0 }
                DisplayName = '{0}\  {1}  Free: {2} GB  Size: {3} GB' -f $_.DeviceID, $(if ($_.VolumeName) { "($($_.VolumeName))" } else { '(External/Removable)' }), $(if ($_.FreeSpace) { [math]::Round($_.FreeSpace / 1GB, 2) } else { 0 }), $(if ($_.Size) { [math]::Round($_.Size / 1GB, 2) } else { 0 })
            }
        }
}

function Show-TtkDataTransferWizard {
    [CmdletBinding()]
    param(
        [string]$DefaultDestination = (Join-Path $env:USERPROFILE 'TransferredData')
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $result = [pscustomobject]@{
        Completed   = $false
        Cancelled   = $false
        Source      = $null
        Destination = $null
        LogPath     = $null
        ExitCode    = $null
        Summary     = 'No transfer was started.'
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Data Transfer Wizard'
    $form.StartPosition = 'CenterScreen'
    $form.Size = New-Object System.Drawing.Size(780, 560)
    $form.MinimumSize = New-Object System.Drawing.Size(720, 520)
    $form.Font = New-Object System.Drawing.Font('Segoe UI', 10)
    $form.BackColor = [System.Drawing.Color]::FromArgb(245, 248, 255)

    $layout = New-Object System.Windows.Forms.TableLayoutPanel
    $layout.Dock = 'Fill'
    $layout.RowCount = 7
    $layout.ColumnCount = 1
    $layout.Padding = New-Object System.Windows.Forms.Padding(14)
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 130)))
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 44)))
    [void]$form.Controls.Add($layout)

    $header = New-Object System.Windows.Forms.Label
    $header.Text = "Transfer files from an old PC, external drive, or network folder. The wizard checks for attached external drives first, but you can browse to any source folder."
    $header.AutoSize = $true
    $header.MaximumSize = New-Object System.Drawing.Size(730, 0)
    $header.ForeColor = [System.Drawing.Color]::FromArgb(31, 41, 80)
    $header.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)
    [void]$layout.Controls.Add($header, 0, 0)

    $driveList = New-Object System.Windows.Forms.ListBox
    $driveList.Dock = 'Fill'
    $driveList.DisplayMember = 'DisplayName'
    $driveList.BackColor = [System.Drawing.Color]::White
    [void]$layout.Controls.Add($driveList, 0, 1)

    $sourcePanel = New-Object System.Windows.Forms.TableLayoutPanel
    $sourcePanel.Dock = 'Top'
    $sourcePanel.ColumnCount = 3
    $sourcePanel.RowCount = 1
    [void]$sourcePanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 95)))
    [void]$sourcePanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$sourcePanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 110)))
    [void]$layout.Controls.Add($sourcePanel, 0, 2)

    $sourceLabel = New-Object System.Windows.Forms.Label
    $sourceLabel.Text = 'Source:'
    $sourceLabel.Dock = 'Fill'
    $sourceLabel.TextAlign = 'MiddleLeft'
    [void]$sourcePanel.Controls.Add($sourceLabel, 0, 0)

    $sourceBox = New-Object System.Windows.Forms.TextBox
    $sourceBox.Dock = 'Fill'
    [void]$sourcePanel.Controls.Add($sourceBox, 1, 0)

    $browseSource = New-Object System.Windows.Forms.Button
    $browseSource.Text = 'Browse...'
    $browseSource.Dock = 'Fill'
    [void]$sourcePanel.Controls.Add($browseSource, 2, 0)

    $destPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $destPanel.Dock = 'Top'
    $destPanel.ColumnCount = 3
    $destPanel.RowCount = 1
    [void]$destPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 95)))
    [void]$destPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$destPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 110)))
    [void]$layout.Controls.Add($destPanel, 0, 3)

    $destLabel = New-Object System.Windows.Forms.Label
    $destLabel.Text = 'Destination:'
    $destLabel.Dock = 'Fill'
    $destLabel.TextAlign = 'MiddleLeft'
    [void]$destPanel.Controls.Add($destLabel, 0, 0)

    $destBox = New-Object System.Windows.Forms.TextBox
    $destBox.Dock = 'Fill'
    $destBox.Text = $DefaultDestination
    [void]$destPanel.Controls.Add($destBox, 1, 0)

    $browseDest = New-Object System.Windows.Forms.Button
    $browseDest.Text = 'Browse...'
    $browseDest.Dock = 'Fill'
    [void]$destPanel.Controls.Add($browseDest, 2, 0)

    $moveCheck = New-Object System.Windows.Forms.CheckBox
    $moveCheck.Text = 'Move files instead of copying them (advanced; copy is safer)'
    $moveCheck.AutoSize = $true
    $moveCheck.Margin = New-Object System.Windows.Forms.Padding(0, 8, 0, 8)
    [void]$layout.Controls.Add($moveCheck, 0, 4)

    $logBox = New-Object System.Windows.Forms.TextBox
    $logBox.Multiline = $true
    $logBox.ReadOnly = $true
    $logBox.ScrollBars = 'Vertical'
    $logBox.Dock = 'Fill'
    $logBox.BackColor = [System.Drawing.Color]::White
    [void]$layout.Controls.Add($logBox, 0, 5)

    $buttons = New-Object System.Windows.Forms.FlowLayoutPanel
    $buttons.Dock = 'Fill'
    $buttons.FlowDirection = 'RightToLeft'
    [void]$layout.Controls.Add($buttons, 0, 6)

    $close = New-Object System.Windows.Forms.Button
    $close.Text = 'Close'
    $close.Size = New-Object System.Drawing.Size(110, 32)
    $close.Add_Click({ $result.Cancelled = -not $result.Completed; $form.Close() })
    [void]$buttons.Controls.Add($close)

    $start = New-Object System.Windows.Forms.Button
    $start.Text = 'Start Transfer'
    $start.Size = New-Object System.Drawing.Size(140, 32)
    [void]$buttons.Controls.Add($start)

    $refresh = New-Object System.Windows.Forms.Button
    $refresh.Text = 'Refresh Drives'
    $refresh.Size = New-Object System.Drawing.Size(130, 32)
    [void]$buttons.Controls.Add($refresh)

    $appendLog = {
        param([string]$Line)
        $logBox.AppendText(("[{0:HH:mm:ss}] {1}{2}" -f (Get-Date), $Line, [Environment]::NewLine))
    }

    $loadDrives = {
        $driveList.Items.Clear()
        $drives = @(Get-TtkExternalTransferDrives)
        foreach ($drive in $drives) { [void]$driveList.Items.Add($drive) }
        if ($drives.Count -eq 0) {
            & $appendLog 'No external/removable drive detected. Plug in the old drive, click Refresh Drives, or use Browse to pick a network/source folder.'
        }
        else {
            & $appendLog ("Detected {0} external/removable source candidate(s)." -f $drives.Count)
            $driveList.SelectedIndex = 0
            $sourceBox.Text = $drives[0].Drive + '\'
        }
    }

    $driveList.Add_SelectedIndexChanged({
        if ($driveList.SelectedItem) {
            $sourceBox.Text = $driveList.SelectedItem.Drive + '\'
        }
    })

    $browseFolder = {
        param([System.Windows.Forms.TextBox]$TargetBox)
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = 'Choose a folder'
        if ($TargetBox.Text -and (Test-Path -LiteralPath $TargetBox.Text)) {
            $dlg.SelectedPath = $TargetBox.Text
        }
        if ($dlg.ShowDialog($form) -eq 'OK') {
            $TargetBox.Text = $dlg.SelectedPath
        }
    }

    $browseSource.Add_Click({ & $browseFolder $sourceBox })
    $browseDest.Add_Click({ & $browseFolder $destBox })
    $refresh.Add_Click($loadDrives)

    $start.Add_Click({
        if (-not (Test-Path -LiteralPath $sourceBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show('Choose a valid source folder first.', 'Data Transfer Wizard', 'OK', 'Warning') | Out-Null
            return
        }
        if ([string]::IsNullOrWhiteSpace($destBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show('Choose a destination folder first.', 'Data Transfer Wizard', 'OK', 'Warning') | Out-Null
            return
        }

        $root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
        $logDir = Join-Path $root 'Logs'
        if (-not (Test-Path -LiteralPath $logDir)) { New-Item -Path $logDir -ItemType Directory -Force | Out-Null }
        $transferLog = Join-Path $logDir ("DataTransfer-{0:yyyyMMdd-HHmmss}.log" -f (Get-Date))

        $start.Enabled = $false
        & $appendLog "Starting robocopy transfer."
        & $appendLog "Source: $($sourceBox.Text)"
        & $appendLog "Destination: $($destBox.Text)"
        [System.Windows.Forms.Application]::DoEvents()

        try {
            $copyResult = Start-TtkFileTransfer -Source $sourceBox.Text -Destination $destBox.Text -Move:($moveCheck.Checked) -LogPath $transferLog
            $result.Completed = $true
            $result.Source = $sourceBox.Text
            $result.Destination = $destBox.Text
            $result.LogPath = $copyResult.LogPath
            $result.ExitCode = $copyResult.ExitCode
            $result.Summary = "Robocopy completed with exit code $($copyResult.ExitCode). Copied/skipped/failed details are in $($copyResult.LogPath)."
            & $appendLog $result.Summary
            [System.Windows.Forms.MessageBox]::Show($result.Summary, 'Transfer Complete', 'OK', 'Information') | Out-Null
        }
        catch {
            $result.Completed = $false
            $result.LogPath = $transferLog
            $result.Summary = $_.Exception.Message
            & $appendLog "Transfer failed: $($result.Summary)"
            [System.Windows.Forms.MessageBox]::Show($result.Summary, 'Transfer Failed', 'OK', 'Error') | Out-Null
        }
        finally {
            $start.Enabled = $true
        }
    })

    & $loadDrives
    [void]$form.ShowDialog()
    return $result
}

function Get-TtkDiskInventory {
    [CmdletBinding()]
    param()

    Get-Disk | Select-Object Number, FriendlyName, PartitionStyle, Size, OperationalStatus, BusType
}

function Invoke-TtkCloneDiskGuide {
    [CmdletBinding()]
    param()

    return Show-TtkCloneDiskWizard
}

function Show-TtkCloneDiskWizard {
    [CmdletBinding()]
    param()

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
    $logDir = Join-Path $root 'Logs'
    if (-not (Test-Path -LiteralPath $logDir)) {
        New-Item -Path $logDir -ItemType Directory -Force | Out-Null
    }

    $result = [pscustomobject]@{
        ReportPath = $null
        SourceDisk = $null
        DestinationDisk = $null
        OpenedDiskManagement = $false
        OpenedBackup = $false
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Clone Disk Wizard'
    $form.StartPosition = 'CenterScreen'
    $form.Size = New-Object System.Drawing.Size(900, 650)
    $form.MinimumSize = New-Object System.Drawing.Size(820, 560)
    $form.Font = New-Object System.Drawing.Font('Segoe UI', 10)
    $form.BackColor = [System.Drawing.Color]::FromArgb(245, 248, 255)

    $layout = New-Object System.Windows.Forms.TableLayoutPanel
    $layout.Dock = 'Fill'
    $layout.RowCount = 6
    $layout.ColumnCount = 1
    $layout.Padding = New-Object System.Windows.Forms.Padding(14)
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 95)))
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 55)))
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 42)))
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 42)))
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 45)))
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 50)))
    [void]$form.Controls.Add($layout)

    $intro = New-Object System.Windows.Forms.TextBox
    $intro.Multiline = $true
    $intro.ReadOnly = $true
    $intro.BorderStyle = 'FixedSingle'
    $intro.Dock = 'Fill'
    $intro.BackColor = [System.Drawing.Color]::White
    $intro.Text = "This wizard helps you plan a disk clone safely. It does not silently wipe or clone disks.`r`n`r`nChoose a SOURCE disk and a DESTINATION disk, save the plan, then launch Disk Management, Windows Backup, or your trusted imaging software to perform the actual sector copy. Always back up important data first."
    [void]$layout.Controls.Add($intro, 0, 0)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = 'Fill'
    $grid.ReadOnly = $true
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.AutoSizeColumnsMode = 'Fill'
    $grid.SelectionMode = 'FullRowSelect'
    [void]$layout.Controls.Add($grid, 0, 1)

    $sourcePanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $sourcePanel.Dock = 'Fill'
    $sourcePanel.FlowDirection = 'LeftToRight'
    $sourcePanel.WrapContents = $false
    [void]$layout.Controls.Add($sourcePanel, 0, 2)

    $sourceLabel = New-Object System.Windows.Forms.Label
    $sourceLabel.Text = 'Source disk to copy from:'
    $sourceLabel.Width = 180
    $sourceLabel.TextAlign = 'MiddleLeft'
    [void]$sourcePanel.Controls.Add($sourceLabel)

    $sourceCombo = New-Object System.Windows.Forms.ComboBox
    $sourceCombo.DropDownStyle = 'DropDownList'
    $sourceCombo.Width = 650
    [void]$sourcePanel.Controls.Add($sourceCombo)

    $destPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $destPanel.Dock = 'Fill'
    $destPanel.FlowDirection = 'LeftToRight'
    $destPanel.WrapContents = $false
    [void]$layout.Controls.Add($destPanel, 0, 3)

    $destLabel = New-Object System.Windows.Forms.Label
    $destLabel.Text = 'Destination disk to overwrite:'
    $destLabel.Width = 180
    $destLabel.TextAlign = 'MiddleLeft'
    [void]$destPanel.Controls.Add($destLabel)

    $destCombo = New-Object System.Windows.Forms.ComboBox
    $destCombo.DropDownStyle = 'DropDownList'
    $destCombo.Width = 650
    [void]$destPanel.Controls.Add($destCombo)

    $details = New-Object System.Windows.Forms.TextBox
    $details.Multiline = $true
    $details.ReadOnly = $true
    $details.ScrollBars = 'Vertical'
    $details.Dock = 'Fill'
    $details.Text = @(
        'Clone workflow:',
        '1. Confirm the source disk is the disk you want copied.',
        '2. Confirm the destination disk is the disk that may be overwritten.',
        '3. Save this clone plan report.',
        '4. Use Disk Management, Windows Backup/System Image, or trusted vendor imaging software to perform the clone.',
        '5. After cloning, validate boot order and disk health before wiping the old disk.',
        '',
        'Safety note: a true clone overwrites the destination disk. This toolkit intentionally does not auto-run destructive clone commands without a dedicated imaging engine and explicit operator confirmation.'
    ) -join [Environment]::NewLine
    [void]$layout.Controls.Add($details, 0, 4)

    $buttonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $buttonPanel.Dock = 'Fill'
    $buttonPanel.FlowDirection = 'RightToLeft'
    [void]$layout.Controls.Add($buttonPanel, 0, 5)

    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = 'Close'
    $closeButton.Width = 110
    [void]$buttonPanel.Controls.Add($closeButton)

    $saveButton = New-Object System.Windows.Forms.Button
    $saveButton.Text = 'Save Clone Plan'
    $saveButton.Width = 150
    [void]$buttonPanel.Controls.Add($saveButton)

    $backupButton = New-Object System.Windows.Forms.Button
    $backupButton.Text = 'Open Windows Backup'
    $backupButton.Width = 170
    [void]$buttonPanel.Controls.Add($backupButton)

    $diskMgmtButton = New-Object System.Windows.Forms.Button
    $diskMgmtButton.Text = 'Open Disk Management'
    $diskMgmtButton.Width = 180
    [void]$buttonPanel.Controls.Add($diskMgmtButton)

    $refreshButton = New-Object System.Windows.Forms.Button
    $refreshButton.Text = 'Refresh Disks'
    $refreshButton.Width = 130
    [void]$buttonPanel.Controls.Add($refreshButton)

    $loadDisks = {
        $sourceCombo.Items.Clear()
        $destCombo.Items.Clear()
        try {
            $disks = @(Get-TtkDiskInventory | ForEach-Object {
                    [pscustomobject]@{
                        Number = $_.Number
                        Name = $_.FriendlyName
                        BusType = $_.BusType
                        PartitionStyle = $_.PartitionStyle
                        SizeGB = [math]::Round($_.Size / 1GB, 2)
                        Status = ($_.OperationalStatus -join ', ')
                    }
                })
            $grid.DataSource = $disks
            foreach ($disk in $disks) {
                $display = 'Disk {0}: {1} | {2} | {3} GB | {4}' -f $disk.Number, $disk.Name, $disk.BusType, $disk.SizeGB, $disk.Status
                [void]$sourceCombo.Items.Add($display)
                [void]$destCombo.Items.Add($display)
            }
            if ($sourceCombo.Items.Count -gt 0) { $sourceCombo.SelectedIndex = 0 }
            if ($destCombo.Items.Count -gt 1) { $destCombo.SelectedIndex = 1 }
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Unable to read disk inventory:`r`n$($_.Exception.Message)", 'Disk Inventory', 'OK', 'Warning') | Out-Null
        }
    }

    $refreshButton.Add_Click($loadDisks)
    $diskMgmtButton.Add_Click({
        Start-Process -FilePath 'diskmgmt.msc'
        $result.OpenedDiskManagement = $true
    })
    $backupButton.Add_Click({
        try {
            Start-Process -FilePath 'control.exe' -ArgumentList '/name Microsoft.BackupAndRestore'
        }
        catch {
            Start-Process -FilePath 'sdclt.exe'
        }
        $result.OpenedBackup = $true
    })
    $saveButton.Add_Click({
        if ($sourceCombo.SelectedIndex -lt 0 -or $destCombo.SelectedIndex -lt 0) {
            [System.Windows.Forms.MessageBox]::Show('Select both a source disk and a destination disk first.', 'Clone Plan', 'OK', 'Warning') | Out-Null
            return
        }
        if ($sourceCombo.SelectedItem -eq $destCombo.SelectedItem) {
            [System.Windows.Forms.MessageBox]::Show('Source and destination cannot be the same disk.', 'Clone Plan', 'OK', 'Warning') | Out-Null
            return
        }

        $result.SourceDisk = [string]$sourceCombo.SelectedItem
        $result.DestinationDisk = [string]$destCombo.SelectedItem
        $result.ReportPath = Join-Path $logDir ("CloneDiskPlan-{0:yyyyMMdd-HHmmss}.txt" -f (Get-Date))
        $report = @(
            "Clone Disk Plan - $(Get-Date)",
            '',
            "Source: $($result.SourceDisk)",
            "Destination: $($result.DestinationDisk)",
            '',
            'Important: This report is a plan only. Use Disk Management, Windows Backup/System Image, or trusted imaging software to perform the actual clone.',
            'Do not proceed unless the destination disk is safe to overwrite.',
            '',
            'Detected disks:'
        )
        foreach ($row in @($grid.DataSource)) {
            $report += 'Disk {0}: {1} | {2} | {3} GB | {4} | {5}' -f $row.Number, $row.Name, $row.BusType, $row.SizeGB, $row.PartitionStyle, $row.Status
        }
        Set-Content -LiteralPath $result.ReportPath -Value $report -Encoding UTF8
        [System.Windows.Forms.MessageBox]::Show("Clone plan saved:`r`n$($result.ReportPath)", 'Clone Plan Saved', 'OK', 'Information') | Out-Null
    })
    $closeButton.Add_Click({ $form.Close() })

    & $loadDisks
    [void]$form.ShowDialog()

    return @(
        'Clone Disk Wizard opened.',
        'This toolkit does not silently run destructive sector-copy commands.',
        'Use the wizard to choose source/destination disks, save a clone plan, and launch Disk Management or Windows Backup.',
        $(if ($result.ReportPath) { "Clone plan report: $($result.ReportPath)" } else { 'No clone plan report was saved.' }),
        $(if ($result.OpenedDiskManagement) { 'Disk Management was opened.' } else { 'Disk Management was not opened.' }),
        $(if ($result.OpenedBackup) { 'Windows Backup was opened.' } else { 'Windows Backup was not opened.' })
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
            Details = 'DEPOT-inspired setup step: plug in the old drive or choose a network/source folder, then use the Data Transfer Wizard to copy user files with robocopy logging.'
            SuggestedTool = 'Data Transfer Wizard'
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

    $transferButton = New-Object System.Windows.Forms.Button
    $transferButton.Text = 'Transfer Data'
    $transferButton.Size = New-Object System.Drawing.Size(130, 32)
    [void]$buttonPanel.Controls.Add($transferButton)

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

    $transferButton.Add_Click({
        $transferResult = Show-TtkDataTransferWizard
        if ($transferResult.Completed) {
            for ($i = 0; $i -lt $items.Count; $i++) {
                if ($items[$i].Step -eq 'User data transfer') {
                    $checkList.SetItemChecked($i, $true)
                    $checkList.SelectedIndex = $i
                    break
                }
            }
            $details.Text = "Data transfer complete.`r`n`r`n$($transferResult.Summary)"
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
    'Get-TtkExternalTransferDrives',
    'Show-TtkDataTransferWizard',
    'Get-TtkDiskInventory',
    'Invoke-TtkCloneDiskGuide',
    'Show-TtkCloneDiskWizard',
    'Invoke-TtkNewComputerSetupChecklist',
    'Get-TtkNewComputerSetupChecklistItems',
    'Show-TtkNewComputerSetupChecklist'
)
