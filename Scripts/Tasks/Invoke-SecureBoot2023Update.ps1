<#
.SYNOPSIS
    Lansweeper deployment script for Windows Secure Boot 2023 certificate update remediation.

.DESCRIPTION
    This script attempts the Microsoft Secure Boot 2023 certificate update on eligible
    physical Windows client workstations only.

    It does NOT:
      - Enable Secure Boot
      - Disable Secure Boot
      - Enable UEFI
      - Disable UEFI
      - Convert Legacy/CSM to UEFI
      - Convert MBR to GPT
      - Change BIOS/firmware settings
      - Remediate servers
      - Remediate VMs

    It DOES:
      - Support Test and Live modes
      - Test mode only allows Framework_FW1 / Framework_FW1.albl.com
      - Live mode targets other eligible physical Windows client workstations
      - Blocks servers and VMs
      - Attempts update even on Legacy/CSM systems
      - Writes Microsoft AvailableUpdates trigger
      - Runs Microsoft Secure-Boot-Update task if present
      - Enables the Secure-Boot-Update task if disabled, and leaves it enabled
      - Logs boot mode and Secure Boot state honestly
      - Logs to append-safe Desktop TXT report
      - Backs up/logs BitLocker recovery key if available
      - Suspends BitLocker only before reboot, only if BitLocker is on
      - Reboots only when -Restart is supplied
      - Optionally registers a one-time SYSTEM startup task with -PostRebootCheck to write a post-reboot CSV row

.EXIT CODES
    0 = Success / dry-run success / already updated / queued or attempted without immediate reboot
    1 = Success, reboot was requested and scheduled
    2 = Machine not eligible
    3 = Unexpected failure
    4 = BitLocker failure
    5 = CSV logging failure
    6 = Remediation failure

.NOTES
    Default report directory:
    Desktop\T3CHNRD-SecureBoot2023

    Security note:
    This toolkit copy writes the report to the signed-in user Desktop and may include the BitLocker recovery key when available.
    Protect the Desktop report folder accordingly.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$ScriptVersion = "2026.05.14.1",

    [string]$SchemaVersion = "2026.05.14.1",

    [string]$MappedDriveLetter = "",

    [string]$MappedDriveRoot = "",

    [Parameter(Mandatory = $false)]
    [ValidateSet("Test", "Live")]
    [string]$Mode = "Live",

    [switch]$ToggleMode,

    [switch]$DryRun,

    [switch]$Restart,

    # When used with -Restart, registers a one-time startup task that runs this same script
    # after reboot in post-check-only mode, writes a post-reboot CSV row, then removes itself.
    [switch]$PostRebootCheck,

    # Internal switch used only by the one-time startup task created by -PostRebootCheck.
    # Do not pass this manually unless intentionally running a post-reboot validation-only pass.
    [switch]$PostRebootTaskRun,

    [string]$PostRebootTaskName = "ALBL-SecureBoot2023-PostRebootCheck",

    [int]$PostRebootDelaySeconds = 120,

    [string]$CsvDirectory = (Join-Path ([Environment]::GetFolderPath("Desktop")) "T3CHNRD-SecureBoot2023"),

    [string]$CsvFileName = "SecureBoot2023_Remediation_Results.txt",

    [bool]$WriteLocalLog = $true,

    [string]$LocalLogDirectory = (Join-Path ([Environment]::GetFolderPath("Desktop")) "T3CHNRD-SecureBoot2023"),

    [string]$TestComputerName = "Framework_FW1",

    [string]$TestFqdn = "Framework_FW1.albl.com",

    [string[]]$KnownBadModels = @(
        # Add exact or wildcard model names here.
    ),

    [string[]]$UnsupportedOSVersionPatterns = @(
        # Add OS caption/version/build patterns here if needed.
    ),

    [int]$AvailableUpdatesValue = 0x5944,

    [int]$SecureBootEventLookbackDays = 30
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

# Handle ToggleMode switch - flip between Test and Live.
# Lansweeper should normally run explicit Live mode or default Live mode.
if ($ToggleMode) {
    $Mode = if ($Mode -eq "Test") { "Live" } else { "Test" }
}

# -----------------------------
# Initial state
# -----------------------------

$ScriptStart = Get-Date
$ExitCode = 3

$InvocationCommand = $MyInvocation.Line
$ProductionSafetySummary = "ScriptVersion=$ScriptVersion; SchemaVersion=$SchemaVersion; Mode=$Mode; DryRun=$DryRun; Restart=$Restart; PostRebootCheck=$PostRebootCheck; CsvDirectory=$CsvDirectory"
$ReportingDriveStatus = "Not evaluated"

$ComputerName = $env:COMPUTERNAME
$FullComputerName = $ComputerName

$OSCaption = ""
$OSVersion = ""
$OSBuild = ""
$Manufacturer = ""
$Model = ""
$SerialNumber = ""
$BaseBoardManufacturer = ""
$BaseBoardProduct = ""
$CsProductVendor = ""
$CsProductName = ""
$CsProductUuid = ""

$IsServer = $false
$IsVM = $false
$VMVendor = ""

$BootMode = "Unknown"
$FirmwareType = "Unknown"
$UEFIStatus = "Unknown"
$SecureBootStatus = "Unknown"

$MACAddress = ""
$IPAddress = ""

$BitLockerStatus = "Unknown"
$BitLockerRawProtectionStatus = ""
$BitLockerSuspended = $false
$BitLockerRecoveryKey = ""

$SecureBootTaskExists = $false
$SecureBootTaskState = "Unknown"
$SecureBootTaskEnabledByScript = $false

$AvailableUpdatesBefore = ""
$AvailableUpdatesAfter = ""
$UEFICA2023StatusBefore = ""
$UEFICA2023StatusAfter = ""
$UEFICA2023ErrorBefore = ""
$UEFICA2023ErrorAfter = ""
$UEFICA2023ErrorEventBefore = ""
$UEFICA2023ErrorEventAfter = ""
$SecureBootRegistryRawBefore = ""
$SecureBootRegistryRawAfter = ""

$RecentSecureBootEvents = ""

$PrecheckResult = "Not evaluated"
$RemediationAttempted = $false
$RegistryTriggerWritten = $false
$TaskRunAttempted = $false
$TaskRunResult = "Not attempted"
$CertificateUpdateResult = "Not attempted"
$RebootRequired = $false
$RestartScheduled = $false
$PostRebootTaskRegistered = $false
$PostRebootTaskRegisterResult = ""
$IsPostRebootCheck = [bool]$PostRebootTaskRun
$PostCheckResult = "Not run"
$FinalStatus = "failed"
$ErrorMessage = ""

$FirmwareWarning = ""
$TaskWarning = ""
$DryRunEffective = $false
$DurationSeconds = ""

if ($DryRun -or $WhatIfPreference) {
    $DryRunEffective = $true
}

$CsvColumns = @(
    "Timestamp",
    "ScriptVersion",
    "SchemaVersion",
    "ProductionSafetySummary",
    "InvocationCommand",
    "ReportingDriveStatus",
    "Mode",
    "DryRun",
    "IsPostRebootCheck",
    "PostRebootTaskName",
    "PostRebootTaskRegistered",
    "PostRebootTaskRegisterResult",
    "ComputerName",
    "FullComputerName",
    "MACAddress",
    "IPAddress",
    "OSCaption",
    "OSVersion",
    "OSBuild",
    "Manufacturer",
    "Model",
    "SerialNumber",
    "BaseBoardManufacturer",
    "BaseBoardProduct",
    "CsProductVendor",
    "CsProductName",
    "CsProductUuid",
    "IsServer",
    "IsVM",
    "VMVendor",
    "BootMode",
    "FirmwareType",
    "UEFIStatus",
    "SecureBootStatus",
    "FirmwareWarning",
    "BitLockerStatus",
    "BitLockerRawProtectionStatus",
    "BitLockerSuspended",
    "BitLockerRecoveryKey",
    "SecureBootTaskExists",
    "SecureBootTaskState",
    "SecureBootTaskEnabledByScript",
    "TaskWarning",
    "AvailableUpdatesBefore",
    "AvailableUpdatesAfter",
    "UEFICA2023StatusBefore",
    "UEFICA2023StatusAfter",
    "UEFICA2023ErrorBefore",
    "UEFICA2023ErrorAfter",
    "UEFICA2023ErrorEventBefore",
    "UEFICA2023ErrorEventAfter",
    "SecureBootRegistryRawBefore",
    "SecureBootRegistryRawAfter",
    "RecentSecureBootEvents",
    "PrecheckResult",
    "RemediationAttempted",
    "RegistryTriggerWritten",
    "TaskRunAttempted",
    "TaskRunResult",
    "CertificateUpdateResult",
    "RebootRequired",
    "RestartScheduled",
    "PostCheckResult",
    "FinalStatus",
    "ErrorMessage",
    "DurationSeconds"
)

# -----------------------------
# Utility functions
# -----------------------------

function Write-LocalLog {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    if (-not $WriteLocalLog) {
        return
    }

    try {
        if (-not (Test-Path -LiteralPath $LocalLogDirectory)) {
            New-Item -Path $LocalLogDirectory -ItemType Directory -Force | Out-Null
        }

        $logPath = Join-Path $LocalLogDirectory "SecureBoot2023_Remediation.log"
        $line = "{0} [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"), $Mode, $Message
        Add-Content -LiteralPath $logPath -Value $line -Encoding UTF8
        Write-Output $line
    }
    catch {
        $null = $_
        # Local logging failure should not block CSV logging/remediation.
    }
}

function Test-AdminPrivilege {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)

    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw "Script must run elevated as Administrator."
    }
}

function Get-CsvDriveName {
    param(
        [Parameter(Mandatory = $true)][string]$Path
    )

    if ($Path -match '^([A-Za-z]):\\') {
        return $Matches[1].ToUpperInvariant()
    }

    return ""
}

function Initialize-CsvDrive {
    param(
        [Parameter(Mandatory = $true)][string]$CsvDirectory,
        [Parameter(Mandatory = $true)][string]$MappedDriveLetter,
        [Parameter(Mandatory = $true)][string]$MappedDriveRoot
    )

    $driveName = Get-CsvDriveName -Path $CsvDirectory

    if ([string]::IsNullOrWhiteSpace($driveName)) {
        $script:ReportingDriveStatus = "CSV path is not drive-letter based; no drive mapping required"
        return
    }

    $driveRoot = "$driveName`:"
    if (Test-Path -LiteralPath "$driveRoot\" -ErrorAction SilentlyContinue) {
        $script:ReportingDriveStatus = "Drive $driveRoot is already available"
        return
    }

    $configuredDriveName = ($MappedDriveLetter -replace ':', '').ToUpperInvariant()
    if ($driveName -ne $configuredDriveName) {
        throw "CSV path requires drive $driveRoot, but configured mapped drive is $MappedDriveLetter. Cannot safely map automatically."
    }

    if ([string]::IsNullOrWhiteSpace($MappedDriveRoot)) {
        throw "CSV path requires drive $driveRoot, but MappedDriveRoot is blank."
    }

    Write-LocalLog "Drive $driveRoot is not available. Attempting temporary mapping to $MappedDriveRoot."

    try {
        New-PSDrive -Name $driveName -PSProvider FileSystem -Root $MappedDriveRoot -Scope Global -ErrorAction Stop | Out-Null
        if (Test-Path -LiteralPath "$driveRoot\" -ErrorAction SilentlyContinue) {
            $script:ReportingDriveStatus = "Mapped drive $driveRoot to $MappedDriveRoot using New-PSDrive"
            Write-LocalLog $script:ReportingDriveStatus
            return
        }
    }
    catch {
        Write-LocalLog "New-PSDrive mapping failed for $driveRoot to $MappedDriveRoot. Error: $($_.Exception.Message)"
    }

    try {
        $netUseTarget = "$driveRoot $MappedDriveRoot /persistent:no"
        $process = Start-Process -FilePath "net.exe" -ArgumentList "use $netUseTarget" -NoNewWindow -Wait -PassThru
        if ($process.ExitCode -eq 0 -and (Test-Path -LiteralPath "$driveRoot\" -ErrorAction SilentlyContinue)) {
            $script:ReportingDriveStatus = "Mapped drive $driveRoot to $MappedDriveRoot using net use"
            Write-LocalLog $script:ReportingDriveStatus
            return
        }

        Write-LocalLog "net use mapping returned exit code $($process.ExitCode) for $driveRoot to $MappedDriveRoot."
    }
    catch {
        Write-LocalLog "net use mapping failed for $driveRoot to $MappedDriveRoot. Error: $($_.Exception.Message)"
    }

    if (-not (Test-Path -LiteralPath "$driveRoot\" -ErrorAction SilentlyContinue)) {
        throw "CSV path requires drive $driveRoot, but the drive is not available and automatic mapping to $MappedDriveRoot failed."
    }
}

function Test-CsvSchema {
    param(
        [Parameter(Mandatory = $true)][string]$CsvPath,
        [Parameter(Mandatory = $true)][string]$ExpectedHeaderLine
    )

    if (-not (Test-Path -LiteralPath $CsvPath -PathType Leaf -ErrorAction SilentlyContinue)) {
        return
    }

    $existingHeader = ""
    try {
        $existingHeader = Get-Content -LiteralPath $CsvPath -TotalCount 1 -ErrorAction Stop
    }
    catch {
        throw "Unable to read CSV header for schema check. Path: $CsvPath. Error: $($_.Exception.Message)"
    }

    if ($existingHeader -eq $ExpectedHeaderLine) {
        return
    }

    $directory = Split-Path -Path $CsvPath -Parent
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($CsvPath)
    $extension = [System.IO.Path]::GetExtension($CsvPath)
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $rolloverPath = Join-Path $directory ("{0}_SCHEMA_MISMATCH_{1}{2}" -f $baseName, $timestamp, $extension)

    if (Test-Path -LiteralPath $rolloverPath -ErrorAction SilentlyContinue) {
        $rolloverPath = Join-Path $directory ("{0}_SCHEMA_MISMATCH_{1}_{2}{3}" -f $baseName, $timestamp, ([guid]::NewGuid().ToString("N").Substring(0, 8)), $extension)
    }

    Rename-Item -LiteralPath $CsvPath -NewName (Split-Path -Path $rolloverPath -Leaf) -ErrorAction Stop
    Write-LocalLog "CSV schema/header mismatch detected. Existing CSV rolled over to: $rolloverPath. A fresh CSV will be created at: $CsvPath."
}

function Test-PreflightAccess {
    param(
        [Parameter(Mandatory = $true)][string]$CsvDirectory,
        [Parameter(Mandatory = $true)][string]$LocalLogDirectory,
        [bool]$WriteLocalLog,
        [Parameter(Mandatory = $true)][string]$MappedDriveLetter,
        [Parameter(Mandatory = $true)][string]$MappedDriveRoot
    )

    if ($PSVersionTable.PSVersion.Major -lt 5) {
        throw "This script requires PowerShell 5.0 or later. Current version: $($PSVersionTable.PSVersion)."
    }

    Initialize-CsvDrive -CsvDirectory $CsvDirectory -MappedDriveLetter $MappedDriveLetter -MappedDriveRoot $MappedDriveRoot

    try {
        if (-not (Test-Path -LiteralPath $CsvDirectory -PathType Container -ErrorAction SilentlyContinue)) {
            New-Item -Path $CsvDirectory -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }

        $testFile = Join-Path $CsvDirectory (".write_test_{0}.tmp" -f ([guid]::NewGuid().ToString("N")))
        [System.IO.File]::WriteAllText($testFile, "test")
        Remove-Item -LiteralPath $testFile -Force -ErrorAction Stop
    }
    catch {
        throw "Cannot access or create CSV directory: $CsvDirectory. Error: $($_.Exception.Message)."
    }

    if ($WriteLocalLog) {
        try {
            if (-not (Test-Path -LiteralPath $LocalLogDirectory -PathType Container -ErrorAction SilentlyContinue)) {
                New-Item -Path $LocalLogDirectory -ItemType Directory -Force -ErrorAction Stop | Out-Null
            }

            $testLogFile = Join-Path $LocalLogDirectory (".write_test_{0}.tmp" -f ([guid]::NewGuid().ToString("N")))
            [System.IO.File]::WriteAllText($testLogFile, "test")
            Remove-Item -LiteralPath $testLogFile -Force -ErrorAction Stop
        }
        catch {
            throw "Cannot write to local log directory: $LocalLogDirectory. Error: $($_.Exception.Message)."
        }
    }
}

function ConvertTo-CsvSafeField {
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return '""'
    }

    $stringValue = [string]$Value

    # Prevent formula execution when CSV is opened in Excel.
    if ($stringValue -match '^[=+\-@]') {
        $stringValue = "'" + $stringValue
    }

    $stringValue = $stringValue -replace '"', '""'
    return '"' + $stringValue + '"'
}

function Write-AppendOnlyCsvRow {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Row
    )

    Initialize-CsvDrive -CsvDirectory $CsvDirectory -MappedDriveLetter $MappedDriveLetter -MappedDriveRoot $MappedDriveRoot

    if (-not (Test-Path -LiteralPath $CsvDirectory)) {
        New-Item -Path $CsvDirectory -ItemType Directory -Force | Out-Null
    }

    $csvPath = Join-Path $CsvDirectory $CsvFileName

    $headerLine = ($CsvColumns | ForEach-Object { ConvertTo-CsvSafeField $_ }) -join ","
    Test-CsvSchema -CsvPath $csvPath -ExpectedHeaderLine $headerLine

    $rowLine = ($CsvColumns | ForEach-Object { ConvertTo-CsvSafeField $Row[$_] }) -join ","

    $maxAttempts = 60
    $sleepMs = 500

    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        try {
            $fileStream = [System.IO.FileStream]::new(
                $csvPath,
                [System.IO.FileMode]::OpenOrCreate,
                [System.IO.FileAccess]::ReadWrite,
                [System.IO.FileShare]::None
            )

            try {
                $writeHeader = ($fileStream.Length -eq 0)
                [void]$fileStream.Seek(0, [System.IO.SeekOrigin]::End)

                $streamWriter = [System.IO.StreamWriter]::new(
                    $fileStream,
                    [System.Text.UTF8Encoding]::new($false)
                )

                try {
                    if ($writeHeader) {
                        $streamWriter.WriteLine($headerLine)
                    }

                    $streamWriter.WriteLine($rowLine)
                    $streamWriter.Flush()
                    $fileStream.Flush($true)
                }
                finally {
                    $streamWriter.Dispose()
                }
            }
            finally {
                $fileStream.Dispose()
            }

            return
        }
        catch {
            if ($attempt -eq $maxAttempts) {
                throw "CSV write failed after $maxAttempts attempts. Path: $csvPath. Error: $($_.Exception.Message)"
            }

            Start-Sleep -Milliseconds $sleepMs
        }
    }
}

function Get-FullComputerNameSafe {
    try {
        $props = [System.Net.NetworkInformation.IPGlobalProperties]::GetIPGlobalProperties()
        if (-not [string]::IsNullOrWhiteSpace($props.DomainName)) {
            return "$env:COMPUTERNAME.$($props.DomainName)"
        }
    }
    catch {
        Write-LocalLog "Failed to determine DNS full computer name: $($_.Exception.Message)"
    }

    return $env:COMPUTERNAME
}

function Get-NetworkIdentity {
    $result = [ordered]@{
        IPAddress  = ""
        MACAddress = ""
    }

    try {
        $configs = Get-CimInstance Win32_NetworkAdapterConfiguration |
            Where-Object {
                $_.IPEnabled -eq $true -and
                $_.IPAddress -and
                $_.MACAddress
            }

        $ips = New-Object System.Collections.Generic.List[string]
        $macs = New-Object System.Collections.Generic.List[string]

        foreach ($config in $configs) {
            foreach ($ip in $config.IPAddress) {
                if ($ip -match '^\d{1,3}(\.\d{1,3}){3}$' -and $ip -notmatch '^169\.254\.') {
                    [void]$ips.Add($ip)
                }
            }

            if ($config.MACAddress) {
                [void]$macs.Add($config.MACAddress)
            }
        }

        $result.IPAddress = (($ips | Select-Object -Unique) -join ";")
        $result.MACAddress = (($macs | Select-Object -Unique) -join ";")
    }
    catch {
        Write-LocalLog "Failed to collect network identity: $($_.Exception.Message)"
    }

    return $result
}

function Test-WildcardMatchList {
    param(
        [AllowNull()]
        [string]$Value,

        [string[]]$Patterns
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $false
    }

    foreach ($pattern in $Patterns) {
        if (-not [string]::IsNullOrWhiteSpace($pattern)) {
            if ($Value -like $pattern) {
                return $true
            }
        }
    }

    return $false
}

function Get-VmDetection {
    param(
        [AllowNull()][string]$Manufacturer,
        [AllowNull()][string]$Model,
        [AllowNull()][string]$SerialNumber,
        [AllowNull()][string]$BiosVersion,
        [AllowNull()][string]$BaseBoardManufacturer,
        [AllowNull()][string]$BaseBoardProduct,
        [AllowNull()][string]$CsProductVendor,
        [AllowNull()][string]$CsProductName,
        [AllowNull()][string]$CsProductUuid
    )

    $haystack = "$Manufacturer $Model $SerialNumber $BiosVersion $BaseBoardManufacturer $BaseBoardProduct $CsProductVendor $CsProductName $CsProductUuid"

    $vmSignatures = @(
        @{ Pattern = "VMware"; Vendor = "VMware" },
        @{ Pattern = "VirtualBox"; Vendor = "Oracle VirtualBox" },
        @{ Pattern = "Oracle Corporation"; Vendor = "Oracle VirtualBox" },
        @{ Pattern = "KVM"; Vendor = "KVM" },
        @{ Pattern = "QEMU"; Vendor = "QEMU" },
        @{ Pattern = "Xen"; Vendor = "Xen" },
        @{ Pattern = "Hyper-V"; Vendor = "Microsoft Hyper-V" },
        @{ Pattern = "Microsoft Corporation Virtual Machine"; Vendor = "Microsoft Virtual Machine" },
        @{ Pattern = "Virtual Machine"; Vendor = "Microsoft Virtual Machine" },
        @{ Pattern = "Parallels"; Vendor = "Parallels" },
        @{ Pattern = "Bochs"; Vendor = "Bochs" },
        @{ Pattern = "OpenStack"; Vendor = "OpenStack" },
        @{ Pattern = "Amazon EC2"; Vendor = "Amazon EC2" },
        @{ Pattern = "Google Compute Engine"; Vendor = "Google Compute Engine" },
        @{ Pattern = "Azure"; Vendor = "Microsoft Azure" },
        @{ Pattern = "HVM domU"; Vendor = "Xen/AWS" }
    )

    foreach ($signature in $vmSignatures) {
        if ($haystack -match [regex]::Escape($signature.Pattern)) {
            return [ordered]@{
                IsVM     = $true
                VMVendor = $signature.Vendor
            }
        }
    }

    return [ordered]@{
        IsVM     = $false
        VMVendor = ""
    }
}

function Get-FirmwareSecureBootState {
    $result = [ordered]@{
        BootMode         = "Unknown"
        FirmwareType     = "Unknown"
        UEFIStatus       = "Unknown"
        SecureBootStatus = "Unknown"
    }

    try {
        $computerInfo = Get-ComputerInfo -Property BiosFirmwareType -ErrorAction Stop

        if ($computerInfo.BiosFirmwareType) {
            $result.FirmwareType = [string]$computerInfo.BiosFirmwareType
        }
    }
    catch {
        Write-LocalLog "Get-ComputerInfo BiosFirmwareType failed: $($_.Exception.Message)"
    }

    if ($result.FirmwareType -match "Uefi") {
        $result.BootMode = "UEFI"
        $result.UEFIStatus = "UEFI"
    }
    elseif ($result.FirmwareType -match "Bios|Legacy") {
        $result.BootMode = "Legacy"
        $result.UEFIStatus = "Legacy"
        $result.SecureBootStatus = "Legacy"
        return $result
    }

    try {
        $secureBoot = Confirm-SecureBootUEFI -ErrorAction Stop

        if ($secureBoot -eq $true) {
            $result.SecureBootStatus = "On"
            if ($result.BootMode -eq "Unknown") {
                $result.BootMode = "UEFI"
                $result.UEFIStatus = "UEFI"
            }
        }
        elseif ($secureBoot -eq $false) {
            $result.SecureBootStatus = "Off"
            if ($result.BootMode -eq "Unknown") {
                $result.BootMode = "UEFI"
                $result.UEFIStatus = "UEFI"
            }
        }
    }
    catch {
        $message = $_.Exception.Message

        if ($result.BootMode -eq "Legacy") {
            $result.SecureBootStatus = "Legacy"
        }
        elseif ($message -match "Cmdlet not supported|not supported|unsupported|Secure Boot") {
            $result.SecureBootStatus = "Unsupported"
        }
        else {
            $result.SecureBootStatus = "Unknown: $message"
        }
    }

    return $result
}

function Get-SecureBootRegistryState {
    $path = "HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot"

    $state = [ordered]@{
        AvailableUpdates     = ""
        UEFICA2023Status     = ""
        UEFICA2023Error      = ""
        UEFICA2023ErrorEvent = ""
        Raw                  = ""
    }

    try {
        if (Test-Path -LiteralPath $path) {
            $properties = Get-ItemProperty -LiteralPath $path -ErrorAction Stop

            foreach ($name in @("AvailableUpdates", "UEFICA2023Status", "UEFICA2023Error", "UEFICA2023ErrorEvent")) {
                if ($properties.PSObject.Properties.Name -contains $name) {
                    $state[$name] = [string]$properties.$name
                }
            }

            $rawPairs = New-Object System.Collections.Generic.List[string]
            foreach ($property in $properties.PSObject.Properties) {
                if ($property.Name -notmatch '^PS') {
                    [void]$rawPairs.Add(("{0}={1}" -f $property.Name, $property.Value))
                }
            }

            $state.Raw = ($rawPairs -join "; ")
        }
    }
    catch {
        Write-LocalLog "Failed to read Secure Boot registry state: $($_.Exception.Message)"
    }

    return $state
}

function Test-SecureBoot2023Updated {
    param(
        [AllowNull()][string]$Status
    )

    if ($Status -match "Updated|Complete|Completed|Success") {
        return $true
    }

    # Some systems may expose numeric status values. Keep this conservative:
    # do not claim success unless the status is obviously successful.
    return $false
}

function Get-SecureBootTaskState {
    $result = [ordered]@{
        Exists = $false
        State  = "Missing"
    }

    try {
        $task = Get-ScheduledTask -TaskPath "\Microsoft\Windows\PI\" -TaskName "Secure-Boot-Update" -ErrorAction Stop
        $result.Exists = $true
        $result.State = [string]$task.State
    }
    catch {
        $result.Exists = $false
        $result.State = "Missing"
    }

    return $result
}

function Get-RecentSecureBootEvent {
    param(
        [int]$LookbackDays
    )

    try {
        $startTime = (Get-Date).AddDays(-1 * [Math]::Abs($LookbackDays))

        $eventIds = @(
            1032,
            1036,
            1043,
            1044,
            1045,
            1795,
            1796,
            1799,
            1801,
            1802,
            1803,
            1808,
            1034
        )

        $events = Get-WinEvent -FilterHashtable @{
            LogName   = "System"
            Id        = $eventIds
            StartTime = $startTime
        } -MaxEvents 25 -ErrorAction SilentlyContinue

        if (-not $events) {
            return ""
        }

        return (($events | Sort-Object TimeCreated -Descending | ForEach-Object {
            "{0}|ID={1}|Provider={2}" -f $_.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss"), $_.Id, $_.ProviderName
        }) -join " ; ")
    }
    catch {
        Write-LocalLog "Failed to read recent Secure Boot events: $($_.Exception.Message)"
        return ""
    }
}

function ConvertTo-BitLockerStatus {
    param(
        [AllowNull()][object]$Value
    )

    if ($null -eq $Value) {
        return "Unknown"
    }

    $text = [string]$Value

    if ($text -match '^\s*1\s*$|On|ProtectionOn|True') {
        return "On"
    }

    if ($text -match '^\s*0\s*$|Off|ProtectionOff|False') {
        return "Off"
    }

    return "Unknown"
}

function Get-OsDriveBitLockerInfo {
    $result = [ordered]@{
        Status              = "Unknown"
        RawProtectionStatus = ""
        RecoveryKey         = ""
    }

    try {
        $volume = Get-BitLockerVolume -MountPoint $env:SystemDrive -ErrorAction Stop

        $result.RawProtectionStatus = [string]$volume.ProtectionStatus
        $result.Status = ConvertTo-BitLockerStatus -Value $volume.ProtectionStatus

        $recoveryProtector = $volume.KeyProtector |
            Where-Object {
                $_.KeyProtectorType -eq "RecoveryPassword" -and
                -not [string]::IsNullOrWhiteSpace($_.RecoveryPassword)
            } |
            Select-Object -First 1

        if ($recoveryProtector) {
            $result.RecoveryKey = [string]$recoveryProtector.RecoveryPassword
        }
    }
    catch {
        Write-LocalLog "Get-BitLockerVolume failed: $($_.Exception.Message)"
    }

    try {
        $statusOutput = & manage-bde.exe -status $env:SystemDrive 2>&1
        $statusText = $statusOutput | Out-String

        if ($result.Status -eq "Unknown") {
            if ($statusText -match "Protection Status:\s+Protection On") {
                $result.Status = "On"
            }
            elseif ($statusText -match "Protection Status:\s+Protection Off") {
                $result.Status = "Off"
            }
        }

        if ([string]::IsNullOrWhiteSpace($result.RawProtectionStatus)) {
            if ($statusText -match "Protection Status:\s+([^\r\n]+)") {
                $result.RawProtectionStatus = $Matches[1].Trim()
            }
        }
    }
    catch {
        Write-LocalLog "manage-bde -status failed: $($_.Exception.Message)"
    }

    try {
        if ([string]::IsNullOrWhiteSpace($result.RecoveryKey)) {
            $protectorOutput = & manage-bde.exe -protectors -get $env:SystemDrive 2>&1
            $protectorText = $protectorOutput | Out-String

            if ($protectorText -match "([0-9]{6}-[0-9]{6}-[0-9]{6}-[0-9]{6}-[0-9]{6}-[0-9]{6}-[0-9]{6}-[0-9]{6})") {
                $result.RecoveryKey = $Matches[1]
            }
        }
    }
    catch {
        Write-LocalLog "manage-bde -protectors -get failed: $($_.Exception.Message)"
    }

    return $result
}

function Suspend-BitLockerForOneRebootIfNeeded {
    param(
        [AllowNull()][string]$CurrentStatus
    )

    if ($CurrentStatus -ne "On") {
        Write-LocalLog "BitLocker status is '$CurrentStatus'. No BitLocker suspension needed."
        return $false
    }

    if ($DryRunEffective) {
        Write-LocalLog "DryRun: would suspend BitLocker for one reboot."
        return $false
    }

    try {
        Suspend-BitLocker -MountPoint $env:SystemDrive -RebootCount 1 -ErrorAction Stop | Out-Null
        Write-LocalLog "BitLocker suspended for one reboot using Suspend-BitLocker."
        return $true
    }
    catch {
        Write-LocalLog "Suspend-BitLocker failed, trying manage-bde: $($_.Exception.Message)"
    }

    try {
        & manage-bde.exe -protectors -disable $env:SystemDrive -RebootCount 1 | Out-Null
        Write-LocalLog "BitLocker suspended for one reboot using manage-bde."
        return $true
    }
    catch {
        throw "Failed to suspend BitLocker on $env:SystemDrive. Error: $($_.Exception.Message)"
    }
}


function ConvertTo-PowerShellLiteral {
    param(
        [AllowNull()][string]$Value
    )

    if ($null -eq $Value) {
        return "''"
    }

    return "'" + ($Value -replace "'", "''") + "'"
}


function ConvertTo-WindowsArgument {
    param(
        [AllowNull()][string]$Value
    )

    if ($null -eq $Value) {
        return '""'
    }

    return '"' + ($Value -replace '"', '\"') + '"'
}

function Invoke-PostRebootCheckTaskRegistration {
    param(
        [Parameter(Mandatory = $true)][string]$ScriptPath,
        [Parameter(Mandatory = $true)][string]$TaskName,
        [Parameter(Mandatory = $true)][string]$Mode,
        [Parameter(Mandatory = $true)][string]$CsvDirectory,
        [Parameter(Mandatory = $true)][string]$CsvFileName,
        [Parameter(Mandatory = $true)][string]$LocalLogDirectory,
        [int]$DelaySeconds = 120
    )

    if (-not (Test-Path -LiteralPath $ScriptPath)) {
        throw "Cannot register post-reboot check task because script path was not found: $ScriptPath"
    }

    $helperDirectory = Join-Path $env:ProgramData "ALBL\SecureBoot2023"
    if (-not (Test-Path -LiteralPath $helperDirectory)) {
        New-Item -Path $helperDirectory -ItemType Directory -Force | Out-Null
    }

    $helperScript = Join-Path $helperDirectory "Invoke-SecureBoot2023-PostRebootCheck.ps1"

    $helperContent = @"
Start-Sleep -Seconds $DelaySeconds
& $(ConvertTo-PowerShellLiteral $ScriptPath) -Mode $(ConvertTo-PowerShellLiteral $Mode) -PostRebootCheck -PostRebootTaskRun -CsvDirectory $(ConvertTo-PowerShellLiteral $CsvDirectory) -CsvFileName $(ConvertTo-PowerShellLiteral $CsvFileName) -LocalLogDirectory $(ConvertTo-PowerShellLiteral $LocalLogDirectory) -PostRebootTaskName $(ConvertTo-PowerShellLiteral $TaskName) -MappedDriveLetter $(ConvertTo-PowerShellLiteral $MappedDriveLetter) -MappedDriveRoot $(ConvertTo-PowerShellLiteral $MappedDriveRoot) -ScriptVersion $(ConvertTo-PowerShellLiteral $ScriptVersion) -SchemaVersion $(ConvertTo-PowerShellLiteral $SchemaVersion)
"@

    Set-Content -LiteralPath $helperScript -Value $helperContent -Encoding UTF8 -Force

    try {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    }
    catch {
        Write-LocalLog "Existing post-reboot task cleanup before registration failed or was not needed: $($_.Exception.Message)"
    }

    $action = New-ScheduledTaskAction `
        -Execute "powershell.exe" `
        -Argument ("-NoProfile -ExecutionPolicy Bypass -File " + (ConvertTo-WindowsArgument $helperScript))

    $trigger = New-ScheduledTaskTrigger -AtStartup
    $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -RunLevel Highest
    $settings = New-ScheduledTaskSettingsSet `
        -StartWhenAvailable `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries `
        -MultipleInstances IgnoreNew `
        -ExecutionTimeLimit (New-TimeSpan -Minutes 20)

    Register-ScheduledTask `
        -TaskName $TaskName `
        -Action $action `
        -Trigger $trigger `
        -Principal $principal `
        -Settings $settings `
        -Force | Out-Null

    return "Registered startup task '$TaskName' as SYSTEM. HelperScript=$helperScript DelaySeconds=$DelaySeconds"
}

function Invoke-PostRebootCheckTaskCleanup {
    param(
        [Parameter(Mandatory = $true)][string]$TaskName
    )

    try {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
        Write-LocalLog "Post-reboot task cleanup attempted for '$TaskName'."
    }
    catch {
        Write-LocalLog "Post-reboot task cleanup failed for '$TaskName': $($_.Exception.Message)"
    }

    try {
        $helperScript = Join-Path $env:ProgramData "ALBL\SecureBoot2023\Invoke-SecureBoot2023-PostRebootCheck.ps1"
        if (Test-Path -LiteralPath $helperScript) {
            Remove-Item -LiteralPath $helperScript -Force -ErrorAction SilentlyContinue
        }
    }
    catch {
        Write-LocalLog "Post-reboot helper cleanup failed: $($_.Exception.Message)"
    }
}

function Get-PostRebootCertificateStatus {
    param(
        [AllowNull()][string]$Status,
        [AllowNull()][string]$AvailableUpdates,
        [AllowNull()][string]$RecentEvents
    )

    if (Test-SecureBoot2023Updated -Status $Status) {
        return [ordered]@{
            FinalStatus = "success"
            Result      = "Post-reboot check: UEFICA2023Status indicates updated"
        }
    }

    if ($RecentEvents -match "ID=1808") {
        return [ordered]@{
            FinalStatus = "success"
            Result      = "Post-reboot check: event log includes TPM-WMI 1808, indicating Secure Boot certificate application success"
        }
    }

    if ($RecentEvents -match "ID=1801|ID=1803") {
        return [ordered]@{
            FinalStatus = "post-check-not-fully-applied"
            Result      = "Post-reboot check: TPM-WMI 1801/1803 still present; certificates appear staged/updated in Windows but not confirmed applied to firmware"
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($AvailableUpdates) -and $AvailableUpdates -ne "0") {
        return [ordered]@{
            FinalStatus = "post-check-pending"
            Result      = "Post-reboot check: AvailableUpdates is still non-zero ($AvailableUpdates); update is not confirmed complete"
        }
    }

    return [ordered]@{
        FinalStatus = "post-check-not-confirmed"
        Result      = "Post-reboot check completed, but no success status/event was detected"
    }
}

function Invoke-SecureBoot2023UpdateAttempt {
    param(
        [int]$Value,
        [bool]$TaskExists
    )

    $result = [ordered]@{
        RegistryTriggerWritten = $false
        TaskRunAttempted       = $false
        TaskRunResult          = "Not attempted"
        CertificateResult      = "Not attempted"
    }

    $registryPath = "HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot"

    if ($DryRunEffective) {
        $result.CertificateResult = "DryRun: would write AvailableUpdates trigger and run or enable task if available"
        Write-LocalLog $result.CertificateResult
        return $result
    }

    try {
        if (-not (Test-Path -LiteralPath $registryPath)) {
            New-Item -Path $registryPath -Force | Out-Null
        }

        New-ItemProperty `
            -Path $registryPath `
            -Name "AvailableUpdates" `
            -Value $Value `
            -PropertyType DWord `
            -Force | Out-Null

        $result.RegistryTriggerWritten = $true
        Write-LocalLog ("AvailableUpdates registry trigger written: 0x{0:X}" -f $Value)
    }
    catch {
        throw "Failed to write AvailableUpdates registry trigger. Error: $($_.Exception.Message)"
    }

    if (-not $TaskExists) {
        $result.TaskRunResult = "Task missing"
        $result.CertificateResult = "Registry trigger written; Secure-Boot-Update task missing; queued only"
        Write-LocalLog $result.CertificateResult
        return $result
    }

    try {
        $freshTask = Get-ScheduledTask -TaskPath "\Microsoft\Windows\PI\" -TaskName "Secure-Boot-Update" -ErrorAction Stop
        $currentTaskState = [string]$freshTask.State

        if ($currentTaskState -eq "Disabled") {
            Enable-ScheduledTask -TaskPath "\Microsoft\Windows\PI\" -TaskName "Secure-Boot-Update" -ErrorAction Stop | Out-Null
            $script:SecureBootTaskEnabledByScript = $true
            $currentTaskState = "Ready"
            Write-LocalLog "Secure-Boot-Update task was disabled. Enabled and left enabled."
        }

        $result.TaskRunAttempted = $true

        $process = Start-Process `
            -FilePath "schtasks.exe" `
            -ArgumentList '/Run /TN "\Microsoft\Windows\PI\Secure-Boot-Update"' `
            -NoNewWindow `
            -Wait `
            -PassThru

        if ($process.ExitCode -eq 0) {
            $result.TaskRunResult = "Task start accepted"
            $result.CertificateResult = "Registry trigger written; Secure-Boot-Update task start accepted; not confirmed updated"
        }
        else {
            $result.TaskRunResult = "schtasks exit code $($process.ExitCode)"
            $result.CertificateResult = "Registry trigger written; task run returned exit code $($process.ExitCode); not confirmed updated"
        }

        Write-LocalLog $result.CertificateResult
        return $result
    }
    catch {
        $result.TaskRunAttempted = $true
        $result.TaskRunResult = "Task run failed: $($_.Exception.Message)"
        $result.CertificateResult = "Registry trigger written; task run failed; not confirmed updated"
        Write-LocalLog $result.TaskRunResult
        return $result
    }
}

function Get-EligibilityResult {
    param(
        [string]$Mode,
        [string]$ComputerName,
        [string]$FullComputerName,
        [string]$TestComputerName,
        [string]$TestFqdn,
        [string]$OSCaption,
        [string]$OSVersion,
        [string]$OSBuild,
        [bool]$IsServer,
        [bool]$IsVM,
        [string]$VMVendor,
        [string]$Model,
        [int]$PCSystemType,
        [string[]]$KnownBadModels,
        [string[]]$UnsupportedOSVersionPatterns
    )

    $reasons = New-Object System.Collections.Generic.List[string]

    $isTestTarget = (
        $ComputerName -ieq $TestComputerName -or
        $FullComputerName -ieq $TestFqdn
    )

    if ($Mode -eq "Test" -and -not $isTestTarget) {
        [void]$reasons.Add("Test mode only allows $TestComputerName / $TestFqdn")
    }

    if ($Mode -eq "Live" -and $isTestTarget) {
        [void]$reasons.Add("Live mode excludes the designated test machine")
    }

    if ($IsServer) {
        [void]$reasons.Add("Server OS detected")
    }

    if ($IsVM) {
        [void]$reasons.Add("Virtual machine detected: $VMVendor")
    }

    if ($OSCaption -notmatch "Windows 10|Windows 11") {
        [void]$reasons.Add("Unsupported Windows client OS: $OSCaption")
    }

    if (Test-WildcardMatchList -Value $OSCaption -Patterns $UnsupportedOSVersionPatterns) {
        [void]$reasons.Add("OS caption excluded by unsupported OS list: $OSCaption")
    }

    if (Test-WildcardMatchList -Value $OSVersion -Patterns $UnsupportedOSVersionPatterns) {
        [void]$reasons.Add("OS version excluded by unsupported OS list: $OSVersion")
    }

    if (Test-WildcardMatchList -Value $OSBuild -Patterns $UnsupportedOSVersionPatterns) {
        [void]$reasons.Add("OS build excluded by unsupported OS list: $OSBuild")
    }

    if (Test-WildcardMatchList -Value $Model -Patterns $KnownBadModels) {
        [void]$reasons.Add("Known-bad workstation model excluded: $Model")
    }

    # Only allow workstation/desktop/laptop class systems.
    # 1 = Desktop, 2 = Mobile, 3 = Workstation.
    # Some physical machines report 0 = Unspecified, so do not block solely on 0.
    if ($PCSystemType -notin @(0, 1, 2, 3)) {
        [void]$reasons.Add("Computer system type is not desktop/laptop/workstation. PCSystemType=$PCSystemType")
    }

    if ($reasons.Count -gt 0) {
        return [ordered]@{
            Eligible = $false
            Reason   = "fail: " + (($reasons | Select-Object -Unique) -join "; ")
        }
    }

    return [ordered]@{
        Eligible = $true
        Reason   = "pass"
    }
}

# -----------------------------
# Main
# -----------------------------

try {
    Write-LocalLog "Script started. Mode=$Mode DryRun=$DryRunEffective Restart=$Restart PostRebootCheck=$PostRebootCheck PostRebootTaskRun=$PostRebootTaskRun"
    Write-LocalLog "PowerShell version: $($PSVersionTable.PSVersion)"
    Write-LocalLog "CSV directory: $CsvDirectory"
    Write-LocalLog "CSV file: $CsvFileName"
    Write-LocalLog "Local log directory: $LocalLogDirectory"
    Write-LocalLog "Effective mode: $Mode"
    Write-LocalLog "Invocation command: $InvocationCommand"
    Write-LocalLog "ScriptVersion=$ScriptVersion SchemaVersion=$SchemaVersion"
    if ($ToggleMode) {
        Write-LocalLog "ToggleMode switch was applied. Effective Mode=$Mode"
    }
    if ($DryRunEffective) {
        Write-LocalLog "DryRun/WhatIf mode ACTIVE - no changes will be made to system state"
    }

    Test-AdminPrivilege
    Write-LocalLog "Running as admin: True"

    Test-PreflightAccess -CsvDirectory $CsvDirectory -LocalLogDirectory $LocalLogDirectory -WriteLocalLog $WriteLocalLog -MappedDriveLetter $MappedDriveLetter -MappedDriveRoot $MappedDriveRoot

    $FullComputerName = Get-FullComputerNameSafe

    $networkIdentity = Get-NetworkIdentity
    $IPAddress = $networkIdentity.IPAddress
    $MACAddress = $networkIdentity.MACAddress

    $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
    $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
    $bios = Get-CimInstance Win32_BIOS -ErrorAction Stop

    try {
        $baseBoard = Get-CimInstance Win32_BaseBoard -ErrorAction Stop
        $BaseBoardManufacturer = [string]$baseBoard.Manufacturer
        $BaseBoardProduct = [string]$baseBoard.Product
    }
    catch {
        Write-LocalLog "BaseBoard CIM collection failed: $($_.Exception.Message)"
    }

    try {
        $csProduct = Get-CimInstance Win32_ComputerSystemProduct -ErrorAction Stop
        $CsProductVendor = [string]$csProduct.Vendor
        $CsProductName = [string]$csProduct.Name
        $CsProductUuid = [string]$csProduct.UUID
    }
    catch {
        Write-LocalLog "ComputerSystemProduct CIM collection failed: $($_.Exception.Message)"
    }

    $OSCaption = [string]$os.Caption
    $OSVersion = [string]$os.Version
    $OSBuild = [string]$os.BuildNumber
    $Manufacturer = [string]$cs.Manufacturer
    $Model = [string]$cs.Model
    $SerialNumber = [string]$bios.SerialNumber

    $IsServer = (($os.ProductType -ne 1) -or ($OSCaption -match "Server"))

    $vmResult = Get-VmDetection `
        -Manufacturer $Manufacturer `
        -Model $Model `
        -SerialNumber $SerialNumber `
        -BiosVersion (($bios.SMBIOSBIOSVersion | Out-String).Trim()) `
        -BaseBoardManufacturer $BaseBoardManufacturer `
        -BaseBoardProduct $BaseBoardProduct `
        -CsProductVendor $CsProductVendor `
        -CsProductName $CsProductName `
        -CsProductUuid $CsProductUuid

    $IsVM = [bool]$vmResult.IsVM
    $VMVendor = [string]$vmResult.VMVendor

    $firmwareState = Get-FirmwareSecureBootState
    $BootMode = $firmwareState.BootMode
    $FirmwareType = $firmwareState.FirmwareType
    $UEFIStatus = $firmwareState.UEFIStatus
    $SecureBootStatus = $firmwareState.SecureBootStatus

    if ($BootMode -eq "Legacy" -or $SecureBootStatus -eq "Legacy") {
        $FirmwareWarning = "Legacy/CSM detected. Update will still be attempted. Script will not change firmware mode. Actual certificate installation may depend on Windows/firmware support."
    }
    elseif ($SecureBootStatus -eq "Off") {
        $FirmwareWarning = "Secure Boot is off. Update will still be attempted. Script will not enable Secure Boot."
    }
    elseif ($SecureBootStatus -eq "Unsupported") {
        $FirmwareWarning = "Secure Boot unsupported or not exposed. Update will be attempted only through available Windows servicing mechanism."
    }

    $taskState = Get-SecureBootTaskState
    $SecureBootTaskExists = [bool]$taskState.Exists
    $SecureBootTaskState = [string]$taskState.State

    if (-not $SecureBootTaskExists) {
        $TaskWarning = "Secure-Boot-Update task is missing. Registry trigger will still be written if remediation is attempted, but servicing cannot progress until the task exists."
    }
    elseif ($SecureBootTaskState -eq "Disabled") {
        $TaskWarning = "Secure-Boot-Update task is disabled. Script is approved to enable it and leave it enabled if remediation is attempted."
    }

    $registryBefore = Get-SecureBootRegistryState
    $AvailableUpdatesBefore = $registryBefore.AvailableUpdates
    $UEFICA2023StatusBefore = $registryBefore.UEFICA2023Status
    $UEFICA2023ErrorBefore = $registryBefore.UEFICA2023Error
    $UEFICA2023ErrorEventBefore = $registryBefore.UEFICA2023ErrorEvent
    $SecureBootRegistryRawBefore = $registryBefore.Raw

    $RecentSecureBootEvents = Get-RecentSecureBootEvent -LookbackDays $SecureBootEventLookbackDays

    $bitLockerInfo = Get-OsDriveBitLockerInfo
    $BitLockerStatus = $bitLockerInfo.Status
    $BitLockerRawProtectionStatus = $bitLockerInfo.RawProtectionStatus
    $BitLockerRecoveryKey = $bitLockerInfo.RecoveryKey

    $eligibility = Get-EligibilityResult `
        -Mode $Mode `
        -ComputerName $ComputerName `
        -FullComputerName $FullComputerName `
        -TestComputerName $TestComputerName `
        -TestFqdn $TestFqdn `
        -OSCaption $OSCaption `
        -OSVersion $OSVersion `
        -OSBuild $OSBuild `
        -IsServer $IsServer `
        -IsVM $IsVM `
        -VMVendor $VMVendor `
        -Model $Model `
        -PCSystemType ([int]$cs.PCSystemType) `
        -KnownBadModels $KnownBadModels `
        -UnsupportedOSVersionPatterns $UnsupportedOSVersionPatterns

    $PrecheckResult = $eligibility.Reason

    if ($PostRebootTaskRun) {
        # Post-reboot validation-only mode. Do not write registry triggers, run update tasks, suspend BitLocker, or reboot.
        $RemediationAttempted = $false
        $RegistryTriggerWritten = $false
        $TaskRunAttempted = $false
        $TaskRunResult = "PostRebootCheck: not attempted"
        $AvailableUpdatesAfter = $AvailableUpdatesBefore
        $UEFICA2023StatusAfter = $UEFICA2023StatusBefore
        $UEFICA2023ErrorAfter = $UEFICA2023ErrorBefore
        $UEFICA2023ErrorEventAfter = $UEFICA2023ErrorEventBefore
        $SecureBootRegistryRawAfter = $SecureBootRegistryRawBefore

        if (-not $eligibility.Eligible) {
            $CertificateUpdateResult = "PostRebootCheck: not eligible: $($eligibility.Reason)"
            $PostCheckResult = "Post-reboot check ran, but machine is not eligible"
            $FinalStatus = "not-eligible"
            $ExitCode = 2
        }
        else {
            $postResult = Get-PostRebootCertificateStatus `
                -Status $UEFICA2023StatusBefore `
                -AvailableUpdates $AvailableUpdatesBefore `
                -RecentEvents $RecentSecureBootEvents

            $CertificateUpdateResult = $postResult.Result
            $PostCheckResult = $postResult.Result
            $FinalStatus = $postResult.FinalStatus
            $ExitCode = 0
        }

        Invoke-PostRebootCheckTaskCleanup -TaskName $PostRebootTaskName
    }
    elseif (-not $eligibility.Eligible) {
        $FinalStatus = "not-eligible"
        $CertificateUpdateResult = "Not attempted: not eligible"
        $PostCheckResult = "Not run"
        $ExitCode = 2
        Write-LocalLog "Machine not eligible: $PrecheckResult"
    }
    else {
        $alreadyUpdatedBefore = Test-SecureBoot2023Updated -Status $UEFICA2023StatusBefore

        if ($alreadyUpdatedBefore) {
            $FinalStatus = "success"
            $CertificateUpdateResult = "Already updated before remediation"
            $PostCheckResult = "UEFICA2023Status indicates updated"
            $ExitCode = 0
            Write-LocalLog "Machine already appears updated."
        }
        elseif ($DryRunEffective) {
            $RemediationAttempted = $false
            $RegistryTriggerWritten = $false
            $TaskRunAttempted = $false
            $TaskRunResult = if ($WhatIfPreference) { "WhatIf: not attempted" } else { "DryRun: not attempted" }
            $CertificateUpdateResult = if ($WhatIfPreference) { "WhatIf: eligible; no changes made" } else { "DryRun: eligible; no changes made" }
            $PostCheckResult = if ($WhatIfPreference) { "WhatIf only" } else { "DryRun only" }
            $FinalStatus = "success"
            $ExitCode = 0
            Write-LocalLog "$CertificateUpdateResult"
        }
        else {
            if ($PSCmdlet.ShouldProcess($ComputerName, "Attempt Microsoft Secure Boot 2023 certificate update")) {
                $RemediationAttempted = $true

                $attemptResult = Invoke-SecureBoot2023UpdateAttempt `
                    -Value $AvailableUpdatesValue `
                    -TaskExists $SecureBootTaskExists

                $RegistryTriggerWritten = [bool]$attemptResult.RegistryTriggerWritten
                $TaskRunAttempted = [bool]$attemptResult.TaskRunAttempted
                $TaskRunResult = [string]$attemptResult.TaskRunResult
                $CertificateUpdateResult = [string]$attemptResult.CertificateResult

                # Refresh task state after possible enable.
                $taskStateAfter = Get-SecureBootTaskState
                $SecureBootTaskExists = [bool]$taskStateAfter.Exists
                $SecureBootTaskState = [string]$taskStateAfter.State

                Start-Sleep -Seconds 5

                $registryAfter = Get-SecureBootRegistryState
                $AvailableUpdatesAfter = $registryAfter.AvailableUpdates
                $UEFICA2023StatusAfter = $registryAfter.UEFICA2023Status
                $UEFICA2023ErrorAfter = $registryAfter.UEFICA2023Error
                $UEFICA2023ErrorEventAfter = $registryAfter.UEFICA2023ErrorEvent
                $SecureBootRegistryRawAfter = $registryAfter.Raw

                $updatedAfter = Test-SecureBoot2023Updated -Status $UEFICA2023StatusAfter

                if ($updatedAfter) {
                    $RebootRequired = $false
                    $PostCheckResult = "UEFICA2023Status indicates updated"
                    $FinalStatus = "success"
                    $ExitCode = 0
                }
                elseif ($RegistryTriggerWritten) {
                    # Do not assume that writing the trigger means a reboot is required.
                    # Microsoft's process can remain queued/in-progress and retry through the scheduled task.
                    if ($BootMode -eq "Legacy" -or $SecureBootStatus -eq "Legacy") {
                        $PostCheckResult = "Attempted/queued with legacy warning. Not confirmed updated."
                        $FinalStatus = "queued-with-legacy-warning"
                    }
                    elseif (-not $SecureBootTaskExists) {
                        $PostCheckResult = "Registry trigger written, but Secure-Boot-Update task missing. Not confirmed updated."
                        $FinalStatus = "task-missing"
                    }
                    elseif ($TaskRunResult -match "failed|exit code") {
                        $PostCheckResult = "Registry trigger written, but task did not complete cleanly. Not confirmed updated."
                        $FinalStatus = "task-start-failed-not-confirmed"
                    }
                    else {
                        $PostCheckResult = "Registry trigger written and update task start attempted/accepted. Not confirmed updated."
                        $FinalStatus = "task-started-not-confirmed"
                    }

                    $RebootRequired = $false
                    $ExitCode = 0
                }
                else {
                    $PostCheckResult = "Remediation did not write registry trigger"
                    $FinalStatus = "failed"
                    $ExitCode = 6
                }

                if ($Restart -and $RegistryTriggerWritten -and -not $updatedAfter) {
                    if ($PostRebootCheck) {
                        try {
                            $scriptPath = if (-not [string]::IsNullOrWhiteSpace($PSCommandPath)) { $PSCommandPath } else { $MyInvocation.MyCommand.Path }
                            $PostRebootTaskRegisterResult = Invoke-PostRebootCheckTaskRegistration `
                                -ScriptPath $scriptPath `
                                -TaskName $PostRebootTaskName `
                                -Mode $Mode `
                                -CsvDirectory $CsvDirectory `
                                -CsvFileName $CsvFileName `
                                -LocalLogDirectory $LocalLogDirectory `
                                -DelaySeconds $PostRebootDelaySeconds
                            $PostRebootTaskRegistered = $true
                            Write-LocalLog $PostRebootTaskRegisterResult
                        }
                        catch {
                            $PostRebootTaskRegisterResult = "Failed to register post-reboot check task: $($_.Exception.Message)"
                            Write-LocalLog $PostRebootTaskRegisterResult
                            throw $PostRebootTaskRegisterResult
                        }
                    }

                    try {
                        $BitLockerSuspended = Suspend-BitLockerForOneRebootIfNeeded -CurrentStatus $BitLockerStatus
                    }
                    catch {
                        $ErrorMessage = $_.Exception.Message
                        $FinalStatus = "failed"
                        $ExitCode = 4
                        throw
                    }

                    $RebootRequired = $true
                    $RestartScheduled = $true
                    $FinalStatus = if ($PostRebootCheck) { "reboot-required-post-check-registered" } else { "reboot-required" }
                    $PostCheckResult = if ($PostRebootCheck) { "Post-reboot check task registered; machine will reboot and write a follow-up CSV row after startup." } else { $PostCheckResult }
                    $ExitCode = 1

                    Write-LocalLog "Restart requested after update trigger. Restarting in 60 seconds."
                    shutdown.exe /r /t 60 /c "Secure Boot 2023 certificate update attempted by Lansweeper. Restart requested for completion/post-check."
                }
            }
            else {
                # Safety net for any non-DryRun ShouldProcess decline.
                $RemediationAttempted = $false
                $CertificateUpdateResult = "ShouldProcess declined; no changes made"
                $PostCheckResult = "No changes made"
                $FinalStatus = "success"
                $ExitCode = 0
            }
        }
    }

    if ([string]::IsNullOrWhiteSpace($AvailableUpdatesAfter)) {
        $registryFinal = Get-SecureBootRegistryState
        $AvailableUpdatesAfter = $registryFinal.AvailableUpdates
        $UEFICA2023StatusAfter = $registryFinal.UEFICA2023Status
        $UEFICA2023ErrorAfter = $registryFinal.UEFICA2023Error
        $UEFICA2023ErrorEventAfter = $registryFinal.UEFICA2023ErrorEvent
        $SecureBootRegistryRawAfter = $registryFinal.Raw
    }
}
catch {
    if ([string]::IsNullOrWhiteSpace($ErrorMessage)) {
        $ErrorMessage = $_.Exception.Message
    }

    Write-LocalLog "ERROR: $ErrorMessage"

    if ($ExitCode -lt 3) {
        $ExitCode = 3
    }

    $FinalStatus = "failed"
}
finally {
    try {
        $DurationSeconds = [math]::Round(((Get-Date) - $ScriptStart).TotalSeconds, 2)

        $ProductionSafetySummary = "ScriptVersion=$ScriptVersion; SchemaVersion=$SchemaVersion; Mode=$Mode; DryRun=$DryRunEffective; Restart=$Restart; PostRebootCheck=$PostRebootCheck; IsPostRebootCheck=$IsPostRebootCheck; CsvDirectory=$CsvDirectory; ReportingDriveStatus=$ReportingDriveStatus"

        $row = @{
            Timestamp                     = (Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff zzz")
            ScriptVersion                 = $ScriptVersion
            SchemaVersion                 = $SchemaVersion
            ProductionSafetySummary       = $ProductionSafetySummary
            InvocationCommand             = $InvocationCommand
            ReportingDriveStatus          = $ReportingDriveStatus
            Mode                          = $Mode
            DryRun                        = [string]$DryRunEffective
            IsPostRebootCheck             = [string]$IsPostRebootCheck
            PostRebootTaskName            = $PostRebootTaskName
            PostRebootTaskRegistered      = [string]$PostRebootTaskRegistered
            PostRebootTaskRegisterResult  = $PostRebootTaskRegisterResult
            ComputerName                  = $ComputerName
            FullComputerName              = $FullComputerName
            MACAddress                    = $MACAddress
            IPAddress                     = $IPAddress
            OSCaption                     = $OSCaption
            OSVersion                     = $OSVersion
            OSBuild                       = $OSBuild
            Manufacturer                  = $Manufacturer
            Model                         = $Model
            SerialNumber                  = $SerialNumber
            BaseBoardManufacturer         = $BaseBoardManufacturer
            BaseBoardProduct              = $BaseBoardProduct
            CsProductVendor               = $CsProductVendor
            CsProductName                 = $CsProductName
            CsProductUuid                 = $CsProductUuid
            IsServer                      = [string]$IsServer
            IsVM                          = [string]$IsVM
            VMVendor                      = $VMVendor
            BootMode                      = $BootMode
            FirmwareType                  = $FirmwareType
            UEFIStatus                    = $UEFIStatus
            SecureBootStatus              = $SecureBootStatus
            FirmwareWarning               = $FirmwareWarning
            BitLockerStatus               = $BitLockerStatus
            BitLockerRawProtectionStatus  = $BitLockerRawProtectionStatus
            BitLockerSuspended            = [string]$BitLockerSuspended
            BitLockerRecoveryKey          = $BitLockerRecoveryKey
            SecureBootTaskExists          = [string]$SecureBootTaskExists
            SecureBootTaskState           = $SecureBootTaskState
            SecureBootTaskEnabledByScript = [string]$SecureBootTaskEnabledByScript
            TaskWarning                   = $TaskWarning
            AvailableUpdatesBefore        = $AvailableUpdatesBefore
            AvailableUpdatesAfter         = $AvailableUpdatesAfter
            UEFICA2023StatusBefore        = $UEFICA2023StatusBefore
            UEFICA2023StatusAfter         = $UEFICA2023StatusAfter
            UEFICA2023ErrorBefore         = $UEFICA2023ErrorBefore
            UEFICA2023ErrorAfter          = $UEFICA2023ErrorAfter
            UEFICA2023ErrorEventBefore    = $UEFICA2023ErrorEventBefore
            UEFICA2023ErrorEventAfter     = $UEFICA2023ErrorEventAfter
            SecureBootRegistryRawBefore   = $SecureBootRegistryRawBefore
            SecureBootRegistryRawAfter    = $SecureBootRegistryRawAfter
            RecentSecureBootEvents        = $RecentSecureBootEvents
            PrecheckResult                = $PrecheckResult
            RemediationAttempted          = [string]$RemediationAttempted
            RegistryTriggerWritten        = [string]$RegistryTriggerWritten
            TaskRunAttempted              = [string]$TaskRunAttempted
            TaskRunResult                 = $TaskRunResult
            CertificateUpdateResult       = $CertificateUpdateResult
            RebootRequired                = [string]$RebootRequired
            RestartScheduled              = [string]$RestartScheduled
            PostCheckResult               = $PostCheckResult
            FinalStatus                   = $FinalStatus
            ErrorMessage                  = $ErrorMessage
            DurationSeconds               = [string]$DurationSeconds
        }

        Write-AppendOnlyCsvRow -Row $row
    }
    catch {
        Write-LocalLog "CSV logging failure: $($_.Exception.Message)"

        if ($ExitCode -eq 0 -or $ExitCode -eq 1 -or $ExitCode -eq 2) {
            $ExitCode = 5
        }
    }

    Write-LocalLog "Script finished. ExitCode=$ExitCode FinalStatus=$FinalStatus"
    exit $ExitCode
}
