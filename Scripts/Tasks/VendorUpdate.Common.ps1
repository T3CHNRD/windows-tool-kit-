[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-TaskToolkitRoot {
    return (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
}

function Get-TaskSettings {
    $settingsPath = Join-Path (Get-TaskToolkitRoot) 'Config\Toolkit.Settings.psd1'
    if (Test-Path -LiteralPath $settingsPath) {
        return Import-PowerShellDataFile -Path $settingsPath
    }

    return @{}
}

function Test-TaskIsAdmin {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Ensure-TaskAdmin {
    if (-not (Test-TaskIsAdmin)) {
        throw 'Administrator rights are required for this task.'
    }
}

function Ensure-TaskDirectory {
    param(
        [Parameter(Mandatory = $true)][string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }

    return $Path
}

function Get-UpdateWorkspace {
    $root = Get-TaskToolkitRoot
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $basePath = Ensure-TaskDirectory -Path (Join-Path $root 'Logs\UpdateTools')
    $runPath = Ensure-TaskDirectory -Path (Join-Path $basePath $stamp)

    return [pscustomobject]@{
        Root = $root
        BasePath = $basePath
        RunPath = $runPath
        DownloadPath = (Ensure-TaskDirectory -Path (Join-Path $runPath 'Downloads'))
        ReportPath = (Ensure-TaskDirectory -Path (Join-Path $runPath 'Reports'))
    }
}

function Resolve-FirstExistingPath {
    param(
        [Parameter(Mandatory = $true)][string[]]$Candidates
    )

    foreach ($candidate in $Candidates) {
        if (-not $candidate) {
            continue
        }

        if (Test-Path -LiteralPath $candidate) {
            return (Resolve-Path -LiteralPath $candidate).ProviderPath
        }
    }

    return $null
}

function Get-SystemIdentity {
    $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem
    $bios = Get-CimInstance -ClassName Win32_BIOS

    $manufacturerRaw = [string]$computerSystem.Manufacturer
    $manufacturerKey = switch -Regex ($manufacturerRaw) {
        'Dell' { 'Dell'; break }
        'HP|Hewlett' { 'HP'; break }
        'Lenovo' { 'Lenovo'; break }
        'Framework' { 'Framework'; break }
        default { 'Unsupported' }
    }

    return [pscustomobject]@{
        Manufacturer = $manufacturerRaw
        ManufacturerKey = $manufacturerKey
        Model = [string]$computerSystem.Model
        BIOSVersion = [string]($bios.SMBIOSBIOSVersion -join ', ')
    }
}

function Invoke-LoggedProcess {
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [string]$Arguments = '',
        [Parameter(Mandatory = $true)][string]$Label,
        [string[]]$AllowedExitCodes = @('0')
    )

    $workspace = Get-UpdateWorkspace
    $stdoutPath = Join-Path $workspace.RunPath "$Label-stdout.log"
    $stderrPath = Join-Path $workspace.RunPath "$Label-stderr.log"

    Write-Output ("Running {0}: {1} {2}" -f $Label, $FilePath, $Arguments)

    $process = Start-Process -FilePath $FilePath `
        -ArgumentList $Arguments `
        -Wait `
        -PassThru `
        -NoNewWindow `
        -RedirectStandardOutput $stdoutPath `
        -RedirectStandardError $stderrPath

    if (Test-Path -LiteralPath $stdoutPath) {
        Get-Content -LiteralPath $stdoutPath | Write-Output
    }

    if (Test-Path -LiteralPath $stderrPath) {
        $stderrLines = Get-Content -LiteralPath $stderrPath
        if ($stderrLines) {
            $stderrLines | Write-Output
        }
    }

    if ($AllowedExitCodes -notcontains ([string]$process.ExitCode)) {
        throw "{0} failed with exit code {1}. Review logs in {2}." -f $Label, $process.ExitCode, $workspace.RunPath
    }

    Write-Output ("{0} finished with exit code {1}." -f $Label, $process.ExitCode)
    return $workspace
}

function Ensure-PowerShellGalleryModule {
    param(
        [Parameter(Mandatory = $true)][string]$Name
    )

    if (Get-Module -ListAvailable -Name $Name) {
        return
    }

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Install-Module -Name $Name -Repository PSGallery -Scope CurrentUser -Force -AllowClobber -AcceptLicense
}

function Get-DellCommandUpdateCliPath {
    # FIX: MED-10 - dynamically discover Dell Command Update when versioned install paths change.
    $knownPath = Resolve-FirstExistingPath -Candidates @(
        'C:\Program Files\Dell\CommandUpdate\dcu-cli.exe',
        'C:\Program Files (x86)\Dell\CommandUpdate\dcu-cli.exe'
    )
    if ($knownPath) { return $knownPath }

    $discovered = Get-ChildItem -Path 'C:\Program Files', 'C:\Program Files (x86)' -Filter 'dcu-cli.exe' -Recurse -ErrorAction SilentlyContinue |
        Select-Object -First 1 -ExpandProperty FullName
    return $discovered
}

function Ensure-HpImageAssistantPath {
    $existing = Resolve-FirstExistingPath -Candidates @(
        'C:\Program Files\HP\HPIA\HPImageAssistant.exe',
        'C:\Program Files (x86)\HP\HPIA\HPImageAssistant.exe',
        (Join-Path (Get-TaskToolkitRoot) 'Tools\HP\HPIA\HPImageAssistant.exe')
    )

    if ($existing) {
        return $existing
    }

    Ensure-PowerShellGalleryModule -Name 'HPCMSL'
    Import-Module HPCMSL -Force

    if (-not (Get-Command -Name Install-HPImageAssistant -ErrorAction SilentlyContinue)) {
        throw 'HPCMSL is installed but Install-HPImageAssistant is not available.'
    }

    $installRoot = Ensure-TaskDirectory -Path (Join-Path (Get-TaskToolkitRoot) 'Tools\HP\HPIA')
    Install-HPImageAssistant -Extract -DestinationPath $installRoot | Out-Null

    $installed = Resolve-FirstExistingPath -Candidates @(
        (Join-Path $installRoot 'HPImageAssistant.exe')
    )

    if (-not $installed) {
        throw 'HP Image Assistant could not be installed automatically.'
    }

    return $installed
}

function Get-LenovoSystemUpdatePath {
    # FIX: MED-10 - dynamically discover Lenovo System Update across versioned install paths.
    $knownPath = Resolve-FirstExistingPath -Candidates @(
        'C:\Program Files (x86)\Lenovo\System Update\tvsu.exe',
        'C:\Program Files\Lenovo\System Update\tvsu.exe'
    )
    if ($knownPath) { return $knownPath }

    $discovered = Get-ChildItem -Path 'C:\Program Files', 'C:\Program Files (x86)' -Filter 'tvsu.exe' -Recurse -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -match 'Lenovo' } |
        Select-Object -First 1 -ExpandProperty FullName
    return $discovered
}

function Invoke-DellVendorUpdates {
    param(
        [Parameter(Mandatory = $true)][string]$UpdateType
    )

    $settings = Get-TaskSettings
    $cliPath = Get-DellCommandUpdateCliPath
    if (-not $cliPath) {
        $url = $settings.UpdateTools.DellCommandUpdatePage
        throw "Dell Command | Update is not installed. Install it from Dell's official support page first: $url"
    }

    $workspace = Get-UpdateWorkspace
    $logPath = Join-Path $workspace.RunPath 'DellCommandUpdate.log'
    $arguments = "/applyUpdates -silent -reboot=disable -autoSuspendBitLocker=enable -updateType=$UpdateType -outputLog=`"$logPath`""
    Invoke-LoggedProcess -FilePath $cliPath -Arguments $arguments -Label 'DellVendorUpdate' -AllowedExitCodes @('0', '1', '3010') | Out-Null
}

function Invoke-HpVendorUpdates {
    param(
        [Parameter(Mandatory = $true)][string]$Category
    )

    $workspace = Get-UpdateWorkspace
    $hpiaPath = Ensure-HpImageAssistantPath
    $arguments = @(
        '/Operation:Analyze'
        '/Selection:All'
        '/Action:Install'
        "/Category:$Category"
        "/ReportFolder:`"$($workspace.ReportPath)`""
        "/SoftpaqDownloadFolder:`"$($workspace.DownloadPath)`""
        '/Silent'
        '/AutoCleanup'
    ) -join ' '

    Invoke-LoggedProcess -FilePath $hpiaPath -Arguments $arguments -Label 'HPImageAssistant' -AllowedExitCodes @('0', '256', '257', '3010') | Out-Null
}

function Invoke-LenovoVendorUpdates {
    param(
        [Parameter(Mandatory = $true)][string]$PackageTypes
    )

    $settings = Get-TaskSettings
    $systemUpdatePath = Get-LenovoSystemUpdatePath
    if (-not $systemUpdatePath) {
        $url = $settings.UpdateTools.LenovoSystemUpdatePage
        throw "Lenovo System Update is not installed. Install it from Lenovo's official support page first: $url"
    }

    $policyPath = 'HKLM:\Software\Policies\Lenovo\System Update\UserSettings\General'
    if (-not (Test-Path -LiteralPath $policyPath)) {
        New-Item -Path $policyPath -Force | Out-Null
    }

    $commandLine = "-search A -action INSTALL -packagetypes $PackageTypes -noicon -noreboot -includerebootpackages 1,3,4,5"
    Set-ItemProperty -Path $policyPath -Name AdminCommandLine -Type String -Value $commandLine

    Invoke-LoggedProcess -FilePath $systemUpdatePath -Arguments '/CM' -Label 'LenovoSystemUpdate' -AllowedExitCodes @('0', '3010') | Out-Null
}

function Invoke-FrameworkVendorUpdates {
    param(
        [Parameter(Mandatory = $true)][ValidateSet('BIOS', 'Firmware', 'Drivers')]$Mode
    )

    $settings = Get-TaskSettings
    $supportUrl = $settings.UpdateTools.FrameworkBiosDriversPage
    if (-not $supportUrl) {
        $supportUrl = 'https://knowledgebase.frame.work/bios-and-drivers-downloads-rJ3PaCexh'
    }

    Write-Output "Framework detected. Opening Framework's official BIOS and driver downloads page."
    Write-Output "Framework does not currently provide a silent vendor CLI in this toolkit. Automated install skipped; use the official Framework workflow for $Mode."
    Write-Output "Official Framework support page: $supportUrl"

    try {
        Start-Process $supportUrl
        Write-Output 'Completed: opened Framework official support page.'
    }
    catch {
        Write-Output "Skipped opening browser automatically: $($_.Exception.Message)"
    }
}

function Invoke-VendorMaintenanceUpdate {
    param(
        [Parameter(Mandatory = $true)][ValidateSet('BIOS', 'Firmware', 'Drivers')]$Mode
    )

    Ensure-TaskAdmin
    $identity = Get-SystemIdentity

    Write-Output ("Detected manufacturer: {0}" -f $identity.Manufacturer)
    Write-Output ("Detected model: {0}" -f $identity.Model)
    Write-Output ("Detected BIOS version: {0}" -f $identity.BIOSVersion)

    switch ($identity.ManufacturerKey) {
        'Dell' {
            $updateType = switch ($Mode) {
                'BIOS' { 'bios' }
                'Firmware' { 'firmware' }
                'Drivers' { 'driver' }
            }
            Invoke-DellVendorUpdates -UpdateType $updateType
            break
        }
        'HP' {
            $category = switch ($Mode) {
                'BIOS' { 'BIOS' }
                'Firmware' { 'Firmware' }
                'Drivers' { 'Drivers' }
            }
            Invoke-HpVendorUpdates -Category $category
            break
        }
        'Lenovo' {
            $packageTypes = switch ($Mode) {
                'BIOS' { '3' }
                'Firmware' { '4' }
                'Drivers' { '2' }
            }
            Invoke-LenovoVendorUpdates -PackageTypes $packageTypes
            break
        }
        'Framework' {
            Invoke-FrameworkVendorUpdates -Mode $Mode
            break
        }
        default {
            $settings = Get-TaskSettings
            $supported = ($settings.UpdateTools.SupportedManufacturers -join ', ')
            throw "This vendor is not currently supported for vendor update automation. Supported manufacturers: $supported. Detected manufacturer: $($identity.Manufacturer)"
        }
    }

    Write-Output ("{0} update workflow completed for {1}." -f $Mode, $identity.ManufacturerKey)
}

function Get-PendingWindowsUpdates {
    $session = New-Object -ComObject Microsoft.Update.Session
    $searcher = $session.CreateUpdateSearcher()
    $searchResult = $searcher.Search('IsInstalled=0')
    $updates = @()

    for ($index = 0; $index -lt $searchResult.Updates.Count; $index++) {
        $update = $searchResult.Updates.Item($index)
        $kbIds = @($update.KBArticleIDs | ForEach-Object { 'KB{0}' -f $_ })
        $updates += [pscustomobject]@{
            Title = [string]$update.Title
            Identity = [string]$update.Identity.UpdateID
            KB = @($kbIds)
            IsHidden = [bool]$update.IsHidden
            UpdateObject = $update
        }
    }

    return $updates
}

function Test-WindowsUpdateSkipMatch {
    param(
        [Parameter(Mandatory = $true)]$UpdateRecord,
        [Parameter(Mandatory = $true)][string[]]$SkipEntries
    )

    foreach ($entry in $SkipEntries) {
        if (-not $entry) {
            continue
        }

        $trimmed = $entry.Trim()
        if (-not $trimmed) {
            continue
        }

        if ($trimmed -match '^KB\d+$') {
            if ($UpdateRecord.KB -contains $trimmed.ToUpperInvariant()) {
                return $true
            }
        }
        elseif ($UpdateRecord.Title -like "*$trimmed*") {
            return $true
        }
    }

    return $false
}

function Invoke-WindowsUpdateInstallation {
    param(
        [string]$SkipSelectionFile
    )

    Ensure-TaskAdmin

    $availableUpdates = @(Get-PendingWindowsUpdates)
    if ($availableUpdates.Count -eq 0) {
        Write-Output 'No pending Windows updates were found.'
        return
    }

    Write-Output ("Pending updates found: {0}" -f $availableUpdates.Count)

    $skipEntries = @()
    if ($SkipSelectionFile -and (Test-Path -LiteralPath $SkipSelectionFile)) {
        $skipEntries = @(Get-Content -LiteralPath $SkipSelectionFile | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }

    $installCollection = New-Object -ComObject Microsoft.Update.UpdateColl
    foreach ($record in $availableUpdates) {
        if ($skipEntries.Count -gt 0 -and (Test-WindowsUpdateSkipMatch -UpdateRecord $record -SkipEntries $skipEntries)) {
            try {
                $record.UpdateObject.IsHidden = $true
                Write-Output ("Skipping and hiding update: {0}" -f $record.Title)
            }
            catch {
                Write-Output ("Could not hide update '{0}': {1}" -f $record.Title, $_.Exception.Message)
            }
            continue
        }

        if (-not $record.UpdateObject.EulaAccepted) {
            $record.UpdateObject.AcceptEula()
        }

        [void]$installCollection.Add($record.UpdateObject)
    }

    if ($installCollection.Count -eq 0) {
        Write-Output 'No Windows updates remain after applying skip rules.'
        return
    }

    $session = New-Object -ComObject Microsoft.Update.Session
    $downloader = $session.CreateUpdateDownloader()
    $downloader.Updates = $installCollection
    $downloadResult = $downloader.Download()
    Write-Output ("Windows Update download result code: {0}" -f $downloadResult.ResultCode)

    $installer = $session.CreateUpdateInstaller()
    $installer.Updates = $installCollection
    $installResult = $installer.Install()
    Write-Output ("Windows Update install result code: {0}" -f $installResult.ResultCode)

    for ($index = 0; $index -lt $installCollection.Count; $index++) {
        $update = $installCollection.Item($index)
        $result = $installResult.GetUpdateResult($index)
        Write-Output ("Update result: {0} => {1}" -f $update.Title, $result.ResultCode)
    }

    if ($installResult.RebootRequired) {
        Write-Output 'A reboot is required to finish Windows updates.'
    }
}
