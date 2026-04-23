param(
    [string[]]$ComputerName = @($env:COMPUTERNAME),
    [pscredential]$Credential,
    [switch]$ApplyUpdate
)

$scriptBlock = {
    param($DoUpdate)

    $result = [ordered]@{
        ComputerName                = $env:COMPUTERNAME
        ScanStatus                  = 'Scanned'
        BIOSMode                    = $null
        SecureBootState             = $null
        SecureBootEnabled           = $false
        WindowsUEFICA2023_DB        = $false
        WindowsUEFICA2023_DBDefault = $false
        UpdateTaskExists            = $false
        RecommendedAction           = $null
        UpdateAttempted             = $false
        Error                       = $null
    }

    try {
        $cs = Get-CimInstance Win32_ComputerSystem
        $result.BIOSMode = if ($cs.BootupState) { $cs.BootupState } else { $null }
    }
    catch {
        $result.BIOSMode = $null
    }

    try {
        $sb = Confirm-SecureBootUEFI
        $result.SecureBootEnabled = [bool]$sb
        $result.SecureBootState = if ($sb) { 'On' } else { 'Off' }
    }
    catch {
        try {
            $reg = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\State" -ErrorAction Stop
            if ($reg.UEFISecureBootEnabled -eq 1) {
                $result.SecureBootEnabled = $true
                $result.SecureBootState = 'On (Registry)'
            }
            else {
                $result.SecureBootEnabled = $false
                $result.SecureBootState = 'Off (Registry)'
            }
        }
        catch {
            $result.SecureBootEnabled = $false
            $result.SecureBootState = 'Unknown'
        }
    }

    try {
        $db = (Get-SecureBootUEFI db).bytes
        $text = [System.Text.Encoding]::ASCII.GetString($db)
        $result.WindowsUEFICA2023_DB = $text -match 'Windows UEFI CA 2023'
    }
    catch {
        $result.WindowsUEFICA2023_DB = $false
    }

    try {
        $dbdefault = (Get-SecureBootUEFI dbdefault).bytes
        $textDefault = [System.Text.Encoding]::ASCII.GetString($dbdefault)
        $result.WindowsUEFICA2023_DBDefault = $textDefault -match 'Windows UEFI CA 2023'
    }
    catch {
        $result.WindowsUEFICA2023_DBDefault = $false
    }

    try {
        $task = Get-ScheduledTask -TaskName 'Secure-Boot-Update' -TaskPath '\Microsoft\Windows\PI\' -ErrorAction Stop
        if ($task) { $result.UpdateTaskExists = $true }
    }
    catch {
        $result.UpdateTaskExists = $false
    }

    if (-not $result.SecureBootEnabled) {
        $result.RecommendedAction = 'Enable UEFI Secure Boot first'
    }
    elseif ($result.WindowsUEFICA2023_DB) {
        $result.RecommendedAction = 'Already updated'
    }
    else {
        $result.RecommendedAction = 'Eligible for update'
    }

    if ($DoUpdate -and $result.SecureBootEnabled -and -not $result.WindowsUEFICA2023_DB) {
        $result.UpdateAttempted = $true

        try {
            Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot' -Name 'AvailableUpdates' -Value 0x5944 -ErrorAction Stop
        }
        catch {}

        try {
            Start-ScheduledTask -TaskName 'Secure-Boot-Update' -TaskPath '\Microsoft\Windows\PI\' -ErrorAction Stop
        }
        catch {}
    }

    [pscustomobject]$result
}

$results = foreach ($c in $ComputerName) {
    try {
        if ($c -eq $env:COMPUTERNAME -or $c -eq 'localhost' -or $c -eq '.') {
            & $scriptBlock $ApplyUpdate
        }
        elseif ($PSBoundParameters.ContainsKey('Credential')) {
            Invoke-Command -ComputerName $c -Credential $Credential -ScriptBlock $scriptBlock -ArgumentList $ApplyUpdate
        }
        else {
            Invoke-Command -ComputerName $c -ScriptBlock $scriptBlock -ArgumentList $ApplyUpdate
        }
    }
    catch {
        [pscustomobject]@{
            ComputerName                = $c
            ScanStatus                  = 'Network Scan Failed - Check Locally'
            BIOSMode                    = $null
            SecureBootState             = 'Error'
            SecureBootEnabled           = $false
            WindowsUEFICA2023_DB        = $false
            WindowsUEFICA2023_DBDefault = $false
            UpdateTaskExists            = $false
            RecommendedAction           = 'Error'
            UpdateAttempted             = $false
            Error                       = $_.Exception.Message
        }
    }
}

$results | Format-Table -AutoSize

$timestamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
$outputFolder = '\\alblnetapp02\public\IT Tracking - Requests_Projects\bootcheck_results'
$outputFile = Join-Path $outputFolder "SecureBootCheckResults_$timestamp.csv"

if (-not (Test-Path $outputFolder)) {
    New-Item -Path $outputFolder -ItemType Directory -Force | Out-Null
}

$results | Export-Csv -Path $outputFile -NoTypeInformation

Write-Host "Results saved to: $outputFile" -ForegroundColor Green

if ($ApplyUpdate) {
    Write-Host "If any targets were eligible, reboot them and run the script again to verify." -ForegroundColor Cyan
}