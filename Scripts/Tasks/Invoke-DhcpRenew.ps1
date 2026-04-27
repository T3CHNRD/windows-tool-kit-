[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Test-IsAdmin {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-IsAdmin)) {
    throw 'Release/renew IP and DHCP reset requires administrator rights.'
}

$completed = New-Object System.Collections.Generic.List[string]
$skipped = New-Object System.Collections.Generic.List[string]
$failed = New-Object System.Collections.Generic.List[string]

function Invoke-DhcpStep {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][scriptblock]$Action,
        [switch]$Optional
    )

    Write-Output "Starting: $Name"
    try {
        & $Action
        [void]$completed.Add($Name)
        Write-Output "Completed: $Name"
    }
    catch {
        if ($Optional) {
            [void]$skipped.Add("$Name - $($_.Exception.Message)")
            Write-Output "Skipped: $Name - $($_.Exception.Message)"
            return
        }

        [void]$failed.Add("$Name - $($_.Exception.Message)")
        Write-Output "Failed: $Name - $($_.Exception.Message)"
    }
}

Write-Output 'Starting IP release/renew and DHCP reset workflow.'

Invoke-DhcpStep -Name 'ipconfig /release' -Action {
    ipconfig.exe /release | Write-Output
}

Invoke-DhcpStep -Name 'restart DHCP Client service' -Action {
    Restart-Service -Name Dhcp -Force -ErrorAction Stop
}

Invoke-DhcpStep -Name 'ipconfig /renew' -Action {
    ipconfig.exe /renew | Write-Output
}

Invoke-DhcpStep -Name 'flush DNS resolver cache' -Action {
    ipconfig.exe /flushdns | Write-Output
}

Invoke-DhcpStep -Name 'register DNS records' -Action {
    ipconfig.exe /registerdns | Write-Output
} -Optional

Write-Output 'Network DHCP summary:'
Write-Output ("Completed: {0}" -f ($(if ($completed.Count -gt 0) { $completed -join '; ' } else { 'none' })))
Write-Output ("Skipped: {0}" -f ($(if ($skipped.Count -gt 0) { $skipped -join '; ' } else { 'none' })))
Write-Output ("Failed: {0}" -f ($(if ($failed.Count -gt 0) { $failed -join '; ' } else { 'none' })))

if ($failed.Count -gt 0) {
    throw "One or more DHCP reset steps failed: $($failed -join '; ')"
}

Write-Output 'IP release/renew and DHCP reset workflow completed.'
