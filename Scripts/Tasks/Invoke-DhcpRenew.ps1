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

function Invoke-NativeDhcpCommand {
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [Parameter(Mandatory = $true)][string[]]$Arguments
    )

    & $FilePath @Arguments | Write-Output
    if ($LASTEXITCODE -ne 0) {
        throw "$FilePath $($Arguments -join ' ') exited with code $LASTEXITCODE."
    }
}

Write-Output 'Starting IP release/renew and DHCP reset workflow.'

$primaryAdapter = $null
try {
    # FIX: MED-12 - target the primary route instead of releasing every adapter at once.
    $primaryAdapter = (Get-NetRoute -DestinationPrefix '0.0.0.0/0' -ErrorAction Stop |
        Sort-Object RouteMetric, InterfaceMetric |
        Select-Object -First 1 -ExpandProperty InterfaceAlias)
    if ($primaryAdapter) {
        Write-Output "Primary network adapter detected: $primaryAdapter"
    }
}
catch {
    Write-Output "Could not detect a primary adapter. Falling back to ipconfig default behavior: $($_.Exception.Message)"
}

$isStaticPrimary = $false
if ($primaryAdapter) {
    try {
        # FIX: MED-05 - do not force DHCP renew on a static-IP primary adapter.
        $primaryConfig = Get-NetIPConfiguration -InterfaceAlias $primaryAdapter -ErrorAction Stop
        $primaryInterface = Get-NetIPInterface -InterfaceAlias $primaryAdapter -AddressFamily IPv4 -ErrorAction Stop | Select-Object -First 1
        $isStaticPrimary = ($primaryConfig.IPv4DefaultGateway -and $primaryInterface.Dhcp -eq 'Disabled')
        if ($isStaticPrimary) {
            Write-Output "Static IPv4 configuration detected on $primaryAdapter. DHCP release/renew will be skipped to avoid breaking connectivity."
        }
    }
    catch {
        Write-Output "Could not determine DHCP/static state for ${primaryAdapter}: $($_.Exception.Message)"
    }
}

Invoke-DhcpStep -Name 'ipconfig /release' -Action {
    if ($isStaticPrimary) { throw "Skipped because $primaryAdapter uses static IPv4 configuration." }
    $arguments = if ($primaryAdapter) { @('/release', $primaryAdapter) } else { @('/release') }
    Invoke-NativeDhcpCommand -FilePath 'ipconfig.exe' -Arguments $arguments
} -Optional

Invoke-DhcpStep -Name 'ipconfig /renew' -Action {
    if ($isStaticPrimary) { throw "Skipped because $primaryAdapter uses static IPv4 configuration." }
    $arguments = if ($primaryAdapter) { @('/renew', $primaryAdapter) } else { @('/renew') }
    Invoke-NativeDhcpCommand -FilePath 'ipconfig.exe' -Arguments $arguments
} -Optional

Invoke-DhcpStep -Name 'flush DNS resolver cache' -Action {
    Invoke-NativeDhcpCommand -FilePath 'ipconfig.exe' -Arguments @('/flushdns')
}

Invoke-DhcpStep -Name 'restart DHCP Client service' -Action {
    Restart-Service -Name Dhcp -Force -ErrorAction Stop
} -Optional

Invoke-DhcpStep -Name 'register DNS records' -Action {
    Invoke-NativeDhcpCommand -FilePath 'ipconfig.exe' -Arguments @('/registerdns')
} -Optional

Write-Output 'Network DHCP summary:'
Write-Output ("Completed: {0}" -f ($(if ($completed.Count -gt 0) { $completed -join '; ' } else { 'none' })))
Write-Output ("Skipped: {0}" -f ($(if ($skipped.Count -gt 0) { $skipped -join '; ' } else { 'none' })))
Write-Output ("Failed: {0}" -f ($(if ($failed.Count -gt 0) { $failed -join '; ' } else { 'none' })))

if ($failed.Count -gt 0) {
    throw "One or more DHCP reset steps failed: $($failed -join '; ')"
}

Write-Output 'IP release/renew and DHCP reset workflow completed.'
