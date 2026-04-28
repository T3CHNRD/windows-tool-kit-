[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Write-Output 'Starting local open ports and services audit.'
Write-Output 'Scope: local listening TCP ports only. This does not scan remote systems.'

if (-not (Get-Command Get-NetTCPConnection -ErrorAction SilentlyContinue)) {
    throw 'Get-NetTCPConnection is not available on this system.'
}

$listeners = @(Get-NetTCPConnection -State Listen -ErrorAction Stop |
    Sort-Object LocalPort, LocalAddress)

if ($listeners.Count -eq 0) {
    Write-Output 'No listening TCP ports were reported.'
    return
}

Write-Output "Listening TCP endpoints detected: $($listeners.Count)"
$summary = $listeners | Group-Object LocalPort | Sort-Object {[int]$_.Name}
foreach ($group in $summary) {
    $port = [int]$group.Name
    $sample = $group.Group | Select-Object -First 1
    $processName = 'Unknown'
    $serviceNames = @()

    try {
        $proc = Get-Process -Id $sample.OwningProcess -ErrorAction Stop
        $processName = $proc.ProcessName
    }
    catch {}

    try {
        $serviceNames = @(Get-CimInstance Win32_Service -Filter "ProcessId=$($sample.OwningProcess)" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name)
    }
    catch {}

    $addresses = (($group.Group | Select-Object -ExpandProperty LocalAddress -Unique) -join ', ')
    $serviceText = if ($serviceNames.Count -gt 0) { $serviceNames -join ', ' } else { 'n/a' }
    $risk = if ($addresses -match '0\.0\.0\.0|::' -and $port -notin @(135, 139, 445)) { 'REVIEW' } else { 'INFO' }
    Write-Output ("[{0}] Port {1}: Process={2} PID={3} Services={4} Addresses={5}" -f $risk, $port, $processName, $sample.OwningProcess, $serviceText, $addresses)
}

Write-Output 'Firewall profile summary:'
if (Get-Command Get-NetFirewallProfile -ErrorAction SilentlyContinue) {
    foreach ($profile in Get-NetFirewallProfile) {
        Write-Output ("Firewall {0}: Enabled={1}, DefaultInbound={2}, DefaultOutbound={3}" -f $profile.Name, $profile.Enabled, $profile.DefaultInboundAction, $profile.DefaultOutboundAction)
    }
}
else {
    Write-Output 'Get-NetFirewallProfile unavailable.'
}

Write-Output 'Open ports and services audit complete.'
