[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Write-Output 'Starting Windows security event review.'
Write-Output 'Defensive objective: spot signs of repeated failed logons, privilege use, account changes, and lockouts.'

$startTime = (Get-Date).AddDays(-7)
$checks = @(
    @{ Id = 4625; Name = 'Failed logons'; Limit = 20 },
    @{ Id = 4624; Name = 'Successful logons'; Limit = 12 },
    @{ Id = 4672; Name = 'Special privilege logons'; Limit = 12 },
    @{ Id = 4720; Name = 'User account created'; Limit = 20 },
    @{ Id = 4726; Name = 'User account deleted'; Limit = 20 },
    @{ Id = 4740; Name = 'Account lockout'; Limit = 20 },
    @{ Id = 7045; Name = 'Service installed'; Limit = 20 }
)

foreach ($check in $checks) {
    Write-Output "Checking event ID $($check.Id): $($check.Name)"
    try {
        $events = @(Get-WinEvent -FilterHashtable @{ LogName = 'Security'; Id = $check.Id; StartTime = $startTime } -MaxEvents $check.Limit -ErrorAction Stop)
    }
    catch {
        Write-Output "[WARN] Could not read event ID $($check.Id): $($_.Exception.Message)"
        continue
    }

    if ($events.Count -eq 0) {
        Write-Output "[OK] $($check.Name): no events found in the last 7 days."
        continue
    }

    Write-Output "[REVIEW] $($check.Name): $($events.Count) recent event(s) found. Showing newest first."
    foreach ($event in $events) {
        $message = ($event.Message -replace '\s+', ' ').Trim()
        if ($message.Length -gt 280) {
            $message = $message.Substring(0, 280) + '...'
        }
        Write-Output ("    {0:u} ID={1} Provider={2} {3}" -f $event.TimeCreated, $event.Id, $event.ProviderName, $message)
    }
}

Write-Output 'Security event review complete.'
