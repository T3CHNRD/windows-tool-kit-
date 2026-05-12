Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'
$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot); $logDir=Join-Path $root 'Logs'; New-Item $logDir -ItemType Directory -Force | Out-Null
$report = Join-Path $logDir ("KillerScanLocal-{0:yyyyMMdd-HHmmss}.csv" -f (Get-Date))
Write-Output 'KillerScan-inspired local network scan: gateway subnet discovery, ping sweep, DNS names, and common port probes.'
$config = Get-NetIPConfiguration | Where-Object { $_.IPv4DefaultGateway -and $_.IPv4Address } | Select-Object -First 1
if (-not $config) { Write-Output 'No active IPv4 gateway found. Scan skipped.'; exit 0 }
$ip = $config.IPv4Address.IPAddress; $prefix = ($ip -replace '\.\d+$','')
$ports = @(22,80,135,139,443,445,3389,5985,8080,9100)
$results = New-Object System.Collections.Generic.List[object]
1..254 | ForEach-Object {
    $target = "$prefix.$_"
    if (Test-Connection -ComputerName $target -Count 1 -Quiet -TimeoutSeconds 1) {
        $name = try { [Net.Dns]::GetHostEntry($target).HostName } catch { '' }
        $open = foreach($p in $ports){ $client=New-Object Net.Sockets.TcpClient; try { $iar=$client.BeginConnect($target,$p,$null,$null); if($iar.AsyncWaitHandle.WaitOne(200,$false)){ $client.EndConnect($iar); $p } } catch {} finally { $client.Close() } }
        $obj=[pscustomobject]@{IP=$target; Hostname=$name; OpenPorts=($open -join ',')}
        $results.Add($obj); Write-Output ("Found {0} {1} ports: {2}" -f $target,$name,($open -join ','))
    }
}
$results | Export-Csv -NoTypeInformation -Path $report
Write-Output "Scan complete. Hosts found: $($results.Count). Report: $report"
exit 0
