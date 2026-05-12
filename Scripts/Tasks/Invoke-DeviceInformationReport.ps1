Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'
$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$logDir = Join-Path $root 'Logs'; New-Item -Path $logDir -ItemType Directory -Force | Out-Null
$report = Join-Path $logDir ("DeviceInformation-{0:yyyyMMdd-HHmmss}.txt" -f (Get-Date))
Write-Output 'Collecting KillerTools device-information inspired Windows hardware/software report.'
$lines = @()
$lines += "Device Information Report - $(Get-Date)"
$lines += 'Credit: inspired by KillerTools Device Information (https://killertools.net/device-information)'
$cs = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue
$bios = Get-CimInstance Win32_BIOS -ErrorAction SilentlyContinue
$os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
$cpu = Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1
$lines += "Computer: $env:COMPUTERNAME"
$lines += "Manufacturer/Model: $($cs.Manufacturer) $($cs.Model)"
$lines += "BIOS: $($bios.SMBIOSBIOSVersion) $($bios.ReleaseDate)"
$lines += "OS: $($os.Caption) build $($os.BuildNumber)"
$lines += "CPU: $($cpu.Name)"
$lines += "RAM GB: $([math]::Round($cs.TotalPhysicalMemory/1GB,2))"
$lines += ''; $lines += 'Disks:'; $lines += (Get-Disk | Select Number,FriendlyName,BusType,OperationalStatus,@{n='SizeGB';e={[math]::Round($_.Size/1GB,2)}} | Format-Table -AutoSize | Out-String)
$lines += ''; $lines += 'Network adapters:'; $lines += (Get-NetAdapter -ErrorAction SilentlyContinue | Select Name,Status,MacAddress,LinkSpeed | Format-Table -AutoSize | Out-String)
$lines | Set-Content -LiteralPath $report -Encoding UTF8
$lines | ForEach-Object { Write-Output $_ }
Write-Output "Report saved: $report"
exit 0
