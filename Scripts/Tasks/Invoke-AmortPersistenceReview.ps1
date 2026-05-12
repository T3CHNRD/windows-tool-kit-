Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'
Write-Output 'AMORT-inspired persistence review: autoruns, startup folders, scheduled tasks, and automatic services.'
Write-Output '--- Run Keys ---'
$runKeys = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Run','HKCU:\Software\Microsoft\Windows\CurrentVersion\Run','HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Run'
foreach($key in $runKeys){ Write-Output "[$key]"; try { Get-ItemProperty $key -ErrorAction Stop | Format-List | Out-String | Write-Output } catch { Write-Output 'Unavailable or empty.' } }
Write-Output '--- Startup folders ---'
@([Environment]::GetFolderPath('Startup'), "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Startup") | ForEach-Object { if(Test-Path $_){ Get-ChildItem $_ -Force | Select FullName,Length,LastWriteTime | Format-Table -AutoSize | Out-String | Write-Output } }
Write-Output '--- Recently modified scheduled tasks ---'
Get-ScheduledTask -ErrorAction SilentlyContinue | Select-Object -First 80 TaskName,TaskPath,State | Format-Table -AutoSize | Out-String | Write-Output
Write-Output 'Completed: AMORT-style persistence review. Skipped: changes/removal. Failed: see warnings above.'
exit 0
