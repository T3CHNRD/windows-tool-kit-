Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'
Write-Output 'SHADE-inspired privacy audit. This tool reports privacy-sensitive settings and does not make changes.'
$checks = @(
    @{Name='Advertising ID'; Path='HKCU:\Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo'; Value='Enabled'},
    @{Name='Telemetry AllowTelemetry'; Path='HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection'; Value='AllowTelemetry'},
    @{Name='Location service'; Path='HKLM:\SYSTEM\CurrentControlSet\Services\lfsvc\Service\Configuration'; Value='Status'},
    @{Name='Activity history publish'; Path='HKLM:\SOFTWARE\Policies\Microsoft\Windows\System'; Value='PublishUserActivities'}
)
foreach($c in $checks){
    try { $v=(Get-ItemProperty -Path $c.Path -Name $c.Value -ErrorAction Stop).($c.Value); Write-Output "$($c.Name): $v" } catch { Write-Output "$($c.Name): Not configured or unavailable" }
}
Write-Output 'Recent files setting:'
try { Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Name Start_TrackDocs -ErrorAction Stop | Select Start_TrackDocs | Format-List | Out-String | Write-Output } catch { Write-Output 'Recent files setting unavailable.' }
Write-Output 'Completed: privacy audit. Skipped: no changes applied. Failed: none unless errors listed above.'
exit 0
