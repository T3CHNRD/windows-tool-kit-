Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'
Write-Output 'BERET-inspired browser extension review: lists installed browser extensions for manual trust review.'
$locations = @(
    "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Extensions",
    "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Extensions",
    "$env:APPDATA\Mozilla\Firefox\Profiles"
)
foreach($loc in $locations){
    Write-Output "--- $loc ---"
    if(Test-Path $loc){ Get-ChildItem -Path $loc -Directory -Recurse -Depth 2 -ErrorAction SilentlyContinue | Select-Object FullName,LastWriteTime | Format-Table -AutoSize | Out-String | Write-Output } else { Write-Output 'Not found.' }
}
Write-Output 'Completed: browser extension inventory. Skipped: removal. Failed: none unless errors listed above.'
exit 0
