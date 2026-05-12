Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'
Write-Output 'MACE-inspired cleanup: safe cache cleanup pass with detailed completed/skipped/failed summary.'
$completed = New-Object System.Collections.Generic.List[string]
$skipped = New-Object System.Collections.Generic.List[string]
$failed = New-Object System.Collections.Generic.List[string]
$targets = @($env:TEMP, "$env:LOCALAPPDATA\Temp", 'C:\Windows\Temp') | Where-Object { $_ -and (Test-Path $_) } | Select-Object -Unique
foreach($target in $targets){
    Write-Output "Cleaning target: $target"
    try {
        Get-ChildItem -LiteralPath $target -Force -ErrorAction SilentlyContinue | ForEach-Object {
            try { Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction Stop; $completed.Add($_.FullName) } catch { $skipped.Add($_.FullName) }
        }
    } catch { $failed.Add("$target :: $($_.Exception.Message)") }
}
try { Clear-RecycleBin -Force -ErrorAction Stop; $completed.Add('Recycle Bin') } catch { $skipped.Add('Recycle Bin: ' + $_.Exception.Message) }
Write-Output "Cleanup summary: Completed=$($completed.Count); Skipped=$($skipped.Count); Failed=$($failed.Count)"
if($skipped.Count){ Write-Output 'Skipped examples:'; $skipped | Select-Object -First 20 | ForEach-Object { Write-Output "  $_" } }
if($failed.Count){ Write-Output 'Failed:'; $failed | ForEach-Object { Write-Output "  $_" } }
exit $(if($failed.Count -gt 0){1}else{0})
