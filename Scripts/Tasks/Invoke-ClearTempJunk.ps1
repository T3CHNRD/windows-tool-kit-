[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$targets = @($env:TEMP, "$env:WINDIR\Temp")

foreach ($target in $targets) {
    if (-not (Test-Path $target)) {
        Write-Output "Skipping missing path: $target"
        continue
    }

    Write-Output "Cleaning: $target"
    Get-ChildItem -Path $target -Force -ErrorAction SilentlyContinue | ForEach-Object {
        try {
            Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction Stop
        }
        catch {
            Write-Output "Could not remove: $($_.FullName) - $($_.Exception.Message)"
        }
    }
}

Write-Output 'Temp and junk cleanup complete.'
