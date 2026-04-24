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
    foreach ($item in @(Get-ChildItem -Path $target -Force -ErrorAction SilentlyContinue)) {
        $itemPath = $null
        if ($item -and $item.PSObject.Properties['FullName']) {
            $itemPath = [string]$item.FullName
        }
        elseif ($item) {
            $itemPath = [string]$item
        }

        if ([string]::IsNullOrWhiteSpace($itemPath)) {
            continue
        }

        try {
            Remove-Item -LiteralPath $itemPath -Recurse -Force -ErrorAction Stop
        }
        catch {
            Write-Output "Could not remove: $itemPath - $($_.Exception.Message)"
        }
    }
}

Write-Output 'Temp and junk cleanup complete.'
