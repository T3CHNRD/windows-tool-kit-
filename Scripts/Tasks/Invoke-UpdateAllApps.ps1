[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Test-IsAdmin {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-IsAdmin)) {
    throw 'Updating all installed apps requires administrator rights.'
}

Write-Output 'Updating all supported apps with winget...'
if (-not (Get-Command winget.exe -ErrorAction SilentlyContinue)) {
    throw 'winget.exe was not found. Install or repair Microsoft App Installer from Microsoft Store, then rerun app updates.'
}

$stdoutPath = Join-Path $env:TEMP ("winget-upgrade-{0}-stdout.log" -f [guid]::NewGuid().ToString('N'))
$stderrPath = Join-Path $env:TEMP ("winget-upgrade-{0}-stderr.log" -f [guid]::NewGuid().ToString('N'))
$stdoutLines = @()
$stderrLines = @()

try {
    $process = Start-Process -FilePath 'winget.exe' `
        -ArgumentList 'upgrade --all --include-unknown --accept-source-agreements --accept-package-agreements --disable-interactivity' `
        -Wait `
        -PassThru `
        -NoNewWindow `
        -RedirectStandardOutput $stdoutPath `
        -RedirectStandardError $stderrPath

    if (Test-Path -LiteralPath $stdoutPath) {
        $stdoutLines = @(Get-Content -LiteralPath $stdoutPath)
        $stdoutLines | Write-Output
    }

    if (Test-Path -LiteralPath $stderrPath) {
        $stderrLines = @(Get-Content -LiteralPath $stderrPath)
        $stderrLines | Write-Output
    }

    if ($process.ExitCode -ne 0) {
        throw "winget upgrade failed with exit code $($process.ExitCode)."
    }

    # FIX: MED-04 - provide a clear completion summary for Update tab tasks.
    $completedCount = @($stdoutLines | Where-Object { $_ -match '\bSuccessfully installed\b|\bSuccessfully upgraded\b|\bUpgraded\b' }).Count
    $skippedCount = @($stdoutLines | Where-Object { $_ -match '\bNo available upgrade\b|\bNo installed package found\b|\bNo applicable update\b' }).Count
    Write-Output ("App update summary: Completed={0}; Skipped/unchanged={1}; Failed=0; ExitCode={2}" -f $completedCount, $skippedCount, $process.ExitCode)
}
finally {
    Remove-Item -LiteralPath $stdoutPath, $stderrPath -Force -ErrorAction SilentlyContinue
}

Write-Output 'App update run completed.'
