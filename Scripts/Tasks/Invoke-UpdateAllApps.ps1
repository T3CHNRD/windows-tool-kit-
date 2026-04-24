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

$stdoutPath = Join-Path $env:TEMP ("winget-upgrade-{0}-stdout.log" -f [guid]::NewGuid().ToString('N'))
$stderrPath = Join-Path $env:TEMP ("winget-upgrade-{0}-stderr.log" -f [guid]::NewGuid().ToString('N'))

try {
    $process = Start-Process -FilePath 'winget.exe' `
        -ArgumentList 'upgrade --all --include-unknown --accept-source-agreements --accept-package-agreements --disable-interactivity' `
        -Wait `
        -PassThru `
        -NoNewWindow `
        -RedirectStandardOutput $stdoutPath `
        -RedirectStandardError $stderrPath

    if (Test-Path -LiteralPath $stdoutPath) {
        Get-Content -LiteralPath $stdoutPath | Write-Output
    }

    if (Test-Path -LiteralPath $stderrPath) {
        Get-Content -LiteralPath $stderrPath | Write-Output
    }

    if ($process.ExitCode -ne 0) {
        throw "winget upgrade failed with exit code $($process.ExitCode)."
    }
}
finally {
    Remove-Item -LiteralPath $stdoutPath, $stderrPath -Force -ErrorAction SilentlyContinue
}

Write-Output 'App update run completed.'
