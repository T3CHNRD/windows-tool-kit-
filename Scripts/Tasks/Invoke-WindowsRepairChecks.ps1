[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Test-IsAdmin {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Invoke-CheckedConsoleTool {
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [Parameter(Mandatory = $true)][string]$Arguments,
        [Parameter(Mandatory = $true)][string]$Label
    )

    $stdoutPath = Join-Path $env:TEMP ("{0}-{1}-stdout.log" -f $Label, [guid]::NewGuid().ToString('N'))
    $stderrPath = Join-Path $env:TEMP ("{0}-{1}-stderr.log" -f $Label, [guid]::NewGuid().ToString('N'))

    try {
        $process = Start-Process -FilePath $FilePath `
            -ArgumentList $Arguments `
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
            throw "{0} failed with exit code {1}." -f $Label, $process.ExitCode
        }
    }
    finally {
        Remove-Item -LiteralPath $stdoutPath, $stderrPath -Force -ErrorAction SilentlyContinue
    }
}

if (-not (Test-IsAdmin)) {
    throw 'Windows repair checks require administrator rights.'
}

Write-Output 'Step 1/3: DISM ScanHealth'
Invoke-CheckedConsoleTool -FilePath 'dism.exe' -Arguments '/Online /Cleanup-Image /ScanHealth' -Label 'DISM-ScanHealth'

Write-Output 'Step 2/3: DISM RestoreHealth'
Invoke-CheckedConsoleTool -FilePath 'dism.exe' -Arguments '/Online /Cleanup-Image /RestoreHealth' -Label 'DISM-RestoreHealth'

Write-Output 'Step 3/3: SFC ScanNow'
Invoke-CheckedConsoleTool -FilePath 'sfc.exe' -Arguments '/scannow' -Label 'SFC-ScanNow'

Write-Output 'Windows repair checks completed.'
