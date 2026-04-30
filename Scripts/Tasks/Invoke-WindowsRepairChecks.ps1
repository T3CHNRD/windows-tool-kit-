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
        [Parameter(Mandatory = $true)][string]$Label,
        [int]$TimeoutSeconds = 1800,
        [switch]$NoThrow
    )

    $stdoutPath = Join-Path $env:TEMP ("{0}-{1}-stdout.log" -f $Label, [guid]::NewGuid().ToString('N'))
    $stderrPath = Join-Path $env:TEMP ("{0}-{1}-stderr.log" -f $Label, [guid]::NewGuid().ToString('N'))

    try {
        $process = Start-Process -FilePath $FilePath `
            -ArgumentList $Arguments `
            -PassThru `
            -NoNewWindow `
            -RedirectStandardOutput $stdoutPath `
            -RedirectStandardError $stderrPath

        if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
            try { Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue } catch {}
            Write-Output "$Label timed out after $TimeoutSeconds seconds."
            if (-not $NoThrow) {
                throw "$Label timed out after $TimeoutSeconds seconds."
            }
            return 1
        }

        if (Test-Path -LiteralPath $stdoutPath) {
            Get-Content -LiteralPath $stdoutPath | Write-Output
        }

        if (Test-Path -LiteralPath $stderrPath) {
            Get-Content -LiteralPath $stderrPath | Write-Output
        }

        if ($process.ExitCode -ne 0 -and -not $NoThrow) {
            throw "{0} failed with exit code {1}." -f $Label, $process.ExitCode
        }

        return [int]$process.ExitCode
    }
    finally {
        Remove-Item -LiteralPath $stdoutPath, $stderrPath -Force -ErrorAction SilentlyContinue
    }
}

if (-not (Test-IsAdmin)) {
    throw 'Windows repair checks require administrator rights.'
}

Write-Output 'Step 1/3: SFC ScanNow'
$sfcExit = Invoke-CheckedConsoleTool -FilePath 'sfc.exe' -Arguments '/scannow' -Label 'SFC-ScanNow' -TimeoutSeconds 1800 -NoThrow

if ($sfcExit -eq 0) {
    Write-Output 'SFC completed successfully. No DISM repair was required.'
}
else {
    Write-Output "SFC returned exit code $sfcExit. Running DISM component store repair, then verifying with SFC again."

    Write-Output 'Step 2/3: DISM RestoreHealth'
    $dismExit = Invoke-CheckedConsoleTool -FilePath 'dism.exe' -Arguments '/Online /Cleanup-Image /RestoreHealth' -Label 'DISM-RestoreHealth' -TimeoutSeconds 1800 -NoThrow

    if ($dismExit -eq 0) {
        Write-Output 'DISM RestoreHealth completed successfully. Running SFC again to verify repairs.'
        Write-Output 'Step 3/3: SFC verification'
        $verifyExit = Invoke-CheckedConsoleTool -FilePath 'sfc.exe' -Arguments '/scannow' -Label 'SFC-Verify' -TimeoutSeconds 1800 -NoThrow
        if ($verifyExit -ne 0) {
            throw "SFC verification completed with exit code $verifyExit. Review CBS.log for unresolved repairs."
        }
    }
    else {
        throw "DISM RestoreHealth completed with exit code $dismExit. Internet/source files may be required."
    }
}

Write-Output 'Windows repair checks completed.'
