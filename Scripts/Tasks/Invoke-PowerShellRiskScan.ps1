[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$toolkitRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$scanRoots = @(
    (Join-Path $toolkitRoot 'Scripts'),
    (Join-Path $toolkitRoot 'Modules'),
    (Join-Path $toolkitRoot 'Integrations'),
    (Join-Path $toolkitRoot 'LegacyScripts')
) | Where-Object { Test-Path -LiteralPath $_ }

$rules = @(
    @{ Severity = 'HIGH'; Pattern = '(?i)\bInvoke-Expression\b|\biex\b'; Reason = 'Dynamic code execution can hide malicious behavior or make review harder.' },
    @{ Severity = 'HIGH'; Pattern = '(?i)EncodedCommand|FromBase64String'; Reason = 'Encoded commands are commonly used to obscure PowerShell behavior.' },
    @{ Severity = 'HIGH'; Pattern = '(?i)DownloadString|DownloadFile|Invoke-WebRequest|Invoke-RestMethod'; Reason = 'Network download/execution paths should be verified and source-pinned.' },
    @{ Severity = 'MED'; Pattern = '(?i)Set-ExecutionPolicy\s+Bypass|ExecutionPolicy\s+Bypass'; Reason = 'Execution policy bypass may be legitimate in tooling but should be intentional.' },
    @{ Severity = 'MED'; Pattern = '(?i)Add-MpPreference|Set-MpPreference|DisableRealtimeMonitoring|ExclusionPath'; Reason = 'Defender preference changes and exclusions should be reviewed carefully.' },
    @{ Severity = 'MED'; Pattern = '(?i)net\s+user|New-LocalUser|Add-LocalGroupMember'; Reason = 'User/account changes are sensitive and should be documented.' },
    @{ Severity = 'LOW'; Pattern = '(?i)password\s*=|credential\s*=|ConvertTo-SecureString\s+.+-AsPlainText'; Reason = 'Potential hardcoded credential or plain-text secret handling.' }
)

Write-Output 'Starting PowerShell script-risk scan.'
Write-Output 'Scope: scans toolkit PowerShell files for risky patterns. This is SAST-style review guidance, not proof of compromise.'
Write-Output ("Scan roots: {0}" -f ($scanRoots -join '; '))

$files = @(Get-ChildItem -Path $scanRoots -Include '*.ps1','*.psm1','*.psd1' -Recurse -File -ErrorAction SilentlyContinue |
    Where-Object {
        $_.FullName -notmatch '\\Integrations\\Install-Microsoft365\\' -and
        $_.Name -ne 'Invoke-PowerShellRiskScan.ps1'
    })

Write-Output "Files scanned: $($files.Count)"
$findings = New-Object System.Collections.Generic.List[object]

foreach ($file in $files) {
    $lineNumber = 0
    foreach ($line in Get-Content -LiteralPath $file.FullName -ErrorAction SilentlyContinue) {
        $lineNumber += 1
        foreach ($rule in $rules) {
            if ($line -match $rule.Pattern) {
                $relativePath = $file.FullName.Substring($toolkitRoot.Length).TrimStart('\')
                $findings.Add([pscustomobject]@{
                    Severity = $rule.Severity
                    File = $relativePath
                    Line = $lineNumber
                    Reason = $rule.Reason
                    Text = $line.Trim()
                })
            }
        }
    }
}

if ($findings.Count -eq 0) {
    Write-Output 'No risky PowerShell patterns were detected by the current rule set.'
}
else {
    Write-Output "Findings detected: $($findings.Count)"
    foreach ($finding in ($findings | Sort-Object Severity, File, Line)) {
        Write-Output ("[{0}] {1}:{2} - {3}" -f $finding.Severity, $finding.File, $finding.Line, $finding.Reason)
        Write-Output ("    Code: {0}" -f $finding.Text)
    }
}

Write-Output 'PowerShell script-risk scan complete.'
