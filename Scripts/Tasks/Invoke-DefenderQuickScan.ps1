[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Write-Output 'Starting Microsoft Defender quick scan.'
Write-Output 'This uses the built-in Defender PowerShell module and may take several minutes.'

if (-not (Get-Command Start-MpScan -ErrorAction SilentlyContinue)) {
    throw 'Start-MpScan is not available. Microsoft Defender PowerShell cmdlets may be missing or managed by another endpoint security product.'
}

if (Get-Command Get-MpComputerStatus -ErrorAction SilentlyContinue) {
    # FIX: MED-09 - report and skip cleanly when Defender is disabled or replaced by third-party AV.
    $before = Get-MpComputerStatus -ErrorAction SilentlyContinue
    if (-not $before -or -not $before.AntivirusEnabled) {
        Write-Output 'Microsoft Defender Antivirus is not active on this system. Quick scan skipped.'
        exit 0
    }
    Write-Output "Before scan: RealTimeProtectionEnabled=$($before.RealTimeProtectionEnabled), AntivirusSignatureAge=$($before.AntivirusSignatureAge), QuickScanAge=$($before.QuickScanAge)"
}

Write-Output 'Launching Defender quick scan now...'
Start-MpScan -ScanType QuickScan
Write-Output 'Defender quick scan command returned.'

if (Get-Command Get-MpComputerStatus -ErrorAction SilentlyContinue) {
    $after = Get-MpComputerStatus
    Write-Output "After scan: QuickScanStartTime=$($after.QuickScanStartTime), QuickScanEndTime=$($after.QuickScanEndTime), QuickScanAge=$($after.QuickScanAge)"
}

if (Get-Command Get-MpThreatDetection -ErrorAction SilentlyContinue) {
    $detections = @(Get-MpThreatDetection | Sort-Object InitialDetectionTime -Descending | Select-Object -First 10)
    if ($detections.Count -eq 0) {
        Write-Output 'Recent Defender detections: none reported by Get-MpThreatDetection.'
    }
    else {
        Write-Output "Recent Defender detections: $($detections.Count) item(s) shown below."
        foreach ($detection in $detections) {
            Write-Output ("Detection: {0}; ActionSuccess={1}; Resources={2}" -f $detection.ThreatName, $detection.ActionSuccess, (($detection.Resources | Select-Object -First 3) -join ', '))
        }
    }
}

Write-Output 'Microsoft Defender quick scan workflow complete.'
