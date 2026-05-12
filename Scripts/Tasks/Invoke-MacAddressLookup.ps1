Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'
Add-Type -AssemblyName Microsoft.VisualBasic
$mac = [Microsoft.VisualBasic.Interaction]::InputBox('Enter a MAC address:', 'MAC Address Lookup', '20:37:06:12:34:56')
if ([string]::IsNullOrWhiteSpace($mac)) { Write-Output 'MAC lookup cancelled.'; exit 0 }
$oui = ($mac.Trim() -replace '[\.:-]','').ToUpper().Substring(0,[Math]::Min(6,($mac.Trim() -replace '[\.:-]','').Length))
Write-Output "KillerTools MAC Address Lookup inspired OUI lookup for $mac (OUI $oui)"
try {
    $vendor = Invoke-RestMethod -Uri "https://api.macvendors.com/$mac" -TimeoutSec 10 -ErrorAction Stop
    Write-Output "Vendor: $vendor"
} catch {
    Write-Output "Online vendor lookup unavailable: $($_.Exception.Message)"
    Write-Output 'Tip: check the first 6 hex digits against an IEEE OUI database when offline.'
}
exit 0
