Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'
Add-Type -AssemblyName Microsoft.VisualBasic
$domain = [Microsoft.VisualBasic.Interaction]::InputBox('Enter a domain name to check:', 'Domain Lookup', 'example.com')
if ([string]::IsNullOrWhiteSpace($domain)) { Write-Output 'Domain lookup cancelled.'; exit 0 }
$domain = $domain.Trim().ToLower() -replace '^https?://','' -replace '/.*$','' -replace '^www\.',''
Write-Output "KillerTools Domain Lookup inspired report for $domain"
foreach($type in 'A','AAAA','MX','NS','TXT','SOA'){
    Write-Output "--- DNS $type ---"
    try { Resolve-DnsName -Name $domain -Type $type -ErrorAction Stop | Format-Table -AutoSize | Out-String | Write-Output } catch { Write-Output "No $type result or lookup failed: $($_.Exception.Message)" }
}
Write-Output 'WHOIS/RDAP note: full registrar details may require browser/RDAP access depending on TLD.'
exit 0
