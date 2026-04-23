$sbEnabled = $null
$ca2023Present = $null

try {
    $sbEnabled = Confirm-SecureBootUEFI
} catch {
    $sbEnabled = $false
}

try {
    $db = (Get-SecureBootUEFI db).bytes
    $text = [System.Text.Encoding]::ASCII.GetString($db)
    $ca2023Present = $text -match 'Windows UEFI CA 2023'
} catch {
    $ca2023Present = $false
}

[pscustomobject]@{
    ComputerName      = $env:COMPUTERNAME
    SecureBootEnabled  = $sbEnabled
    WindowsUEFICA2023  = $ca2023Present
}