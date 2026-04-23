# Prompt the user for the Target Computer Name
$ComputerName = Read-Host -Prompt "Enter the remote computer name"

# Check if the user actually entered a name
if (-not [string]::IsNullOrWhiteSpace($ComputerName)) {
    
    # Construct the UNC path for the C$ admin share
    $NetworkPath = "\\$ComputerName\c$"
    
    # Check if the path is reachable before opening
    if (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet) {
        Write-Host "Connecting to $NetworkPath..." -ForegroundColor Cyan
        Invoke-Item $NetworkPath
    }
    else {
        Write-Warning "Could not ping $ComputerName. Ensure the machine is online and you have network access."
    }
}
else {
    Write-Warning "No computer name entered. Exiting."
}