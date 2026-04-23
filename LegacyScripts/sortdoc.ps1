# --- CONFIGURATION ---
$RemoteDir = "V:\ABCo Systems Documentation"
$VoipNotebookDir = "V:\ABCo Systems Documentation\IT Master Documentation\VOIP System"
$LocalStaging = "C:\ABCo_OneNote_Staging"

# Strict whitelist of allowed extensions
$AllowedExtensions = @(
    ".doc", ".docx", ".xls", ".xlsx",  # Office
    ".vsd", ".vsdx",                   # Visio
    ".pdf",                            # Portable Docs
    ".txt", ".md", ".csv",             # Plain Text / Configs
    ".one",                            # OneNote Sections
    ".jpg", ".jpeg", ".png"            # Images/Topology Screenshots
)

# Organizational Map
$Map = @{
    "01_Infrastructure_Electrical" = @("*Panel*", "*Electrical*", "*Breaker*", "*Power*")
    "02_Infrastructure_Wiring"     = @("*Pair*", "*Wire*", "*Cable*", "*Building*")
    "03_Network_Hardware"          = @("*Switch*", "*Cisco*", "*Netgear*", "*Cluster*", "*VSD*", "*Topology*", "*Config*", "*Gateway*")
    "04_Administration_Personnel"  = @("*Contact*", "*User*", "*Welcome*", "*HR*", "*Personnel*", "*Directory*", "*Phone*", "*Extension*")
    "05_Projects_MACs"             = @("*Aegis*", "*MAC*", "*Order*", "*Project*", "*Move*", "*Change*")
    "06_Telecom_and_AT&T"          = @("*AT&T*", "*Circuit*", "*Account*", "*Agreement*", "*Teleconference*", "*VOIP*", "*PRI*", "*SIP*", "*PBX*")
    "07_Misc_IT_Documentation"     = @("*Doc*", "*Misc*", "*Manual*", "*Instruction*", "*Procedure*", "*ReadMe*", "*Notes*")
}

if (!(Test-Path $LocalStaging)) { New-Item -ItemType Directory -Path $LocalStaging }

# 1. SCAN BOTH DIRECTORIES
$PathsToScan = @($RemoteDir, $VoipNotebookDir)
Write-Host "Strict Scan started. Filtering for Docs and Images only..." -ForegroundColor Yellow

# Collect only files matching our allowed extensions
$AllFiles = $PathsToScan | ForEach-Object { 
    Get-ChildItem -Path $_ -File -Recurse -ErrorAction SilentlyContinue | 
    Where-Object { $AllowedExtensions -contains $_.Extension.ToLower() }
}

$TotalFiles = $AllFiles.Count
if ($TotalFiles -eq 0) { Write-Host "No matching documents found!" -ForegroundColor Red; exit }

$CurrentFileNum = 0

# 2. SORT AND COPY LOOP
foreach ($file in $AllFiles) {
    $CurrentFileNum++
    $Percent = ($CurrentFileNum / $TotalFiles) * 100
    Write-Progress -Activity "Sorting Docs & Images" -Status "Processing: $($file.Name)" -PercentComplete $Percent

    $matched = $false
    $targetFolder = ""

    # Keyword Matching
    foreach ($subject in $Map.Keys) {
        foreach ($pattern in $Map[$subject]) {
            if ($file.Name -like $pattern) {
                $targetFolder = $subject
                $matched = $true
                break 
            }
        }
        if ($matched) { break }
    }

    # If it's an allowed extension but no keyword matched, move to review
    if (!$matched) {
        $targetFolder = "_Uncategorized_Review"
    }

    # Execute Copy with Duplicate Handling
    $targetPath = Join-Path $LocalStaging $targetFolder
    if (!(Test-Path $targetPath)) { New-Item -ItemType Directory -Path $targetPath }
    
    $destination = Join-Path $targetPath $file.Name
    
    # Handle filename collisions (e.g., multiple "site_photo.jpg" files)
    if (Test-Path $destination) {
        $count = 1
        while (Test-Path $destination) {
            $newName = "$($file.BaseName)_$count$($file.Extension)"
            $destination = Join-Path $targetPath $newName
            $count++
        }
    }

    Copy-Item $file.FullName $destination -Force
}

Write-Host "`nStrict Sort Complete!" -ForegroundColor Green
Write-Host "Total Documentation/Image files captured: $TotalFiles" -ForegroundColor White