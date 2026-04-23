<#
.SYNOPSIS
    ABCo Master Import v16 (TOTAL DIRECTORY PURGE)
    - Fixes: Empty subfolders being left behind in staging.
    - Logic: Deletes all files AND subfolders after processing.
    - Verification: Ensures the staging root is a 100% "Clean Slate."
#>

[CmdletBinding()]
param(
    [string]$NotebookName      = "IT Master Documentation",
    [string]$TargetSectionName = "ABCO Documentation",
    [string]$StagingPath       = "V:\ABCo Systems Documentation\IT Master Documentation\ABCo_OneNote_Staging",
    [string]$LogFolder         = "V:\ABCo Systems Documentation\Logs"
)

$ErrorActionPreference = "Stop"
$ArchiveFolder = Join-Path $LogFolder "Staging_Archive_$(Get-Date -Format 'yyyyMMdd_HHmm')"
$LogPath = Join-Path $LogFolder "Import_Summary_$(Get-Date -Format 'yyyyMMdd_HHmm').txt"
$script:Log = New-Object System.Collections.Generic.List[string]

# FILTER: Block system/binary files from import
$ExcludedExtensions = @('.tar','.msi','.exe','.zip','.ps1','.sh','.aes','.bin','.spa','.sig','.rel','.7z','.rar')

function Add-Log {
    param([string]$Message, [ValidateSet("INFO","WARN","ERROR")]$Level = "INFO")
    $ts = (Get-Date).ToString("HH:mm:ss")
    $line = "[$ts][$Level] $Message"
    $script:Log.Add($line) | Out-Null
    $color = if($Level -eq "ERROR"){"Red"}elseif($Level -eq "WARN"){"Yellow"}else{"Gray"}
    Write-Host $line -ForegroundColor $color
}



try {
    if (!(Test-Path $LogFolder)) { New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null }
    if (!(Test-Path $ArchiveFolder)) { New-Item -ItemType Directory -Path $ArchiveFolder -Force | Out-Null }
    
    Add-Log "Starting v16 - Total Staging & Directory Purge..."
    $ON = New-Object -ComObject OneNote.Application
    
    # --- 1. RESOLVE HIERARCHY ---
    [xml]$xmlStr = ""; $ON.GetHierarchy("", 2, [ref]$xmlStr)
    $ns = New-Object System.Xml.XmlNamespaceManager($xmlStr.NameTable)
    $ns.AddNamespace("one", $xmlStr.DocumentElement.NamespaceURI)
    $schema = $xmlStr.DocumentElement.NamespaceURI
    
    $nb = ([System.Xml.XmlElement]$xmlStr.DocumentElement).SelectSingleNode("//one:Notebook[@name='$NotebookName']", $ns)
    $secXml = ""; $ON.GetHierarchy($nb.ID, 1, [ref]$secXml)
    $sec = ([System.Xml.XmlElement]([xml]$secXml).DocumentElement).SelectSingleNode("//one:Section[@name='$TargetSectionName']", $ns)
    $sectionId = $sec.ID

    # --- 2. INDEX & CLEAN BLANKS ---
    [xml]$pgXml = ""; $ON.GetHierarchy($sectionId, 0, [ref]$pgXml)
    $existingTitles = @{}
    foreach($p in $pgXml.SelectNodes("//one:Page", $ns)) {
        if ([string]::IsNullOrWhiteSpace($p.name) -or $p.name -eq "Untitled page") {
            $ON.DeleteHierarchy($p.ID)
        } else { $existingTitles[$p.name.ToLower().Trim()] = $true }
    }

    # --- 3. IMPORT & PURGE FILES ---
    $files = Get-ChildItem -Path $StagingPath -File -Recurse | Where-Object { $_.Extension.ToLower() -notin $ExcludedExtensions }
    Add-Log "Processing $($files.Count) files..."

    $count = 0
    foreach ($f in $files) {
        $count++
        $canonName = $f.BaseName.ToLower().Trim()
        Write-Progress -Activity "ABCo Master Process" -Status "Importing: $($f.Name)" -PercentComplete (($count/$files.Count)*100)

        if ($existingTitles.ContainsKey($canonName)) {
            Add-Log "DUPLICATE: Archiving $($f.Name)." "WARN"
            Move-Item $f.FullName (Join-Path $ArchiveFolder $f.Name) -Force -ErrorAction SilentlyContinue
            continue
        }

        try {
            $newPageId = ""
            $ON.CreateNewPage($sectionId, [ref]$newPageId)
            $cleanTitle = $f.BaseName -replace '[&<>"'']', '' 

            $pageXml = "<?xml version='1.0'?>
            <one:Page xmlns:one='$schema' ID='$newPageId'>
                <one:Title><one:OE><one:T><![CDATA[$cleanTitle]]></one:T></one:OE></one:Title>
                <one:Outline><one:OEChildren><one:OE>
                    <one:InsertedFile pathSource='$($f.FullName)' preferredName='$($f.Name)' />
                </one:OE></one:OEChildren></one:Outline>
            </one:Page>"
            
            $ON.UpdatePageContent($pageXml)
            Start-Sleep -Milliseconds 500
            Remove-Item $f.FullName -Force
            $existingTitles[$canonName] = $true
            Add-Log "SUCCESS: $($f.Name)"
        } catch {
            Add-Log "FAILED: $($f.Name). Archiving." "ERROR"
            if ($newPageId) { try { $ON.DeleteHierarchy($newPageId) } catch {} }
            Start-Sleep -Milliseconds 500
            Move-Item $f.FullName (Join-Path $ArchiveFolder $f.Name) -Force -ErrorAction SilentlyContinue
        }
    }

    # --- 4. NEW: RECURSIVE DIRECTORY PURGE ---
    Add-Log "Cleaning up empty subdirectories in staging..."
    # Get all subdirectories, sorted by depth (deepest first) to ensure clean removal
    $subDirs = Get-ChildItem -Path $StagingPath -Directory -Recurse | Sort-Object { $_.FullName.Length } -Descending
    foreach ($dir in $subDirs) {
        try {
            # Check if directory is empty (or only contains things we want to kill)
            $items = Get-ChildItem -Path $dir.FullName -Recurse
            if ($items.Count -eq 0) {
                Remove-Item $dir.FullName -Force -Recurse
                Add-Log "REMOVED FOLDER: $($dir.Name)"
            } else {
                # Force delete any remaining files (like .DS_Store or hidden logs) then delete folder
                Remove-Item $dir.FullName -Force -Recurse -ErrorAction SilentlyContinue
            }
        } catch {
            Add-Log "Could not remove folder $($dir.Name): $($_.Exception.Message)" "WARN"
        }
    }

    $finalFileCheck = (Get-ChildItem -Path $StagingPath -File -Recurse).Count
    $finalFolderCheck = (Get-ChildItem -Path $StagingPath -Directory -Recurse).Count
    Add-Log "Final Staging State: $finalFileCheck Files, $finalFolderCheck Folders."
    Add-Log "Process Complete. Clean Slate achieved."

} catch {
    Add-Log "FATAL: $($_.Exception.Message)" "ERROR"
} finally {
    Write-Progress -Activity "ABCo Master Process" -Completed
    $script:Log | Out-File -FilePath $LogPath
    Start-Process notepad.exe $LogPath
}