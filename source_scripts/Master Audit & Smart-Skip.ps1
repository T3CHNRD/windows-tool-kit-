# --- CONFIGURATION ---
$LocalStaging = "V:\ABCo_OneNote_Staging"
$NotebookName = "IT Master Documentation" 
$TargetSectionName = "ABCO Documentation"
# Use a safe fallback for LogFile if PSScriptRoot is empty
$LogFolder = "V:\ABCo Systems Documentation\Logs"
if (-not (Test-Path $LogFolder)) { New-Item -Path $LogFolder -ItemType Directory | Out-Null }
$LogFile = Join-Path $LogFolder "Import_Log_$(Get-Date -Format 'yyyyMMdd_HHmm').txt"

function Connect-OneNote {
    Write-Host "Establishing OneNote COM Session..." -ForegroundColor Cyan
    try {
        if (-not (Get-Process "ONENOTE" -ErrorAction SilentlyContinue)) {
            Start-Process "onenote.exe"; Start-Sleep -Seconds 5
        }
        $Global:OneNote = New-Object -ComObject OneNote.Application
        
        # Use Integers instead of Enums: 2=hsNotebooks, 1=hsSections, 0=hsPages
        $XML = ""; $Global:OneNote.GetHierarchy("", 1, [ref]$XML)
        $Global:OneNoteXml = [xml]$XML
        $Global:ns = New-Object Xml.XmlNamespaceManager $Global:OneNoteXml.NameTable
        $Global:ns.AddNamespace("one", $Global:OneNoteXml.DocumentElement.NamespaceURI)
        
        # Robust Selection: Search for notebook then section
        $node = $Global:OneNoteXml.SelectSingleNode("//one:Notebook[@name='$NotebookName']//one:Section[@name='$TargetSectionName']", $Global:ns)
        
        # Fallback if names have slight differences (Contains match)
        if ($null -eq $node) {
             $node = $Global:OneNoteXml.SelectSingleNode("//one:Notebook[contains(@name,'$NotebookName')]//one:Section[contains(@name,'$TargetSectionName')]", $Global:ns)
        }
        
        return $node.ID
    } catch { return $null }
}

# --- INITIALIZATION ---
$StartTime = Get-Date
$SectionID = Connect-OneNote

if ($null -eq $SectionID) { 
    Write-Host "Connection Failed. Ensure the Notebook is open in OneNote Desktop." -ForegroundColor Red
    exit 
}

$ImportedAudit = @() 
$SkippedAudit = @()  
$FailedAudit = @()   
$CleanupStats = @{ Duplicates = 0; Blanks = 0; DeletedFiles = 0 }

# --- PHASE 1: NUCLEAR CLEANUP ---
Write-Host "Scanning OneNote for Cleanup..." -ForegroundColor Yellow
$PagesXML = ""; $Global:OneNote.GetHierarchy($SectionID, 0, [ref]$PagesXML) # 0 = hsPages
$ExistingPages = ([xml]$PagesXML).SelectNodes("//one:Page", $Global:ns)
$OneNoteInventory = @{} 

$cleanupCount = 0
foreach ($Page in $ExistingPages) {
    $cleanupCount++
    $Title = $Page.name
    Write-Progress -Activity "Phase 1: Nuclear Cleanup" -Status "Checking: $Title" -PercentComplete (($cleanupCount / $ExistingPages.Count) * 100)
    
    # Clean Blanks
    if ([string]::IsNullOrWhiteSpace($Title) -or $Title -match "Untitled page") {
        try { $Global:OneNote.DeleteHierarchy($Page.ID); $CleanupStats.Blanks++ } catch {}
        continue
    }
    
    if ($Title -like "*START HERE*") { continue }

    # Clean Duplicates (Improved Regex to handle spaces before underscore)
    if ($Title -match "(\s*_\d+|\s*\(Recovered\))$") {
        try { 
            $Global:OneNote.DeleteHierarchy($Page.ID)
            $CleanupStats.Duplicates++ 
            Write-Host "  [CLEAN] Deleted Duplicate: $Title" -ForegroundColor Gray
        } catch {}
        continue
    }

    # Add to Inventory for Phase 2
    if (-not $OneNoteInventory.ContainsKey($Title)) { $OneNoteInventory.Add($Title, $Page.ID) }
}

# --- PHASE 2: SMART-SKIP IMPORT LOOP ---
if (-not (Test-Path $LocalStaging)) { Write-Host "Staging path not found: $LocalStaging" -ForegroundColor Red; exit }
$FilesToProcess = Get-ChildItem -Path $LocalStaging -File -Recurse | Where-Object { $_.Extension -match "txt|doc|docx|md|rtf|pdf|xls|xlsx|csv|one|jpg|png|jpeg|vsd|vsdx" }
$TotalFiles = $FilesToProcess.Count

$currentFileNum = 0
foreach ($File in $FilesToProcess) {
    $currentFileNum++
    $PageTitle = $File.BaseName
    Write-Progress -Activity "Phase 2: Syncing Files" -Status "Processing: $($File.Name)" -PercentComplete (($currentFileNum / $TotalFiles) * 100)

    if ($OneNoteInventory.ContainsKey($PageTitle)) {
        Write-Host "  [SKIP] '$PageTitle' exists. Clearing staging..." -ForegroundColor Gray
        Remove-Item $File.FullName -Force
        $SkippedAudit += $File.Name
        $CleanupStats.DeletedFiles++
        continue
    }
    
    try {
        $PageID = ""; $Global:OneNote.CreateNewPage($SectionID, [ref]$PageID)
        $tempXmlText = ""; $Global:OneNote.GetPageContent($PageID, [ref]$tempXmlText, 0) # 0 = piBasic
        $nsUri = ([xml]$tempXmlText).DocumentElement.NamespaceURI
        $titleXml = "<one:Title><one:OE><one:T><![CDATA[$PageTitle]]></one:T></one:OE></one:Title>"

        if ($File.Extension -match "txt|md|csv") {
            $rawText = [System.Security.SecurityElement]::Escape((Get-Content $File.FullName -Raw -ErrorAction SilentlyContinue))
            $content = "<one:Outline><one:OEChildren><one:OE><one:T><![CDATA[$rawText]]></one:T></one:OE></one:OEChildren></one:Outline>"
        } else {
            $content = "<one:Outline><one:OEChildren><one:OE><one:InsertedFile pathSource='$($File.FullName)' preferredName='$($File.Name)'/></one:OE></one:OEChildren></one:Outline>"
        }

        $fullXml = "<?xml version='1.0'?><one:Page xmlns:one='$nsUri' ID='$PageID'>$titleXml$content</one:Page>"
        $Global:OneNote.UpdatePageContent($fullXml)
        
        $ImportedAudit += $File.Name
        Write-Host "  [OK] Imported: $($File.Name)" -ForegroundColor Green
        
        Remove-Item $File.FullName -Force
        $CleanupStats.DeletedFiles++

        if ($File.Extension -match "pdf|xls|vsd") { Start-Sleep -Seconds 1 }
    } catch {
        Write-Host "  [ERROR] Failed: $($File.Name)" -ForegroundColor Red
        $FailedAudit += $File.Name
    }
}
Write-Progress -Activity "Syncing Files" -Completed

# --- PHASE 3: FINAL REPORT ---
$Report = @"
===============================================
   ONENOTE SYNC AUDIT
   Run Date: $(Get-Date)
===============================================
Newly Imported:      $($ImportedAudit.Count)
Skipped (Existing):  $($SkippedAudit.Count)
Files Cleared:       $($CleanupStats.DeletedFiles)
Junk Pages Purged:   $($CleanupStats.Blanks + $CleanupStats.Duplicates)
Failed:              $($FailedAudit.Count)
===============================================
"@

$Report | Out-File $LogFile
Write-Host "`nSync Complete. Log: $LogFile" -ForegroundColor Cyan