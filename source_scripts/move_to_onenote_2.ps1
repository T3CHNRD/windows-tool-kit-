# --- CONFIGURATION ---
$LocalStaging = "C:\ABCo_OneNote_Staging"
$NotebookName = "IT Master Documentation" 
$TargetSectionName = "ABCO Documentation"

# 1. FIND THE LATEST LOG FILE
$LogFolder = $PSScriptRoot
$LatestLog = Get-ChildItem -Path $LogFolder -Filter "Import_Log_*.txt" | Sort-Object LastWriteTime -Descending | Select-Object -First 1

if (-not $LatestLog) {
    Write-Host "No log file found in $LogFolder. Please ensure the log is in the same folder as this script." -ForegroundColor Red
    exit
}

Write-Host "Analyzing Log: $($LatestLog.Name)" -ForegroundColor Cyan

# 2. EXTRACT UNIQUE FAILED FILENAMES
$FailedFiles = Get-Content $LatestLog.FullName | Where-Object { $_ -match "FAILED after 3 retries - (.*)" } | ForEach-Object {
    $matches[1].Trim()
} | Select-Object -Unique

if ($FailedFiles.Count -eq 0) {
    Write-Host "No failed files found in the log! Everything is already synced." -ForegroundColor Green
    exit
}

Write-Host "Found $($FailedFiles.Count) unique files to recover." -ForegroundColor Yellow

# 3. CONNECT TO ONENOTE
function Connect-OneNote {
    try {
        if (-not (Get-Process "ONENOTE" -ErrorAction SilentlyContinue)) {
            Start-Process "onenote.exe"; Start-Sleep -Seconds 5
        }
        $Global:OneNote = New-Object -ComObject OneNote.Application
        $XML = ""; $Global:OneNote.GetHierarchy("", [Microsoft.Office.Interop.OneNote.HierarchyScope]::hsSections, [ref]$XML)
        $Global:OneNoteXml = [xml]$XML
        $Global:ns = New-Object Xml.XmlNamespaceManager $Global:OneNoteXml.NameTable
        $Global:ns.AddNamespace("one", $Global:OneNoteXml.DocumentElement.NamespaceURI)
        $node = $Global:OneNoteXml.SelectSingleNode("//one:Notebook[@name='$NotebookName']//one:Section[@name='$TargetSectionName']", $Global:ns)
        return $node.ID
    } catch { return $null }
}

$SectionID = Connect-OneNote
if ($null -eq $SectionID) { Write-Host "Connection Failed." -ForegroundColor Red; exit }

# 4. RECOVERY LOOP
foreach ($FileName in $FailedFiles) {
    $File = Get-ChildItem -Path $LocalStaging -Filter $FileName -Recurse | Select-Object -First 1
    
    if (-not $File) {
        Write-Host "  [SKIP] Could not find $FileName in $LocalStaging" -ForegroundColor Gray
        continue
    }

    $FolderName = $File.Directory.Name
    $PageTitle = "[$FolderName] - $($File.BaseName)"
    Write-Host "`nProcessing: $($File.Name)..." -ForegroundColor White
    
    try {
        $PageID = ""; $OneNote.CreateNewPage($SectionID, [ref]$PageID)
        $tempXmlText = ""; $OneNote.GetPageContent($PageID, [ref]$tempXmlText, [Microsoft.Office.Interop.OneNote.PageInfo]::piBasic)
        $nsUri = ([xml]$tempXmlText).DocumentElement.NamespaceURI
        $titleXml = "<one:Title><one:OE><one:T><![CDATA[$PageTitle]]></one:T></one:OE></one:Title>"

        # STEP A: Try Text/Content Injection (For TXT/CSV/MD)
        if ($File.Extension -match "txt|md|csv") {
            try {
                $rawContent = Get-Content $File.FullName -Raw -ErrorAction SilentlyContinue
                $cleanText = [System.Security.SecurityElement]::Escape($rawContent)
                $content = "<one:Outline><one:OEChildren><one:OE><one:T><![CDATA[$cleanText]]></one:T></one:OE></one:OEChildren></one:Outline>"
                $fullXml = "<?xml version='1.0'?><one:Page xmlns:one='$nsUri' ID='$PageID'>$titleXml$content</one:Page>"
                $OneNote.UpdatePageContent($fullXml)
                Write-Host "  [SUCCESS] Injected as searchable text." -ForegroundColor Green
            } catch {
                # Fallback to standard attachment if injection fails
                $content = "<one:Outline><one:OEChildren><one:OE><one:InsertedFile pathSource='$($File.FullName)' preferredName='$($File.Name)'/></one:OE></one:OEChildren></one:Outline>"
                $fullXml = "<?xml version='1.0'?><one:Page xmlns:one='$nsUri' ID='$PageID'>$titleXml$content</one:Page>"
                $OneNote.UpdatePageContent($fullXml)
                Write-Host "  [SUCCESS] Injection failed; attached as file instead." -ForegroundColor Blue
            }
        } 
        else {
            # STEP B: Try Standard Embedding (For PDF/DOCX)
            try {
                $content = "<one:Outline><one:OEChildren><one:OE><one:InsertedFile pathSource='$($File.FullName)' preferredName='$($File.Name)'/></one:OE></one:OEChildren></one:Outline>"
                $fullXml = "<?xml version='1.0'?><one:Page xmlns:one='$nsUri' ID='$PageID'>$titleXml$content</one:Page>"
                $OneNote.UpdatePageContent($fullXml)
                Write-Host "  [SUCCESS] Embedded as file icon." -ForegroundColor Green
            } catch {
                # STEP C: Hyperlink Fallback (For corrupted/blocked AT&T files)
                $FilePath = "file://$($File.FullName)"
                $content = "<one:Outline><one:OEChildren><one:OE><one:T><![CDATA[<a href='$FilePath'>Click to open: $($File.Name)</a>]]></one:T></one:OE></one:OEChildren></one:Outline>"
                $fullXml = "<?xml version='1.0'?><one:Page xmlns:one='$nsUri' ID='$PageID'>$titleXml$content</one:Page>"
                $OneNote.UpdatePageContent($fullXml)
                Write-Host "  [SUCCESS] Embedding failed; created local hyperlink." -ForegroundColor Yellow
            }
        }
        Start-Sleep -Seconds 2 # Keep the COM interface stable
    } catch {
        Write-Host "  [CRITICAL ERROR] $($File.Name) - $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Final Recovery Attempt Complete." -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan