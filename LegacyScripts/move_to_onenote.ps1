# --- CONFIGURATION ---
$LocalStaging = "C:\ABCo_OneNote_Staging"
$NotebookName = "IT Master Documentation" 
$TargetSectionName = "ABCO Documentation"
$LogFile = "$PSScriptRoot\Import_Log_$(Get-Date -Format 'yyyyMMdd_HHmm').txt"

function Connect-OneNote {
    Write-Host "Establishing OneNote COM Session..." -ForegroundColor Cyan
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

"--- OneNote Import Log: $(Get-Date) ---" | Out-File $LogFile
$SectionID = Connect-OneNote
if ($null -eq $SectionID) { Write-Host "Connection Failed."; exit }

# Cleanup - Wiping old pages BUT EXCLUDING the "START HERE" guide
Write-Host "Wiping old documentation pages..." -ForegroundColor Yellow
$PagesXML = ""; $OneNote.GetHierarchy($SectionID, [Microsoft.Office.Interop.OneNote.HierarchyScope]::hsPages, [ref]$PagesXML)
foreach ($Page in ([xml]$PagesXML).SelectNodes("//one:Page", $ns)) { 
    # This check ensures we don't delete the guide you just made
    if ($Page.name -notlike "*START HERE*") {
        try { $OneNote.DeleteHierarchy($Page.ID) } catch {} 
    }
}

$Folders = Get-ChildItem -Path $LocalStaging -Directory
$FileCounter = 0

foreach ($Folder in $Folders) {
    Write-Host "`n>>> Processing Category: $($Folder.Name)" -ForegroundColor Cyan
    $Files = Get-ChildItem -Path $Folder.FullName -File
    
    foreach ($File in $Files) {
        # Added jpg and png to the filter
        if ($File.Extension -notmatch "txt|doc|docx|md|rtf|pdf|xls|xlsx|csv|one|jpg|png|jpeg|vsd|vsdx") { continue }
        
        # REMOVED FOLDER PREFIX: Now just the filename
        $PageTitle = "$($File.BaseName)" 
        
        $Success = $false
        $RetryCount = 0

        while (-not $Success -and $RetryCount -lt 3) {
            try {
                $PageID = ""; $OneNote.CreateNewPage($SectionID, [ref]$PageID)
                $tempXmlText = ""; $OneNote.GetPageContent($PageID, [ref]$tempXmlText, [Microsoft.Office.Interop.OneNote.PageInfo]::piBasic)
                $nsUri = ([xml]$tempXmlText).DocumentElement.NamespaceURI
                $titleXml = "<one:Title><one:OE><one:T><![CDATA[$PageTitle]]></one:T></one:OE></one:Title>"

                # Logic for injecting text vs embedding file
                if ($File.Extension -match "txt|md|csv") {
                    $rawText = [System.Security.SecurityElement]::Escape((Get-Content $File.FullName -Raw -ErrorAction SilentlyContinue))
                    $content = "<one:Outline><one:OEChildren><one:OE><one:T><![CDATA[$rawText]]></one:T></one:OE></one:OEChildren></one:Outline>"
                } else {
                    # For images and docs, we embed the file
                    $content = "<one:Outline><one:OEChildren><one:OE><one:InsertedFile pathSource='$($File.FullName)' preferredName='$($File.Name)'/></one:OE></one:OEChildren></one:Outline>"
                }

                $fullXml = "<?xml version='1.0'?><one:Page xmlns:one='$nsUri' ID='$PageID'>$titleXml$content</one:Page>"
                $OneNote.UpdatePageContent($fullXml)
                
                Write-Host "  [SUCCESS] $($File.Name)" -ForegroundColor Green
                "$(Get-Date): SUCCESS - $($File.Name)" | Out-File $LogFile -Append
                $Success = $true
                
                $FileCounter++
                # Throttling to keep OneNote stable
                if ($FileCounter % 5 -eq 0) { [System.GC]::Collect(); Start-Sleep -Seconds 3 } else { Start-Sleep -Seconds 1 }

            } catch {
                $RetryCount++
                Write-Host "  [RETRYING $RetryCount/3] OneNote Busy for $($File.Name)..." -ForegroundColor Yellow
                Start-Sleep -Seconds 10 
                $SectionID = Connect-OneNote
            }
        }
        if (-not $Success) {
            "$(Get-Date): FAILED after 3 retries - $($File.Name)" | Out-File $LogFile -Append
        }
    }
}

"--- End of Job: $(Get-Date) ---" | Out-File $LogFile -Append
Invoke-Item $LogFile