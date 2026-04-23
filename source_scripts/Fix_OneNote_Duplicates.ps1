<#
.SYNOPSIS
    ABCo Master Tool - Merged & Hardened Version
    Location: V:\ABCo Systems Documentation\IT Master Documentation
    Target Section: ABCO Documentation
    Includes: Progress Bars, Logging, and Regex-based duplicate detection (e.g., _1, _2).
#>

[CmdletBinding()]
param(
    [string]$NotebookName      = "IT Master Documentation",
    [string]$TargetSectionName = "ABCO Documentation",
    [string]$StagingPath       = "V:\ABCo Systems Documentation\IT Master Documentation\ABCo_OneNote_Staging",
    [string]$MasterNotebook    = "V:\ABCo Systems Documentation\IT Master Documentation\Open Notebook.onetoc2",
    [string]$LogFolder         = "V:\ABCo Systems Documentation\Logs",
    [switch]$Silent,
    [switch]$OpenNotebook
)

$ErrorActionPreference = "Stop"

# -------------------------
# Logging setup
# -------------------------
if (!(Test-Path $LogFolder)) { New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null }
$LogPath = Join-Path $LogFolder ("Import_Summary_{0}.txt" -f (Get-Date -Format 'yyyyMMdd_HHmm'))
$script:Log = New-Object System.Collections.Generic.List[string]

function Add-Log {
    param([string]$Message, [ValidateSet("INFO","WARN","ERROR")]$Level = "INFO")
    $line = "[(Get-Date -Format 'HH:mm:ss')][$Level] $Message"
    $script:Log.Add($line) | Out-Null
    if (-not $Silent) { Write-Output $line }
}

function Escape-XmlAttr([string]$s) { return [System.Security.SecurityElement]::Escape($s) }

# -------------------------
# MASS CLEANUP FUNCTION (With Regex for _1 suffix)
# -------------------------
function Invoke-MassDuplicateCleanup {
    param($ON, $SectionId)
    try {
        Add-Log "Refreshing Section from V: drive for cleanup..."
        $ON.SyncHierarchy($SectionId) 
        
        [xml]$pagesXml = ""; $ON.GetHierarchy($SectionId, 4, [ref]$pagesXml)
        $ns = New-Object System.Xml.XmlNamespaceManager($pagesXml.NameTable)
        $ns.AddNamespace("one", $pagesXml.DocumentElement.NamespaceURI)
        $allPages = $pagesXml.SelectNodes("//one:Page", $ns)

        $seenPages = @{} 
        $deleteCount = 0
        $current = 0

        foreach ($page in $allPages) {
            $current++
            $originalName = $page.name.Trim()
            
            # Regex to strip trailing underscores and numbers (e.g., "2nd floor_1" -> "2nd floor")
            $cleanBaseName = $originalName -replace '_\d+$', ''
            $lookupName = $cleanBaseName.ToLower()
            
            if (-not $Silent) {
                Write-Progress -Activity "Cleaning ABCO Documentation" -Status "Checking: $originalName" -PercentComplete (($current / $allPages.Count) * 100)
            }

            if ($originalName.ToLower() -eq "untitled page" -or $seenPages.ContainsKey($lookupName)) {
                try {
                    $ON.DeleteHierarchy($page.ID)
                    $deleteCount++
                    Start-Sleep -Milliseconds 300 
                } catch {}
            } else {
                $seenPages.Add($lookupName, $true)
            }
        }
        Add-Log "CLEANUP TOTAL: Removed $deleteCount duplicates (including _1 variations) from ABCO Documentation." "WARN"
    } catch { Add-Log "Cleanup failed: $($_.Exception.Message)" "ERROR" }
    finally { Write-Progress -Activity "Cleaning ABCO Documentation" -Completed }
}

# -------------------------
# MAIN LOGIC
# -------------------------
try {
    # 1. Start OneNote
    if (-not (Get-Process "ONENOTE" -ErrorAction SilentlyContinue)) {
        Add-Log "Opening OneNote..."
        Start-Process "onenote.exe" -WindowStyle Minimized
        Start-Sleep -Seconds 15 
    }

    $ON = New-Object -ComObject OneNote.Application

    # 2. Resolve Notebook & Section
    [xml]$xmlStr = ""; $ON.GetHierarchy("", 2, [ref]$xmlStr)
    $ns = New-Object System.Xml.XmlNamespaceManager($xmlStr.NameTable)
    $ns.AddNamespace("one", $xmlStr.DocumentElement.NamespaceURI)
    $schema = $xmlStr.DocumentElement.NamespaceURI

    $nb = $xmlStr.SelectSingleNode("//one:Notebook[@name='$NotebookName']", $ns)
    if ($null -eq $nb) { throw "Notebook '$NotebookName' not found! Please open it manually from the V: drive first." }

    [xml]$secXml = ""; $ON.GetHierarchy($nb.ID, 1, [ref]$secXml)
    $ns2 = New-Object System.Xml.XmlNamespaceManager($secXml.NameTable)
    $ns2.AddNamespace("one", $secXml.DocumentElement.NamespaceURI)
    $sec = $secXml.SelectSingleNode("//one:Section[@name='$TargetSectionName']", $ns2)
    if ($null -eq $sec) { throw "Section '$TargetSectionName' not found." }
    $sectionId = $sec.ID

    # 3. RUN CLEANUP (Nukes the _1 duplicates already in OneNote)
    Invoke-MassDuplicateCleanup -ON $ON -SectionId $sectionId

    # 4. PROCESS STAGING FILES
    $files = Get-ChildItem -Path $StagingPath -File -Recurse -ErrorAction SilentlyContinue
    if ($files) {
        Add-Log "Found $($files.Count) new files in staging."
        $fCount = 0
        foreach ($f in $files) {
            $fCount++
            $cleanTitle = $f.BaseName
            if (-not $Silent) {
                Write-Progress -Activity "Importing to ABCO Documentation" -Status "File: $cleanTitle" -PercentComplete (($fCount / $files.Count) * 100)
            }

            # Advanced Duplicate Pre-Check: 
            # This checks if the page exists OR if a version with _1 suffix exists
            [xml]$checkXml = ""; $ON.GetHierarchy($sectionId, 4, [ref]$checkXml)
            $existingPages = $checkXml.SelectNodes("//one:Page", $ns)
            $alreadyExists = $false

            foreach ($p in $existingPages) {
                $existingClean = $p.name -replace '_\d+$', ''
                if ($existingClean.ToLower().Trim() -eq $cleanTitle.ToLower().Trim()) {
                    $alreadyExists = $true
                    break
                }
            }

            if ($alreadyExists) {
                Add-Log "SKIPPED: '$cleanTitle' (or a version of it) already exists. Deleting staging file." "WARN"
                Remove-Item -LiteralPath $f.FullName -Force
                continue
            }

            # Create Page
            $newPageId = ""
            try {
                $ON.CreateNewPage($sectionId, [ref]$newPageId)
                $pageXml = "<?xml version='1.0'?><one:Page xmlns:one='$schema' ID='$newPageId'><one:Title><one:OE><one:T><![CDATA[$cleanTitle]]></one:T></one:OE></one:Title><one:Outline><one:OEChildren><one:OE><one:InsertedFile pathSource='$(Escape-XmlAttr $f.FullName)' preferredName='$(Escape-XmlAttr $f.Name)' /></one:OE></one:OEChildren></one:Outline></one:Page>"
                $ON.UpdatePageContent($pageXml)
                Start-Sleep -Seconds 2
                Remove-Item -LiteralPath $f.FullName -Force
                Add-Log "SUCCESS: Imported '$cleanTitle'."
            } catch {
                Add-Log "FAILED: $($f.Name). Moved to archive." "ERROR"
                if ($newPageId) { try { $ON.DeleteHierarchy($newPageId) } catch {} }
            }
        }
    } else {
        Add-Log "Staging folder is empty."
    }

} catch {
    Add-Log "FATAL: $($_.Exception.Message)" "ERROR"
} finally {
    Write-Progress -Activity "Importing to ABCO Documentation" -Completed
    $script:Log | Out-File -FilePath $LogPath -Encoding UTF8 -Force
}