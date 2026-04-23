<#
.SYNOPSIS
    ABCo Master Tool (CONSOLIDATED & SILENT-CAPABLE)
    1. Launch OneNote: checks if OneNote is running; starts and waits 15s if not.
    2. Sync/Import: finds files in Staging, locates Notebook + Section via COM.
    3. MASS CLEANUP: Scans the entire section and deletes existing duplicates and "Untitled" pages.
    4. Duplicate Check: verifies if a page with the same name already exists (Case-Insensitive).
    5. Robust Import: Title uses CDATA; file paths are XML-escaped.
    6. Heavy-Duty Cleanup: Removes "Untitled page" items with crash protection 
       and dynamic throttling for network drive (V:) stability.
#>

[CmdletBinding()]
param(
    [string]$NotebookName      = "IT Master Documentation",
    [string]$TargetSectionName = "ABCO Documentation",
    [string]$StagingPath       = "V:\ABCo Systems Documentation\IT Master Documentation\ABCo_OneNote_Staging",
    [string]$MasterNotebook    = "V:\ABCo Systems Documentation\IT Master Documentation\Open Notebook.onetoc2",
    [string]$LogFolder         = "V:\ABCo Systems Documentation\Logs",

    # Force silent mode (EXE auto-enables this)
    [switch]$Silent,

    # Optional: only for interactive use
    [switch]$OpenNotebook
)

$ErrorActionPreference = "Stop"

# -------------------------
# Detect compiled EXE host
# -------------------------
$IsCompiled = $false
try {
    if ($MyInvocation.MyCommand.Path -and ($MyInvocation.MyCommand.Path -notlike "*.ps1")) { $IsCompiled = $true }
    if (-not $MyInvocation.MyCommand.Path) { $IsCompiled = $true }
} catch {}

if ($IsCompiled) { $Silent = $true }
if ($Silent) { $OpenNotebook = $false }

# -------------------------
# Logging setup
# -------------------------
if (!(Test-Path $LogFolder)) { New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null }

$ArchiveFolder = Join-Path $LogFolder ("Staging_Archive_{0}" -f (Get-Date -Format 'yyyyMMdd_HHmm'))
$LogPath       = Join-Path $LogFolder ("Import_Summary_{0}.txt" -f (Get-Date -Format 'yyyyMMdd_HHmm'))
$script:Log    = New-Object System.Collections.Generic.List[string]

function Add-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO","WARN","ERROR")]$Level = "INFO"
    )
    $ts = (Get-Date).ToString("HH:mm:ss")
    $line = "[$ts][$Level] $Message"
    $script:Log.Add($line) | Out-Null
    if (-not $Silent) { Write-Output $line }
}

function Escape-XmlAttr([string]$s) {
    if ($null -eq $s) { return "" }
    return [System.Security.SecurityElement]::Escape($s)
}

# -------------------------
# NEW: MASS DUPLICATE REMOVER
# -------------------------
function Invoke-MassDuplicateCleanup {
    param($ON, $SectionId)
    try {
        Add-Log "Starting Mass Duplicate Scan..."
        $ON.SyncHierarchy($SectionId)
        
        [xml]$pagesXml = ""; $ON.GetHierarchy($SectionId, 4, [ref]$pagesXml)
        $ns = New-Object System.Xml.XmlNamespaceManager($pagesXml.NameTable)
        $ns.AddNamespace("one", $pagesXml.DocumentElement.NamespaceURI)
        $allPages = $pagesXml.SelectNodes("//one:Page", $ns)

        $seenPages = @{} 
        $deleteCount = 0

        foreach ($page in $allPages) {
            $name = $page.name.ToLower().Trim()
            
            # Delete if it's a blank "Untitled page" OR if we've already kept a page with this name
            if ($name -eq "untitled page" -or $seenPages.ContainsKey($name)) {
                try {
                    $ON.DeleteHierarchy($page.ID)
                    $deleteCount++
                    Start-Sleep -Milliseconds 300 # Prevent V: drive lockup
                } catch {}
            } else {
                $seenPages.Add($name, $true)
            }
        }
        if ($deleteCount -gt 0) { Add-Log "MASS CLEANUP: Removed $deleteCount existing duplicates/untitled pages." "WARN" }
    } catch {
        Add-Log "Mass cleanup encountered an error, skipping to import." "WARN"
    }
}

# -------------------------
# HARDENED DUPLICATE CHECK (For new imports)
# -------------------------
function Test-PageExists {
    param($ON, $SectionId, $Title)
    try {
        $ON.SyncHierarchy($SectionId)
        Start-Sleep -Seconds 1 
        [xml]$pagesXml = ""; $ON.GetHierarchy($SectionId, 4, [ref]$pagesXml)
        $ns = New-Object System.Xml.XmlNamespaceManager($pagesXml.NameTable)
        $ns.AddNamespace("one", $pagesXml.DocumentElement.NamespaceURI)
        
        $allPages = $pagesXml.SelectNodes("//one:Page", $ns)
        foreach ($page in $allPages) {
            if ($page.name.ToLower().Trim() -eq $Title.ToLower().Trim()) { return $true }
        }
        return $false
    } catch { return $false }
}

# -------------------------
# MAIN LOGIC
# -------------------------
try {
    if (!(Test-Path $ArchiveFolder)) { New-Item -ItemType Directory -Path $ArchiveFolder -Force | Out-Null }

    # 1. Launch OneNote
    if (-not (Get-Process "ONENOTE" -ErrorAction SilentlyContinue)) {
        Add-Log "OneNote is not running. Launching..."
        Start-Process "onenote.exe" -WindowStyle Minimized
        Start-Sleep -Seconds 15 
    }

    # 2. Connect to API
    $ON = New-Object -ComObject OneNote.Application

    # 3. Resolve Hierarchy
    [xml]$xmlStr = ""; $ON.GetHierarchy("", 2, [ref]$xmlStr)
    $ns1 = New-Object System.Xml.XmlNamespaceManager($xmlStr.NameTable)
    $ns1.AddNamespace("one", $xmlStr.DocumentElement.NamespaceURI)
    $schema = $xmlStr.DocumentElement.NamespaceURI

    $nb = $xmlStr.SelectSingleNode("//one:Notebook[@name='$NotebookName']", $ns1)
    if ($null -eq $nb) { throw "Notebook '$NotebookName' not found." }

    Add-Log "VERIFIED: Targeting Notebook at path: $($nb.path)"

    [xml]$secXml = ""; $ON.GetHierarchy($nb.ID, 1, [ref]$secXml)
    $ns2 = New-Object System.Xml.XmlNamespaceManager($secXml.NameTable)
    $ns2.AddNamespace("one", $secXml.DocumentElement.NamespaceURI)

    $sec = $secXml.SelectSingleNode("//one:Section[@name='$TargetSectionName']", $ns2)
    if ($null -eq $sec) { throw "Section '$TargetSectionName' not found." }
    $sectionId = $sec.ID

    # 4. RUN MASS CLEANUP FIRST (Fixes the existing mess)
    Invoke-MassDuplicateCleanup -ON $ON -SectionId $sectionId

    # 5. Import Staging Files
    $files = Get-ChildItem -Path $StagingPath -File -Recurse -ErrorAction SilentlyContinue
    Add-Log "Found $($files.Count) file(s) in staging."

    foreach ($f in $files) {
        $cleanTitle = $f.BaseName
        
        if (Test-PageExists -ON $ON -SectionId $sectionId -Title $cleanTitle) {
            Add-Log "SKIPPED: '$cleanTitle' already exists. Deleting staging file." "WARN"
            Remove-Item -LiteralPath $f.FullName -Force
            continue
        }

        $newPageId = ""
        try {
            $ON.CreateNewPage($sectionId, [ref]$newPageId)
            $pathAttr = Escape-XmlAttr $f.FullName
            $nameAttr = Escape-XmlAttr $f.Name

            $pageXml = @"
<?xml version='1.0'?>
<one:Page xmlns:one='$schema' ID='$newPageId'>
  <one:Title><one:OE><one:T><![CDATA[$cleanTitle]]></one:T></one:OE></one:Title>
  <one:Outline><one:OEChildren><one:OE>
    <one:InsertedFile pathSource='$pathAttr' preferredName='$nameAttr' />
  </one:OE></one:OEChildren></one:Outline>
</one:Page>
"@
            $ON.UpdatePageContent($pageXml)
            
            Start-Sleep -Seconds 2
            Remove-Item -LiteralPath $f.FullName -Force
            Add-Log "SUCCESS: Imported '$($f.Name)'."
        } catch {
            Add-Log "FAILED: '$($f.Name)'. Archiving." "ERROR"
            if ($newPageId) { try { $ON.DeleteHierarchy($newPageId) } catch {} }
            Move-Item -LiteralPath $f.FullName -Destination $ArchiveFolder -Force -ErrorAction SilentlyContinue
        }
    }

    # 6. FINAL CLEANUP (Purge empty folders)
    $subfolders = Get-ChildItem -Path $StagingPath -Directory -Recurse -ErrorAction SilentlyContinue | Sort-Object FullName -Descending
    foreach ($dir in $subfolders) {
        try { Remove-Item -LiteralPath $dir.FullName -Force -Recurse -ErrorAction Stop } catch { }
    }

    if ($OpenNotebook -and (Test-Path $MasterNotebook)) { Invoke-Item $MasterNotebook }

} catch {
    Add-Log "FATAL ERROR: $($_.Exception.Message)" "ERROR"
} finally {
    $script:Log | Out-File -FilePath $LogPath -Encoding UTF8 -Force
}