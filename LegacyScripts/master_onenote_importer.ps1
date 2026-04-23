<# -------------------------------------------------------------------------
ABCO OneNote Documentation Rebuild (Excel TOC version) - SAFE CLEAR + PROGRESS

Key behavior:
- DOES NOT delete/move any notebook folders on disk for clearing.
- Clears ONLY by deleting sections/pages INSIDE the target notebook via OneNote COM.

This version also:
- $ClearOnlyOurSectionName = $true (only clears section named $SectionName)
- VOIP recovery helper (non-destructive report; optional copy restore)
- Moves any leftover *-ARCHIVE-* folders under V:\ABCo Systems Documentation
  to the local Documents folder (requested)

Requires:
- Windows PowerShell 5.1
- OneNote desktop (COM: OneNote.Application)
- Microsoft Excel + Word desktop installed
---------------------------------------------------------------------------#>

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# --------------------------- CONFIG ----------------------------------------
$TargetNotebookPath = "V:\ABCo Systems Documentation\IT Master Documentation"
$TocPath            = "V:\ABCo Systems Documentation\TOC Calculations.xlsx"

$AdditionalRoots = @(
  "P:\abis List of Changes",
  "P:\ABIS Releases",
  "P:\BOS - ALBL Business Operating System",
  "P:\Buildings",
  "P:\IT Tracking - Requests_Projects"
)

# Section we will (re)create and import into.
$SectionName  = "Imported Documentation"

# Safety switches
$ClearOnlyOurSectionName = $true     # ONLY delete the section named $SectionName (if it exists)
$MaxFilesPerFolder       = 0
$EnablePdfPrintToOneNote = $false

# VOIP recovery helper switches
$EnableVoipRestoreCopy   = $false    # set to $true to COPY discovered VOIP notebooks to a RESTORE folder (non-overwriting)
$VoipNotebookKeyword     = "VOIP"    # folder name match
$RootDocsPath            = "V:\ABCo Systems Documentation"  # where archives likely live

# Move any leftover *-ARCHIVE-* folders off V: to local Documents (requested)
$MoveArchivesToLocalDocuments = $true
$ArchiveDestinationRoot = Join-Path ([Environment]::GetFolderPath("MyDocuments")) "ABCO-OneNote-Archives"

$LogPath = Join-Path $env:TEMP ("ABCO-OneNote-Rebuild-{0:yyyyMMdd-HHmmss}.log" -f (Get-Date))
$VoipRecoveryReportPath = Join-Path $env:TEMP ("ABCO-VOIP-Recovery-Report-{0:yyyyMMdd-HHmmss}.txt" -f (Get-Date))

# --------------------------- LOGGING ---------------------------------------
function Write-Log {
  param([string]$Message, [ValidateSet("INFO","WARN","ERROR")] [string]$Level="INFO")
  $line = "[{0:yyyy-MM-dd HH:mm:ss}] [{1}] {2}" -f (Get-Date), $Level, $Message
  $line | Tee-Object -FilePath $LogPath -Append | Out-Host
}

# --------------------------- PREREQS ---------------------------------------
function Assert-Prereqs {
  Write-Log "Checking prerequisites..."
  if ($PSVersionTable.PSVersion.Major -lt 5) { throw "PowerShell 5.1+ required." }

  try { $null = New-Object -ComObject OneNote.Application }
  catch { throw "OneNote desktop COM automation unavailable. Install OneNote desktop (Win32) and try again." }

  try { $null = New-Object -ComObject Excel.Application }
  catch { throw "Microsoft Excel COM automation unavailable. Excel desktop is required for XLSX import + link discovery." }

  try { $null = New-Object -ComObject Word.Application }
  catch { Write-Log "Microsoft Word COM not available. DOC/DOCX/RTF conversion will fail." "WARN" }

  Write-Log "Prereqs OK."
}

# --------------------------- ONENOTE HELPERS --------------------------------
function New-OneNoteApp { New-Object -ComObject OneNote.Application }

function Open-OneNoteNotebook {
  param([object]$OneNoteApp, [string]$NotebookPath)
  try {
    $null = $OneNoteApp.OpenHierarchy($NotebookPath, "", [ref]([string]$null), 0)
    Write-Log "Requested OneNote open notebook hierarchy: $NotebookPath"
  } catch {
    Write-Log "OpenHierarchy warning: $($_.Exception.Message)" "WARN"
  }
}

function Load-OneNoteHierarchyXmlDocument {
  param([object]$OneNoteApp, [string]$StartNodeId, [int]$Scope)
  $xmlString = ""
  $OneNoteApp.GetHierarchy($StartNodeId, $Scope, [ref]$xmlString)

  $doc = New-Object System.Xml.XmlDocument
  $doc.LoadXml($xmlString)

  $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
  $ns.AddNamespace("one", "http://schemas.microsoft.com/office/onenote/2013/onenote")

  return @{ Doc = $doc; Ns = $ns }
}

function Find-NotebookByPath {
  param([object]$OneNoteApp, [string]$NotebookPath)

  $h = Load-OneNoteHierarchyXmlDocument -OneNoteApp $OneNoteApp -StartNodeId "" -Scope 0 # hsNotebooks
  $doc = $h.Doc
  $ns  = $h.Ns

  $normalizedTarget = (Resolve-Path -LiteralPath $NotebookPath).Path.TrimEnd('\')

  $notebooks = $doc.SelectNodes("//one:Notebook", $ns)
  foreach ($nb in $notebooks) {
    $nbPath = $nb.GetAttribute("path")
    if ([string]::IsNullOrWhiteSpace($nbPath)) { continue }

    try { $nbPathResolved = (Resolve-Path -LiteralPath $nbPath).Path.TrimEnd('\') }
    catch { continue }

    if ($nbPathResolved -ieq $normalizedTarget) {
      return @{
        Id   = $nb.GetAttribute("ID")
        Name = $nb.GetAttribute("name")
        Path = $nbPathResolved
      }
    }
  }

  $visible = $notebooks | ForEach-Object { "{0} ({1})" -f $_.GetAttribute("name"), $_.GetAttribute("path") }
  throw "Target notebook path not found in OneNote COM hierarchy. Make sure it is OPEN in OneNote desktop. Visible notebooks: $($visible -join ' | ')"
}

function Ensure-OneNoteSection {
  param([object]$OneNoteApp, [string]$NotebookId, [string]$SectionName)

  $h = Load-OneNoteHierarchyXmlDocument -OneNoteApp $OneNoteApp -StartNodeId $NotebookId -Scope 2 # hsSections
  $doc = $h.Doc
  $ns  = $h.Ns

  $secNode = $doc.SelectSingleNode("//one:Section[@name=`"$SectionName`"]", $ns)
  if ($secNode) {
    $secId = $secNode.GetAttribute("ID")
    if (-not [string]::IsNullOrWhiteSpace($secId)) { return $secId }
  }

  $newSectionId = ""
  $OneNoteApp.CreateNewSection($NotebookId, $SectionName, [ref]$newSectionId)

  if ([string]::IsNullOrWhiteSpace($newSectionId)) {
    throw "CreateNewSection returned an empty section ID."
  }

  Write-Log "Created section '$SectionName' id=$newSectionId"
  return $newSectionId
}

function Clear-OneNoteNotebookContents {
  param(
    [object]$OneNoteApp,
    [string]$NotebookId,
    [string]$NotebookPath,
    [string]$OnlySectionName,
    [bool]$OnlyThatSection
  )

  $resolvedTarget = (Resolve-Path -LiteralPath $NotebookPath).Path.TrimEnd('\')
  $nbInfo = Find-NotebookByPath -OneNoteApp $OneNoteApp -NotebookPath $NotebookPath

  if ($nbInfo.Path -ine $resolvedTarget -or $nbInfo.Id -ine $NotebookId) {
    throw "Safety check failed: notebook mismatch. Refusing to clear."
  }

  Write-Log "Clearing contents INSIDE notebook '$($nbInfo.Name)' at '$($nbInfo.Path)'"
  if ($OnlyThatSection) {
    Write-Log "Safety mode: ONLY deleting section named '$OnlySectionName' (if it exists)." "WARN"
  } else {
    Write-Log "WARNING: deleting ALL sections in the target notebook." "WARN"
  }

  $h = Load-OneNoteHierarchyXmlDocument -OneNoteApp $OneNoteApp -StartNodeId $NotebookId -Scope 2 # hsSections
  $doc = $h.Doc
  $ns  = $h.Ns

  $sections = $doc.SelectNodes("//one:Section", $ns)
  foreach ($sec in $sections) {
    $secName = $sec.GetAttribute("name")
    $secId   = $sec.GetAttribute("ID")
    if ([string]::IsNullOrWhiteSpace($secId)) { continue }

    if ($OnlyThatSection -and ($secName -ne $OnlySectionName)) { continue }

    Write-Log "Deleting section (and its pages): $secName"
    $OneNoteApp.DeleteHierarchy($secId, 0)
  }
}

function New-OneNotePage {
  param([object]$OneNoteApp, [string]$SectionId)
  $pageId = ""
  $OneNoteApp.CreateNewPage($SectionId, [ref]$pageId, 0)
  return $pageId
}

function Update-OneNotePageContentXml {
  param([object]$OneNoteApp, [xml]$PageXml)
  $OneNoteApp.UpdatePageContent($PageXml.OuterXml, 0)
}

function New-PageXmlWithHtmlBody {
  param([string]$PageId, [string]$Title, [string]$HtmlBody)

  $oneNs = "http://schemas.microsoft.com/office/onenote/2013/onenote"
  $xml = New-Object System.Xml.XmlDocument
  $xml.PreserveWhitespace = $true

  $page = $xml.CreateElement("one", "Page", $oneNs)
  $null = $page.SetAttribute("ID", $PageId)
  $xml.AppendChild($page) | Out-Null

  $title = $xml.CreateElement("one", "Title", $oneNs)
  $tOE = $xml.CreateElement("one", "OE", $oneNs)
  $tT  = $xml.CreateElement("one", "T", $oneNs)
  $tT.InnerText = $Title
  $tOE.AppendChild($tT) | Out-Null
  $title.AppendChild($tOE) | Out-Null
  $page.AppendChild($title) | Out-Null

  $outline = $xml.CreateElement("one", "Outline", $oneNs)
  $oeChildren = $xml.CreateElement("one", "OEChildren", $oneNs)
  $oe = $xml.CreateElement("one", "OE", $oneNs)
  $t = $xml.CreateElement("one", "T", $oneNs)

  $cdata = $xml.CreateCDataSection($HtmlBody)
  $t.AppendChild($cdata) | Out-Null

  $oe.AppendChild($t) | Out-Null
  $oeChildren.AppendChild($oe) | Out-Null
  $outline.AppendChild($oeChildren) | Out-Null
  $page.AppendChild($outline) | Out-Null

  return $xml
}

function New-PageXmlWithFileLink {
  param([string]$PageId, [string]$Title, [string]$FilePath)

  $escaped = [System.Web.HttpUtility]::HtmlEncode($FilePath)
  $uri = ("file:///" + ($FilePath -replace "\\","/"))
  $html = "<p><b>Source file:</b> <a href='$uri'>$escaped</a></p>"

  return New-PageXmlWithHtmlBody -PageId $PageId -Title $Title -HtmlBody $html
}

# --------------------------- CONVERSION ------------------------------------
function Convert-WordToFilteredHtml {
  param([string]$Path)

  $word = $null
  $doc  = $null
  $tempDir = Join-Path $env:TEMP ("ABCO-OneNote-WordHTML-{0:yyyyMMdd-HHmmss}-{1}" -f (Get-Date), (Get-Random))
  New-Item -ItemType Directory -Path $tempDir | Out-Null
  $outHtml = Join-Path $tempDir ([IO.Path]::GetFileNameWithoutExtension($Path) + ".html")

  try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $doc = $word.Documents.Open($Path, $false, $true)
    $doc.SaveAs([ref]$outHtml, [ref]10)
    $doc.Close($false)
    $word.Quit()
    return Get-Content -LiteralPath $outHtml -Raw
  } finally {
    if ($doc)  { try { $doc.Close($false) } catch {} }
    if ($word) { try { $word.Quit() } catch {} }
  }
}

function Convert-ExcelToHtml {
  param([string]$Path)

  $excel = $null
  $wb = $null
  $tempDir = Join-Path $env:TEMP ("ABCO-OneNote-ExcelHTML-{0:yyyyMMdd-HHmmss}-{1}" -f (Get-Date), (Get-Random))
  New-Item -ItemType Directory -Path $tempDir | Out-Null
  $outHtml = Join-Path $tempDir ([IO.Path]::GetFileNameWithoutExtension($Path) + ".html")

  try {
    $excel = New-Object -ComObject Excel.Application
    $excel.DisplayAlerts = $false
    $excel.Visible = $false

    $wb = $excel.Workbooks.Open($Path, $null, $true)
    $wb.SaveAs($outHtml, 44)
    $wb.Close($false)
    $excel.Quit()

    return Get-Content -LiteralPath $outHtml -Raw
  } finally {
    if ($wb)    { try { $wb.Close($false) } catch {} }
    if ($excel) { try { $excel.Quit() } catch {} }
  }
}

function Convert-TextToHtml {
  param([string]$Path)
  $txt = Get-Content -LiteralPath $Path -Raw
  $escaped = [System.Web.HttpUtility]::HtmlEncode($txt)
  return "<pre>$escaped</pre>"
}

function Convert-FileToHtml {
  param([string]$Path)

  $ext = ([IO.Path]::GetExtension($Path)).ToLowerInvariant()
  switch ($ext) {
    ".doc"  { return Convert-WordToFilteredHtml -Path $Path }
    ".docx" { return Convert-WordToFilteredHtml -Path $Path }
    ".rtf"  { return Convert-WordToFilteredHtml -Path $Path }
    ".xls"  { return Convert-ExcelToHtml -Path $Path }
    ".xlsx" { return Convert-ExcelToHtml -Path $Path }
    ".txt"  { return Convert-TextToHtml -Path $Path }
    ".md"   { return Convert-TextToHtml -Path $Path }
    ".html" { return Get-Content -LiteralPath $Path -Raw }
    ".htm"  { return Get-Content -LiteralPath $Path -Raw }
    ".pdf"  { throw "PDF is not imported as editable content; handled by PDF fallback." }
    default { throw "Unsupported extension for editable import: $ext" }
  }
}

# --------------------------- LINK DISCOVERY (EXCEL) -------------------------
function Get-LinkedDocumentsFromExcel {
  param([string]$ExcelPath)

  $excel = $null
  $wb = $null
  $found = New-Object System.Collections.Generic.HashSet[string]

  try {
    $excel = New-Object -ComObject Excel.Application
    $excel.DisplayAlerts = $false
    $excel.Visible = $false

    $wb = $excel.Workbooks.Open($ExcelPath, $null, $true)

    foreach ($ws in $wb.Worksheets) {
      foreach ($hl in $ws.Hyperlinks) {
        $addr = $hl.Address
        if (-not [string]::IsNullOrWhiteSpace($addr)) {
          if ($addr -like "file://*") {
            try { $addr = ([uri]$addr).LocalPath } catch {}
          }
          [void]$found.Add($addr)
        }
      }

      $used = $ws.UsedRange
      if ($used -and $used.Value2) {
        $text = ($used.Text | Out-String)
        $matches = [regex]::Matches($text, '(?i)\b[A-Z]:\\[^:\r\n\t"]+\.(docx?|xlsx?|xls|rtf|txt|md|html?|pdf)\b')
        foreach ($m in $matches) { [void]$found.Add($m.Value) }
      }
    }

    return $found.ToArray() | Sort-Object
  } finally {
    if ($wb)    { try { $wb.Close($false) } catch {} }
    if ($excel) { try { $excel.Quit() } catch {} }
  }
}

# --------------------------- PROGRESS HELPERS -------------------------------
$script:OverallTotal = 0
$script:OverallDone  = 0

function Set-OverallTotal {
  param([int]$Total)
  $script:OverallTotal = [Math]::Max($Total, 1)
  $script:OverallDone = 0
}

function Step-OverallProgress {
  param([string]$Activity, [string]$Status, [string]$CurrentOperation = "")
  $script:OverallDone++
  $pct = [int](($script:OverallDone / $script:OverallTotal) * 100)
  Write-Progress -Id 1 -Activity $Activity -Status $Status -PercentComplete $pct -CurrentOperation $CurrentOperation
}

function Complete-ProgressBars {
  Write-Progress -Id 2 -Activity "Folder import" -Completed
  Write-Progress -Id 1 -Activity "Overall" -Completed
}

# --------------------------- IMPORT PIPELINE --------------------------------
function Import-FileAsEditableOneNotePage {
  param([object]$OneNoteApp, [string]$SectionId, [string]$FilePath, [string]$PageTitlePrefix="")

  $title = if ($PageTitlePrefix) { "$PageTitlePrefix - $(Split-Path $FilePath -Leaf)" } else { (Split-Path $FilePath -Leaf) }

  $html = Convert-FileToHtml -Path $FilePath

  $pageId = New-OneNotePage -OneNoteApp $OneNoteApp -SectionId $SectionId
  $pageXml = New-PageXmlWithHtmlBody -PageId $pageId -Title $title -HtmlBody $html
  Update-OneNotePageContentXml -OneNoteApp $OneNoteApp -PageXml $pageXml

  Write-Log "Imported editable: $FilePath -> pageId=$pageId"
}

function Import-PdfFallback {
  param([object]$OneNoteApp, [string]$SectionId, [string]$PdfPath, [string]$TitlePrefix="PDF")

  $title = "$TitlePrefix - $(Split-Path $PdfPath -Leaf)"

  $pageId = New-OneNotePage -OneNoteApp $OneNoteApp -SectionId $SectionId
  $pageXml = New-PageXmlWithFileLink -PageId $pageId -Title $title -FilePath $PdfPath
  Update-OneNotePageContentXml -OneNoteApp $OneNoteApp -PageXml $pageXml
  Write-Log "PDF linked (non-editable content): $PdfPath -> pageId=$pageId" "WARN"

  if ($EnablePdfPrintToOneNote) {
    try {
      Write-Log "Printing PDF to OneNote (printout): $PdfPath" "WARN"
      Start-Process -FilePath $PdfPath -Verb Print -WindowStyle Hidden
    } catch {
      Write-Log "PDF print-to-OneNote failed: $($_.Exception.Message)" "WARN"
    }
  }
}

function Import-FileAuto {
  param([object]$OneNoteApp, [string]$SectionId, [string]$FilePath, [string]$TitlePrefix="")

  $ext = ([IO.Path]::GetExtension($FilePath)).ToLowerInvariant()

  try {
    if ($ext -eq ".pdf") {
      Import-PdfFallback -OneNoteApp $OneNoteApp -SectionId $SectionId -PdfPath $FilePath -TitlePrefix $TitlePrefix
      return
    }
    Import-FileAsEditableOneNotePage -OneNoteApp $OneNoteApp -SectionId $SectionId -FilePath $FilePath -PageTitlePrefix $TitlePrefix
  }
  catch {
    Write-Log "SKIP: $FilePath :: $($_.Exception.Message)" "WARN"
  }
}

function Import-FolderRecursive {
  param(
    [object]$OneNoteApp,
    [string]$SectionId,
    [string]$RootPath,
    [string]$TitlePrefix,
    [int]$FolderIndex,
    [int]$FolderCount
  )

  if (-not (Test-Path $RootPath)) {
    Write-Log "Source path not found, skipping: $RootPath" "WARN"
    return
  }

  $files = Get-ChildItem -LiteralPath $RootPath -File -Recurse -ErrorAction Stop
  if ($MaxFilesPerFolder -gt 0) { $files = $files | Select-Object -First $MaxFilesPerFolder }

  $total = [Math]::Max($files.Count, 1)
  $done = 0

  Write-Log "Importing folder: $RootPath (files: $($files.Count))"

  foreach ($f in $files) {
    $done++
    $pct = [int](($done / $total) * 100)

    Write-Progress -Id 2 -ParentId 1 `
      -Activity ("Folder import [{0}/{1}]: {2}" -f $FolderIndex, $FolderCount, (Split-Path $RootPath -Leaf)) `
      -Status ("{0}/{1} files" -f $done, $files.Count) `
      -PercentComplete $pct `
      -CurrentOperation $f.FullName

    Import-FileAuto -OneNoteApp $OneNoteApp -SectionId $SectionId -FilePath $f.FullName -TitlePrefix $TitlePrefix
    Step-OverallProgress -Activity "Overall rebuild" -Status ("Imported {0}/{1}" -f $script:OverallDone, $script:OverallTotal) -CurrentOperation $f.FullName
  }
}

function Get-PlannedWorkCount {
  param([string[]]$Roots, [string[]]$TocLinks)

  $count = 0
  $count += 1
  $count += ($TocLinks | Where-Object { Test-Path $_ }).Count

  foreach ($r in $Roots) {
    if (-not (Test-Path $r)) { continue }
    $files = Get-ChildItem -LiteralPath $r -File -Recurse -ErrorAction Stop
    if ($MaxFilesPerFolder -gt 0) { $files = $files | Select-Object -First $MaxFilesPerFolder }
    $count += $files.Count
  }

  return $count
}

# --------------------------- ARCHIVE MOVE HELPER (REQUESTED) ----------------
function Move-ArchiveFoldersToLocalDocuments {
  param(
    [string]$SearchRoot,
    [string]$DestinationRoot
  )

  if (-not $MoveArchivesToLocalDocuments) {
    Write-Log "Archive move disabled by config."
    return
  }

  if (-not (Test-Path $SearchRoot)) {
    Write-Log "Archive search root not found, skipping: $SearchRoot" "WARN"
    return
  }

  New-Item -ItemType Directory -Path $DestinationRoot -Force | Out-Null

  $archives = Get-ChildItem -LiteralPath $SearchRoot -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -like "*-ARCHIVE-*" }

  if (-not $archives -or $archives.Count -eq 0) {
    Write-Log "No *-ARCHIVE-* folders found under: $SearchRoot"
    return
  }

  Write-Log "Found $($archives.Count) archive folder(s). Moving to: $DestinationRoot" "WARN"

  foreach ($a in $archives) {
    $dest = Join-Path $DestinationRoot $a.Name

    # Avoid collisions
    if (Test-Path $dest) {
      $dest = Join-Path $DestinationRoot ("{0}-DUP-{1:yyyyMMdd-HHmmss}" -f $a.Name, (Get-Date))
    }

    Write-Log "Moving archive: $($a.FullName) -> $dest" "WARN"
    Move-Item -LiteralPath $a.FullName -Destination $dest -Force
  }
}

# --------------------------- VOIP RECOVERY HELPERS (NON-DESTRUCTIVE) --------
function Invoke-VoipRecoveryReport {
  <#
    This does NOT delete or overwrite anything.
    It will:
    - Look for prior ARCHIVE folders created by earlier runs under $RootDocsPath
    - Search inside them for folders/files containing $VoipNotebookKeyword
    - Write a report with findings + OneNote backup locations
    - Optionally COPY (not move) found VOIP notebook folders to a RESTORE directory if $EnableVoipRestoreCopy = $true
  #>

  Write-Log "VOIP recovery helper starting (non-destructive)."
  $lines = New-Object System.Collections.Generic.List[string]
  $lines.Add("VOIP Recovery Report - $(Get-Date)")
  $lines.Add("RootDocsPath: $RootDocsPath")
  $lines.Add("Keyword: $VoipNotebookKeyword")
  $lines.Add("")

  $archives = @()
  if (Test-Path $RootDocsPath) {
    $archives = Get-ChildItem -LiteralPath $RootDocsPath -Directory -ErrorAction SilentlyContinue |
      Where-Object { $_.Name -like "*-ARCHIVE-*" }
  }

  $lines.Add("Archive folders found: $($archives.Count)")
  foreach ($a in $archives) { $lines.Add(" - $($a.FullName)") }
  $lines.Add("")

  $hits = @()
  foreach ($a in $archives) {
    try {
      $foundDirs = Get-ChildItem -LiteralPath $a.FullName -Directory -Recurse -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like "*$VoipNotebookKeyword*" }
      $foundOneFiles = Get-ChildItem -LiteralPath $a.FullName -File -Recurse -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension -in ".one",".onetoc2" -and $_.FullName -match [regex]::Escape($VoipNotebookKeyword) }

      foreach ($d in $foundDirs) { $hits += $d.FullName }
      foreach ($f in $foundOneFiles) { $hits += $f.FullName }
    } catch {}
  }

  $hits = $hits | Sort-Object -Unique
  $lines.Add("VOIP-related hits inside archives: $($hits.Count)")
  foreach ($h in $hits) { $lines.Add(" - $h") }
  $lines.Add("")

  $backupCandidates = @(
    Join-Path $env:LOCALAPPDATA "Microsoft\OneNote\16.0\Backup",
    Join-Path $env:LOCALAPPDATA "Microsoft\OneNote\Backup"
  )

  $lines.Add("OneNote backup folder candidates (check these in Explorer):")
  foreach ($b in $backupCandidates) {
    $exists = Test-Path $b
    $lines.Add(" - $b  (exists: $exists)")
  }
  $lines.Add("")
  $lines.Add("OneNote UI recovery: History -> Notebook Recycle Bin (try restoring deleted pages/sections).")
  $lines.Add("")

  if ($EnableVoipRestoreCopy -and $hits.Count -gt 0) {
    $restoreRoot = Join-Path $RootDocsPath ("VOIP-RESTORE-{0:yyyyMMdd-HHmmss}" -f (Get-Date))
    New-Item -ItemType Directory -Path $restoreRoot -Force | Out-Null
    $lines.Add("EnableVoipRestoreCopy is TRUE. Copying candidates to: $restoreRoot")

    foreach ($h in $hits) {
      if (Test-Path $h -PathType Container) {
        $dest = Join-Path $restoreRoot (Split-Path $h -Leaf)
        Write-Log "Copying VOIP candidate folder to restore location: $h -> $dest" "WARN"
        Copy-Item -LiteralPath $h -Destination $dest -Recurse -Force
        $lines.Add("COPIED FOLDER: $h -> $dest")
      } elseif (Test-Path $h -PathType Leaf) {
        $dest = Join-Path $restoreRoot (Split-Path $h -Leaf)
        Write-Log "Copying VOIP candidate file to restore location: $h -> $dest" "WARN"
        Copy-Item -LiteralPath $h -Destination $dest -Force
        $lines.Add("COPIED FILE: $h -> $dest")
      }
    }
    $lines.Add("")
    $lines.Add("Next step after copy: In OneNote desktop, use File -> Open -> Browse and select the restored notebook folder / .onetoc2 if applicable.")
  } else {
    $lines.Add("EnableVoipRestoreCopy is FALSE. No files were copied. (This is the safe default.)")
  }

  $lines | Set-Content -LiteralPath $VoipRecoveryReportPath -Encoding UTF8
  Write-Log "VOIP recovery report written to: $VoipRecoveryReportPath" "WARN"

  try { Start-Process explorer.exe "/select,`"$VoipRecoveryReportPath`"" | Out-Null } catch {}
  foreach ($b in $backupCandidates) {
    if (Test-Path $b) { try { Start-Process explorer.exe "`"$b`"" | Out-Null } catch {} }
  }
}

# --------------------------- MAIN ------------------------------------------
function Invoke-AbcoOneNoteRebuild {
  Assert-Prereqs
  Write-Log "Log file: $LogPath"

  if (-not (Test-Path $TocPath)) { throw "TOC Calculations not found at: $TocPath" }

  $one = New-OneNoteApp

  # NOTE: OneNote COM usually only sees notebooks that have been opened at least once in the OneNote desktop UI.
  Open-OneNoteNotebook -OneNoteApp $one -NotebookPath $TargetNotebookPath

  $nb = Find-NotebookByPath -OneNoteApp $one -NotebookPath $TargetNotebookPath
  Write-Log "Target notebook resolved: Name='$($nb.Name)' Path='$($nb.Path)' ID='$($nb.Id)'"

  # Clear notebook contents safely (ONLY our section, by default)
  Clear-OneNoteNotebookContents -OneNoteApp $one -NotebookId $nb.Id -NotebookPath $TargetNotebookPath -OnlySectionName $SectionName -OnlyThatSection:$ClearOnlyOurSectionName

  # Recreate/ensure our import section exists
  $sectionId = Ensure-OneNoteSection -OneNoteApp $one -NotebookId $nb.Id -SectionName $SectionName

  # Extract TOC links up front so overall progress is accurate
  Write-Log "Extracting TOC links (Excel)..."
  $tocLinks = Get-LinkedDocumentsFromExcel -ExcelPath $TocPath

  $planned = Get-PlannedWorkCount -Roots $AdditionalRoots -TocLinks $tocLinks
  Set-OverallTotal -Total $planned
  Write-Log "Planned work items (for progress): $planned"

  Write-Progress -Id 1 -Activity "Overall rebuild" -Status "Starting..." -PercentComplete 0

  # 1) Import TOC
  Write-Log "Importing TOC Calculations (Excel) as editable content..."
  Write-Progress -Id 2 -ParentId 1 -Activity "Folder import" -Status "TOC import" -PercentComplete 0 -CurrentOperation $TocPath
  Import-FileAuto -OneNoteApp $one -SectionId $sectionId -FilePath $TocPath -TitlePrefix "TOC"
  Step-OverallProgress -Activity "Overall rebuild" -Status ("Imported {0}/{1}" -f $script:OverallDone, $script:OverallTotal) -CurrentOperation $TocPath

  # 2) Import TOC-linked docs
  $existingLinks = $tocLinks | Where-Object { Test-Path $_ }
  $linkTotal = [Math]::Max($existingLinks.Count, 1)
  $linkDone = 0

  Write-Log "Importing TOC-linked documents found: $($existingLinks.Count)"
  foreach ($l in $existingLinks) {
    $linkDone++
    $pct = [int](($linkDone / $linkTotal) * 100)
    Write-Progress -Id 2 -ParentId 1 -Activity "Folder import" -Status ("TOC links {0}/{1}" -f $linkDone, $existingLinks.Count) -PercentComplete $pct -CurrentOperation $l

    Import-FileAuto -OneNoteApp $one -SectionId $sectionId -FilePath $l -TitlePrefix "TOC Link"
    Step-OverallProgress -Activity "Overall rebuild" -Status ("Imported {0}/{1}" -f $script:OverallDone, $script:OverallTotal) -CurrentOperation $l
  }

  # 3) Import additional roots
  $folderCount = $AdditionalRoots.Count
  for ($i = 0; $i -lt $folderCount; $i++) {
    $root = $AdditionalRoots[$i]
    $label = Split-Path $root -Leaf
    Import-FolderRecursive -OneNoteApp $one -SectionId $sectionId -RootPath $root -TitlePrefix $label -FolderIndex ($i+1) -FolderCount $folderCount
  }

  Complete-ProgressBars
  Write-Log "DONE. Review OneNote notebook and import log: $LogPath"

  # Move any leftover archive folders off V: to local Documents (requested)
  Move-ArchiveFoldersToLocalDocuments -SearchRoot $RootDocsPath -DestinationRoot $ArchiveDestinationRoot

  Invoke-VoipRecoveryReport
}

Invoke-AbcoOneNoteRebuild