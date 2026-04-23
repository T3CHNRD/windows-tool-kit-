# --- CONFIGURATION ---
$NotebookName = "IT Master Documentation" 
$TargetSectionName = "ABCO Documentation"

# 1. CONNECT TO ONENOTE
$OneNote = New-Object -ComObject OneNote.Application
$XML = ""; $OneNote.GetHierarchy("", [Microsoft.Office.Interop.OneNote.HierarchyScope]::hsSections, [ref]$XML)
$OneNoteXml = [xml]$XML
$ns = New-Object Xml.XmlNamespaceManager $OneNoteXml.NameTable
$ns.AddNamespace("one", $OneNoteXml.DocumentElement.NamespaceURI)
$Section = $OneNoteXml.SelectSingleNode("//one:Notebook[@name='$NotebookName']//one:Section[@name='$TargetSectionName']", $ns)

if (-not $Section) { Write-Host "Section not found!" -ForegroundColor Red; exit }

# 2. DELETE OLD ATTEMPTS (Keep it clean)
$existingPages = ""; $OneNote.GetHierarchy($Section.ID, [Microsoft.Office.Interop.OneNote.HierarchyScope]::hsPages, [ref]$existingPages)
$pXml = [xml]$existingPages
$oldPages = $pXml.SelectNodes("//one:Page[contains(@name, 'START HERE')]", $ns)
foreach ($old in $oldPages) { $OneNote.DeleteHierarchy($old.ID) }

# 3. CONTENT (Using "!!!" to break the "[" sort bracket)
$PageTitle = "!!! START HERE - Search & Usage Guide"
$BodyText = @"
<one:Outline>
  <one:OEChildren>
    <one:OE><one:T><![CDATA[<span style='font-size:16pt;font-weight:bold;color:#594294'>IT Archive Documentation Manual</span>]]></one:T></one:OE>
    <one:OE><one:T><![CDATA[-----------------------------------------------------------------------]]></one:T></one:OE>
    <one:OE><one:T><![CDATA[<span style='font-size:12pt;font-weight:bold;color:#2E75B6'>🔍 How to Search</span>]]></one:T></one:OE>
    <one:OE><one:T><![CDATA[Use <b>Ctrl+E</b> to search across the entire Notebook. OneNote indexes text <i>inside</i> PDFs and Word docs automatically.]]></one:T></one:OE>
    <one:OE><one:T><![CDATA[]]></one:T></one:OE>
    <one:OE><one:T><![CDATA[<span style='font-size:12pt;font-weight:bold;color:#2E75B6'>📄 File Access</span>]]></one:T></one:OE>
    <one:OE><one:T><![CDATA[• <b>Text:</b> Displayed directly. | • <b>Icons:</b> Double-click to open. | • <b>Links:</b> Local file access.]]></one:T></one:OE>
  </one:OEChildren>
</one:Outline>
"@

# 4. CREATE AND JUMP TO PAGE
try {
    $PageID = ""
    $OneNote.CreateNewPage($Section.ID, [ref]$PageID)
    
    $tempXml = ""; $OneNote.GetPageContent($PageID, [ref]$tempXml, [Microsoft.Office.Interop.OneNote.PageInfo]::piBasic)
    $nsUri = ([xml]$tempXml).DocumentElement.NamespaceURI
    $fullXml = "<?xml version='1.0'?><one:Page xmlns:one='$nsUri' ID='$PageID'><one:Title><one:OE><one:T><![CDATA[$PageTitle]]></one:T></one:OE></one:Title>$BodyText</one:Page>"
    
    $OneNote.UpdatePageContent($fullXml)

    # Force Hierarchy move to index 0
    $SectionXML = ""; $OneNote.GetHierarchy($Section.ID, [Microsoft.Office.Interop.OneNote.HierarchyScope]::hsPages, [ref]$SectionXML)
    $sXml = [xml]$SectionXML
    $newNode = $sXml.SelectSingleNode("//one:Page[@ID='$PageID']", $ns)
    $parent = $newNode.ParentNode
    [void]$parent.RemoveChild($newNode)
    [void]$parent.PrependChild($newNode)
    $OneNote.UpdateHierarchy($sXml.OuterXml)

    # NEW: Navigate the app directly to this page
    $OneNote.NavigateTo($PageID)

    Write-Host "Success! The guide should now be the active page on your screen." -ForegroundColor Green
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}