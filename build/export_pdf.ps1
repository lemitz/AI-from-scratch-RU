param([string]$Docx, [string]$Pdf)
# Opens a .docx in Word (read-only), refreshes the table of contents and exports a print-quality PDF
# with bookmarks from headings. Run as a separate process per file: some Word add-ins hang
# ExportAsFixedFormat when several exports share one automation session.
$ErrorActionPreference = 'Stop'

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $doc = $word.Documents.Open($Docx, $false, $true, $false)
    try {
        $null = $doc.ComputeStatistics(2)
        $doc.Repaginate()
        foreach ($toc in $doc.TablesOfContents) { $toc.Update() }
        $doc.Repaginate()
        foreach ($toc in $doc.TablesOfContents) { $toc.UpdatePageNumbers() }
        # 17 = PDF, OptimizeFor 0 = print, Range 0 = whole doc, Item 0 = content,
        # CreateBookmarks 1 = from headings, DocStructureTags = tagged PDF
        $doc.ExportAsFixedFormat($Pdf, 17, $false, 0, 0, 1, 1, 0, $true, $true, 1, $true, $true, $false)
    }
    finally {
        $doc.Close(0)
    }
}
finally {
    $word.Quit()
}
