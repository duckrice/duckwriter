# Builds a test document in real Word with page and section formatting (landscape pages, page
# numbers that restart, in roman numerals, headers), saves it, and has Word make its PDF.
param([string]$Out)
$ErrorActionPreference = 'Stop'
$w = New-Object -ComObject Word.Application
$w.Visible = $false
try {
  $d = $w.Documents.Add()
  $sel = $w.Selection
  $long = "Measurements were repeated three times at each temperature, and the mean value was used. " * 6
  $n = 0
  function P([string]$text = $long) { $script:n++; $sel.TypeText(("P{0:D2} " -f $script:n) + $text); $sel.TypeParagraph() }

  # Section 1: portrait, footer "Page n", a title page with no number.
  P "Title page"
  P
  $sel.InsertBreak(7)            # page break
  P; P; P; P; P; P; P
  P "Heading kept with next"
  $sel.Paragraphs(1).Range.Paragraphs(1).Previous().KeepWithNext = $true
  P; P; P
  $sel.InsertBreak(2)            # section break, next page
  # Section 2: landscape, numbered i, ii, iii from i.
  P; P; P; P; P; P
  $sel.InsertBreak(3)            # section break, continuous
  P; P
  $sel.InsertBreak(5)            # section break, odd page
  P; P; P; P
  $p = $d.Paragraphs($d.Paragraphs.Count - 2)
  $p.PageBreakBefore = $true
  P; P; P; P; P; P; P; P

  $s1 = $d.Sections(1)
  $s1.PageSetup.DifferentFirstPageHeaderFooter = $true
  $s1.Footers(1).PageNumbers.Add(1, $true) | Out-Null
  $s2 = $d.Sections(2)
  $s2.PageSetup.Orientation = 1
  $s2.Footers(1).LinkToPrevious = $false
  $s2.Footers(1).PageNumbers.RestartNumberingAtSection = $true
  $s2.Footers(1).PageNumbers.StartingNumber = 1
  $s2.Footers(1).PageNumbers.NumberStyle = 2   # lower roman
  $s2.Headers(1).LinkToPrevious = $false
  $s2.Headers(1).Range.Text = "Appendix"
  $s3 = $d.Sections(3)
  $s3.PageSetup.Orientation = 1
  $s4 = $d.Sections(4)
  $s4.PageSetup.Orientation = 0
  $s4.Footers(1).LinkToPrevious = $false
  $s4.Footers(1).PageNumbers.RestartNumberingAtSection = $true
  $s4.Footers(1).PageNumbers.StartingNumber = 10
  $s4.Footers(1).PageNumbers.NumberStyle = 0

  $d.Repaginate()
  $d.SaveAs2($Out, 16)
  $d.ExportAsFixedFormat(($Out -replace '\.docx$', '.pdf'), 17)
  $d.Close($false)
} finally { $w.Quit() }
