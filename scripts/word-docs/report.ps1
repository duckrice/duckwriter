# Builds a test document in real Word: headings, lists, line spacing, a table across pages with a repeated
# header row, footnotes, and headers/footers that differ on the first and even pages. Saves it and its PDF.
param([string]$Out)
$ErrorActionPreference = 'Stop'
$w = New-Object -ComObject Word.Application
$w.Visible = $false
try {
  $d = $w.Documents.Add()
  $sel = $w.Selection
  $long = "Measurements were repeated three times at each temperature, and the mean value was used. " * 5
  $n = 0
  function P([string]$text = $long, [string]$style = "Normal") {
    $script:n++
    $sel.Style = $d.Styles.Item($style)
    $sel.TypeText(("P{0:D2} " -f $script:n) + $text)
    $sel.TypeParagraph()
  }
  P "Report" "Title"
  P "Introduction" "Heading 1"
  P; P
  $sel.Style = $d.Styles.Item("Normal")
  $sel.TypeText("P05 A sentence with a footnote.")
  $d.Footnotes.Add($sel.Range, [Type]::Missing, "A footnote at the bottom of the page.") | Out-Null
  $sel.TypeParagraph(); $n = 5
  P "First point in a list. $long" "List Bullet"
  P "Second point in a list." "List Bullet"
  P "Third point in a list." "List Number"
  P "Fourth point in a list." "List Number"
  P "Method" "Heading 1"
  $sel.Style = $d.Styles.Item("Normal")
  $sel.ParagraphFormat.LineSpacingRule = 1   # 1.5 lines
  P; P
  $sel.ParagraphFormat.LineSpacingRule = 0
  $sel.ParagraphFormat.SpaceBefore = 18
  P; P
  $sel.ParagraphFormat.SpaceBefore = 0
  P "Results" "Heading 1"
  $sel.Style = $d.Styles.Item("Normal")
  $table = $d.Tables.Add($sel.Range, 30, 3)
  $table.Borders.Enable = $true
  $table.Rows(1).HeadingFormat = $true
  $table.Cell(1, 1).Range.Text = "Temperature"; $table.Cell(1, 2).Range.Text = "Value"; $table.Cell(1, 3).Range.Text = "Note"
  for ($r = 2; $r -le 30; $r++) { $table.Cell($r, 1).Range.Text = "R$r"; $table.Cell($r, 2).Range.Text = "$($r * 3)"; $table.Cell($r, 3).Range.Text = "Measured twice to be sure." }
  $sel.EndKey(6) | Out-Null
  $n = 21
  P "Discussion" "Heading 1"
  P; P; P; P; P; P
  P "Conclusion" "Heading 2"
  P; P

  $s = $d.Sections(1)
  $s.PageSetup.DifferentFirstPageHeaderFooter = $true
  $s.PageSetup.OddAndEvenPagesHeaderFooter = $true
  $s.Headers(1).Range.Text = "Odd header"
  $s.Headers(3).Range.Text = "Even header"
  $s.Headers(2).Range.Text = "First header"
  $s.Footers(1).PageNumbers.Add(2, $true) | Out-Null

  $d.Repaginate()
  $d.SaveAs2($Out, 16)
  $d.ExportAsFixedFormat(($Out -replace '\.docx$', '.pdf'), 17)
  $d.Close($false)
} finally { $w.Quit() }
