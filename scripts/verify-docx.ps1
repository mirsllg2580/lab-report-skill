# Format + pagination audit of a .docx through Word COM.
#   .\verify-docx.ps1 -Docx D:\out\report.docx [-Out report-check.txt] [-Pdf preview.pdf]
# Reports: real page count (from PDF, not from Word's read-only statistics),
# per-paragraph page/indent/font, table geometry, footer fields.
param(
  [Parameter(Mandatory = $true)][string]$Docx,
  [string]$Out,
  [string]$Pdf
)
$ErrorActionPreference = 'Stop'
if (-not $Out) { $Out = [System.IO.Path]::ChangeExtension($Docx, '.check.txt') }
if (-not $Pdf) { $Pdf = [System.IO.Path]::ChangeExtension($Docx, '.pdf') }

$lines = New-Object System.Collections.Generic.List[string]
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
  $doc = $word.Documents.Open($Docx, $false, $false)
  $doc.Repaginate()
  $lines.Add("WORD_PAGES=" + $doc.ComputeStatistics(2) + "  (unreliable when opened read-only; trust PDF_PAGES)")

  $tables = $doc.Tables.Count
  $lines.Add("TABLES=$tables")
  for ($t = 1; $t -le $tables; $t++) {
    $tb = $doc.Tables.Item($t)
    $lines.Add("TABLE$t rows=" + $tb.Rows.Count)
    for ($r = 1; $r -le $tb.Rows.Count; $r++) {
      $vals = @(); $sumW = 0
      for ($c = 1; $c -le 6; $c++) {
        try {
          $cell = $tb.Cell($r, $c)
          $vals += ("[" + ($cell.Range.Text -replace "[`r`n`a]", '') + "]")
          $sumW += $cell.Width
        } catch { }
      }
      $h = $tb.Rows.Item($r).Height
      if ($h -gt 1000) { $h = 'auto' }
      $lines.Add(("  R{0} h={1} width={2} {3}" -f $r, $h, [math]::Round($sumW, 1), ($vals -join ' ')))
    }
  }
  $lines.Add("TEXT_WIDTH_PT=" + [math]::Round($doc.PageSetup.TextColumns.Width, 1))

  $prevPage = 0; $i = 0
  foreach ($p in $doc.Paragraphs) {
    $i++
    $txt = $p.Range.Text -replace "[`r`n`a`f]", ''
    if ($txt.Trim().Length -eq 0) { continue }
    $pg = $p.Range.Information(3)
    $mark = ''
    if ($pg -ne $prevPage -and $prevPage -ne 0) { $mark = '   <<< PAGE BREAK' }
    $prevPage = $pg
    $short = $txt; if ($short.Length -gt 30) { $short = $short.Substring(0, 30) }
    $lines.Add(("p{0} pg{1} {2}{3}" -f $i, $pg, $short, $mark))
  }

  foreach ($s in $doc.Sections) {
    $f = $s.Footers.Item(1)
    $lines.Add("FOOTER=[" + ($f.Range.Text -replace "[`r`n`a]", '') + "] fields=" + $f.Range.Fields.Count)
  }

  if ($Pdf) {
    if (Test-Path $Pdf) { Remove-Item $Pdf -Force }
    $doc.ExportAsFixedFormat($Pdf, 17)
    $bytes = [System.IO.File]::ReadAllBytes($Pdf)
    $latin = [System.Text.Encoding]::GetEncoding(28591).GetString($bytes)
    $count = ([regex]::Matches($latin, '/Type\s*/Page[^s]')).Count
    $fonts = [regex]::Matches($latin, '/BaseFont\s*/([A-Za-z0-9+#\-,]+)') | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique
    $lines.Insert(0, "PDF_PAGES=$count")
    $lines.Add("PDF_FONTS=" + ($fonts -join ', '))
  }

  $doc.Close($false)
} finally { $word.Quit() }

[System.IO.File]::WriteAllLines($Out, $lines, (New-Object System.Text.UTF8Encoding($false)))
Write-Output ("WROTE " + $Out + " (" + $lines.Count + " lines)")
