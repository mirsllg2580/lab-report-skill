# Rough visual preview: Word -> filtered HTML -> headless Edge screenshot.
#   .\render-html-preview.ps1 -Docx D:\out\report.docx -Png D:\out\preview.png [-Width 880] [-Height 1200]
# Caveats: HTML export has no pagination and auto-width tables shrink in the
# browser. Use it for glyph / indent / overlap / image-presence checks only;
# trust verify-docx.ps1 for page count and table width, and the exported PDF's
# /Subtype /Image count for whether pictures really made it in.
param(
  [Parameter(Mandatory = $true)][string]$Docx,
  [Parameter(Mandatory = $true)][string]$Png,
  [int]$Width = 880,
  [int]$Height = 1200,
  [string]$Edge = 'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'
)
$ErrorActionPreference = 'Stop'
if (-not (Test-Path -LiteralPath $Edge)) { throw "Edge not found: $Edge (set -Edge)" }
# Edge resolves a relative --screenshot= against its own cwd, and would silently
# leave an old file in place. Always hand it an absolute path and clear the target.
$Png = [System.IO.Path]::GetFullPath($Png)
$pngDir = Split-Path -Parent $Png
if ($pngDir -and -not (Test-Path -LiteralPath $pngDir)) { New-Item -ItemType Directory -Path $pngDir -Force | Out-Null }

# Word names the image folder after the html file, and the html hard-codes that
# folder name, so html + folder must keep matching names. Non-ASCII paths also
# break the file:// URL, so work from an all-ASCII staging dir.
$stage = Join-Path ([System.IO.Path]::GetTempPath()) ('prev-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $stage -Force | Out-Null
$html = Join-Path $stage 'report.preview.html'
$filesDir = Join-Path $stage 'report.preview.files'

$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
  Copy-Item -LiteralPath $Docx -Destination (Join-Path $stage 'report.docx') -Force
  $doc = $word.Documents.Open((Join-Path $stage 'report.docx'), $false, $true)
  $doc.SaveAs2($html, 10)   # wdFormatFilteredHTML
  $doc.Close($false)
} finally { $word.Quit() }

try {
  if (-not (Test-Path -LiteralPath $html)) { throw "html export missing: $html" }
  if (Test-Path -LiteralPath $Png) { Remove-Item -LiteralPath $Png -Force }
  $url = 'file:///' + ($html -replace '\\', '/')
  & $Edge --headless=new --disable-gpu --hide-scrollbars --force-device-scale-factor=1.5 `
    --window-size="$Width,$Height" --screenshot="$Png" $url 2>&1 | Out-String | Write-Verbose
} finally {
  Remove-Item $stage -Recurse -Force -ErrorAction SilentlyContinue
}

if (Test-Path -LiteralPath $Png) { Write-Output ("PNG " + (Get-Item -LiteralPath $Png).Length + " bytes -> " + $Png) }
else { throw 'screenshot not produced' }
