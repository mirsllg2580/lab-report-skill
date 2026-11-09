# Repack an unpacked docx folder into a valid .docx (forward-slash entries),
# then let Word confirm it opens.
#   .\pack-docx.ps1 -SrcDir C:\work\unpacked -OutFile D:\out\report.docx [-SkipWordCheck]
param(
  [Parameter(Mandatory = $true)][string]$SrcDir,
  [Parameter(Mandatory = $true)][string]$OutFile,
  [switch]$SkipWordCheck
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.IO.Compression

# Test-Path treats [] as wildcards, so [Content_Types].xml must use -LiteralPath
$ctPath = Join-Path $SrcDir '[Content_Types].xml'
$docPath = Join-Path $SrcDir 'word\document.xml'
if (-not (Test-Path -LiteralPath $ctPath)) {
  throw "Not a docx folder: $SrcDir (missing [Content_Types].xml)"
}
if (-not (Test-Path -LiteralPath $docPath)) {
  throw "Not a docx folder: $SrcDir (missing word\document.xml)"
}

# 包根目录只允许 [Content_Types].xml；中间产物（解包用的 zip、渲染出的 svg/png）
# 混进来会被一起打包，Word 就报"文件可能已经损坏"
$strays = @(Get-ChildItem -LiteralPath $SrcDir -File | Where-Object { $_.Name -ne '[Content_Types].xml' })
if ($strays.Count) {
  throw ("包根目录有非部件文件，先移走再打包: " + (($strays | ForEach-Object { $_.Name }) -join ', '))
}

# hard guard: catching a truncated document.xml early saves a corrupt file
$docXml = Get-Content -LiteralPath $docPath -Raw -Encoding UTF8
if ($docXml -notmatch '</w:body>') { throw 'word\document.xml is truncated (no </w:body>) - refusing to pack' }
if ($docXml -notmatch '<w:sectPr') { throw 'word\document.xml has no sectPr - refusing to pack' }

# XML 语法守卫：Word 对格式错误的反应只有"无法打开该文件"，不指位置。
# 先自己解析所有 xml/rels 部件，报出具体文件与行列（例如 w:cs 属性重复）。
Get-ChildItem -LiteralPath $SrcDir -Recurse -File | Where-Object { $_.Extension -in '.xml', '.rels' } | ForEach-Object {
  $f = $_
  try { (New-Object System.Xml.XmlDocument).Load($f.FullName) }
  catch {
    $e = $_.Exception
    while ($e.InnerException) { $e = $e.InnerException }
    throw ("XML 不合法: " + $f.FullName.Substring($SrcDir.Length + 1) + " 第 " + $e.LineNumber + " 行第 " + $e.LinePosition + " 列: " + $e.Message)
  }
}

$outDir = Split-Path -Parent $OutFile
if ($outDir -and -not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
if (Test-Path $OutFile) { Remove-Item $OutFile -Force }

$zip = [System.IO.Compression.ZipFile]::Open($OutFile, [System.IO.Compression.ZipArchiveMode]::Create)
try {
  Get-ChildItem -Path $SrcDir -Recurse -File | ForEach-Object {
    # [Content_Types].xml must be the first entry by convention
    $rel = $_.FullName.Substring($SrcDir.Length + 1).Replace('\', '/')
    $entry = $zip.CreateEntry($rel, [System.IO.Compression.CompressionLevel]::Optimal)
    $s = $entry.Open()
    try {
      $bytes = [System.IO.File]::ReadAllBytes($_.FullName)
      $s.Write($bytes, 0, $bytes.Length)
    } finally { $s.Dispose() }
  }
} finally { $zip.Dispose() }

Write-Output ("PACKED " + (Get-Item $OutFile).Length + " bytes -> " + $OutFile)

if (-not $SkipWordCheck) {
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  try {
    $doc = $word.Documents.Open($OutFile, $false, $true)
    $doc.Repaginate()
    Write-Output ("WORD_OK paragraphs=" + $doc.Paragraphs.Count + " tables=" + $doc.Tables.Count)
    $doc.Close($false)
  } finally { $word.Quit() }
}
