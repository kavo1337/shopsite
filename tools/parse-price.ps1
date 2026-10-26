$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Drawing

$src = (Get-ChildItem "C:\Users\clost\Downloads\*.xlsx" | Where-Object { $_.Length -gt 10MB } | Select-Object -First 1).FullName
if (-not $src) { throw "Price xlsx not found in Downloads" }
Write-Host "Source: $src"
$root = "D:\Cursor\Project"
$imgOut = Join-Path $root "images"
$dataOut = Join-Path $root "data"
$extract = Join-Path $env:TEMP "prays_xlsx_extract"

if (Test-Path $extract) { Remove-Item $extract -Recurse -Force }
if (Test-Path $imgOut) { Remove-Item $imgOut -Recurse -Force }
New-Item -ItemType Directory -Force -Path $extract, $imgOut, $dataOut | Out-Null
[System.IO.Compression.ZipFile]::ExtractToDirectory($src, $extract)

function U([int[]]$codes) {
  -join ($codes | ForEach-Object { [char]$_ })
}

$T_WHITE = U @(0x0411,0x0435,0x043B,0x044B,0x0439)
$T_GRAPHITE = U @(0x0413,0x0440,0x0430,0x0444,0x0438,0x0442)
$T_BLACK = U @(0x0427,0x0451,0x0440,0x043D,0x044B,0x0439)
$T_BLACK2 = U @(0x0427,0x0435,0x0440,0x043D,0x044B,0x0439)
$T_CLEAR = U @(0x041F,0x0440,0x043E,0x0437,0x0440,0x0430,0x0447,0x043D,0x044B,0x0439)
$T_PRICES = U @(0x0426,0x0435,0x043D,0x044B)
$T_PIPES = U @(0x0442,0x0440,0x0443,0x0431)
$T_CONN = U @(0x0421,0x043E,0x0435,0x0434,0x0435,0x043D)
$T_SEAL = U @(0x0423,0x043F,0x043B,0x043E,0x0442,0x043D)
$T_PHOTO = U @(0x0424,0x043E,0x0442,0x043E)
$T_DESC = U @(0x041E,0x043F,0x0438,0x0441,0x0430,0x043D,0x0438,0x0435)
$T_ART = U @(0x0410,0x0440,0x0442,0x0438,0x043A,0x0443,0x043B)
$T_PRICE = U @(0x0426,0x0435,0x043D,0x0430)
$T_WH = U @(0x0421,0x043A,0x043B,0x0430,0x0434)
$T_REST = U @(0x041E,0x0441,0x0442,0x0430,0x0442,0x043A,0x0438)
$T_ORDER = U @(0x0417,0x0430,0x043A,0x0430,0x0437)
$T_SUM = U @(0x0418,0x0442,0x043E,0x0433,0x043E)
$T_MAT = U @(0x041C,0x0430,0x0442,0x0435,0x0440,0x0438,0x0430,0x043B)
$T_COLORS = U @(0x0426,0x0432,0x0435,0x0442,0x0430)
$T_TITLE = U @(0x041F,0x0440,0x0430,0x0439,0x0441,0x002D,0x043B,0x0438,0x0441,0x0442)

function Get-SharedStrings([string]$path) {
  [xml]$ssxml = Get-Content -Path $path -Encoding UTF8 -Raw
  $nsmgr = New-Object System.Xml.XmlNamespaceManager($ssxml.NameTable)
  $nsmgr.AddNamespace("m", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
  $list = New-Object System.Collections.Generic.List[string]
  foreach ($si in $ssxml.SelectNodes("//m:si", $nsmgr)) {
    $tNodes = $si.SelectNodes(".//m:t", $nsmgr)
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($t in $tNodes) { $parts.Add([string]$t.InnerText) }
    $list.Add(($parts -join ""))
  }
  return $list
}

function Get-Rels([string]$path) {
  $map = @{}
  if (-not (Test-Path $path)) { return $map }
  [xml]$xml = Get-Content -Path $path -Encoding UTF8 -Raw
  foreach ($rel in $xml.Relationships.Relationship) {
    $map[$rel.Id] = $rel.Target
  }
  return $map
}

function Get-ColRow([string]$ref) {
  if ($ref -notmatch '^([A-Z]+)(\d+)$') { return @{ Col = ""; Row = 0 } }
  return @{ Col = $Matches[1]; Row = [int]$Matches[2] }
}

function Clean-Text([string]$s) {
  if ($null -eq $s) { return "" }
  $s = $s -replace [char]160, " "
  $s = $s.Trim()
  $s = $s -replace '\s+', ' '
  return $s
}

function Parse-Price($raw) {
  if ($null -eq $raw -or $raw -eq "") { return $null }
  $t = Clean-Text ([string]$raw)
  if ($t -eq "" -or $t -eq "#VALUE!") { return $null }
  $n = 0.0
  if ([double]::TryParse($t, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$n)) {
    return [int][Math]::Round($n + 1000)
  }
  return $null
}

function Is-HeaderText([string]$s) {
  $x = Clean-Text $s
  if ($x -eq "") { return $false }
  if ($x.StartsWith($T_PHOTO)) { return $true }
  if ($x.StartsWith($T_DESC)) { return $true }
  if ($x.StartsWith($T_ART)) { return $true }
  if ($x.StartsWith($T_PRICE)) { return $true }
  if ($x.StartsWith($T_WH)) { return $true }
  if ($x.StartsWith($T_REST)) { return $true }
  if ($x.StartsWith($T_ORDER)) { return $true }
  if ($x.StartsWith($T_SUM)) { return $true }
  if ($x.StartsWith($T_MAT)) { return $true }
  if ($x.StartsWith($T_COLORS)) { return $true }
  return $false
}

function Get-SheetMap([string]$sheetPath, $strings) {
  [xml]$xml = Get-Content -Path $sheetPath -Encoding UTF8 -Raw
  $nsmgr = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
  $nsmgr.AddNamespace("m", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
  $rows = @{}
  foreach ($row in $xml.SelectNodes("//m:sheetData/m:row", $nsmgr)) {
    $ridx = [int]$row.r
    $cells = @{}
    foreach ($c in $row.SelectNodes("m:c", $nsmgr)) {
      $cr = Get-ColRow $c.r
      $t = $c.t
      $v = $c.SelectSingleNode("m:v", $nsmgr)
      $is = $c.SelectSingleNode("m:is", $nsmgr)
      $val = ""
      if ($t -eq "s" -and $v) {
        $idx = [int]$v.InnerText
        if ($idx -ge 0 -and $idx -lt $strings.Count) { $val = $strings[$idx] }
      } elseif ($t -eq "inlineStr" -and $is) {
        $val = $is.InnerText
      } elseif ($v) {
        $val = $v.InnerText
      }
      $cells[$cr.Col] = Clean-Text $val
    }
    $rows[$ridx] = $cells
  }
  return $rows
}

function Get-RowImages([string]$drawingPath, [string]$drawingRelsPath) {
  $byRow = @{}
  if (-not (Test-Path $drawingPath)) { return $byRow }
  $rels = Get-Rels $drawingRelsPath
  $xmlText = Get-Content -Path $drawingPath -Encoding UTF8 -Raw
  $anchorRegex = [regex]'(?s)<xdr:(?:twoCellAnchor|oneCellAnchor)[^>]*>.*?</xdr:(?:twoCellAnchor|oneCellAnchor)>'
  foreach ($m in $anchorRegex.Matches($xmlText)) {
    $block = $m.Value
    $rowM = [regex]::Match($block, '<xdr:row>(\d+)</xdr:row>')
    $embM = [regex]::Match($block, 'r:embed="(rId\d+)"')
    if (-not $rowM.Success -or -not $embM.Success) { continue }
    $excelRow = [int]$rowM.Groups[1].Value + 1
    $rid = $embM.Groups[1].Value
    if (-not $rels.ContainsKey($rid)) { continue }
    $target = $rels[$rid]
    $abs = [System.IO.Path]::GetFullPath((Join-Path (Split-Path $drawingPath) ($target -replace '/','\')))
    if (-not $byRow.ContainsKey($excelRow)) { $byRow[$excelRow] = New-Object System.Collections.Generic.List[string] }
    if (-not $byRow[$excelRow].Contains($abs)) { [void]$byRow[$excelRow].Add($abs) }
  }
  return $byRow
}

function Get-DrawingForSheet([string]$sheetRelsPath) {
  $rels = Get-Rels $sheetRelsPath
  foreach ($k in $rels.Keys) {
    if ($rels[$k] -match 'drawing\d+\.xml$') { return $rels[$k] }
  }
  return $null
}

$copied = @{}
function Copy-ResizedImage([string]$srcPath, [string]$destName) {
  if ($copied.ContainsKey($srcPath)) { return $copied[$srcPath] }
  $dest = Join-Path $imgOut $destName
  $img = $null
  $bmp = $null
  $g = $null
  try {
    $img = [System.Drawing.Image]::FromFile($srcPath)
    $max = 480
    $w = $img.Width; $h = $img.Height
    if ($w -gt $max -or $h -gt $max) {
      $scale = [Math]::Min($max / $w, $max / $h)
      $nw = [Math]::Max(1, [int]($w * $scale))
      $nh = [Math]::Max(1, [int]($h * $scale))
    } else {
      $nw = $w; $nh = $h
    }
    $bmp = New-Object System.Drawing.Bitmap $nw, $nh
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.DrawImage($img, 0, 0, $nw, $nh)
    $codec = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq "image/jpeg" }
    $enc = New-Object System.Drawing.Imaging.EncoderParameters 1
    $enc.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter ([System.Drawing.Imaging.Encoder]::Quality, [long]72)
    $bmp.Save($dest, $codec, $enc)
  } catch {
    Copy-Item $srcPath $dest -Force
  } finally {
    if ($g) { $g.Dispose() }
    if ($bmp) { $bmp.Dispose() }
    if ($img) { $img.Dispose() }
  }
  $rel = "images/$destName"
  $copied[$srcPath] = $rel
  return $rel
}

$imgSeq = 0
function Image-ForRow($rowImages, [int]$row) {
  if (-not $rowImages.ContainsKey($row)) { return $null }
  $srcPath = $rowImages[$row][0]
  if (-not (Test-Path $srcPath)) { return $null }
  $script:imgSeq++
  $name = ("p{0:D4}.jpg" -f $script:imgSeq)
  return Copy-ResizedImage $srcPath $name
}

$strings = Get-SharedStrings (Join-Path $extract "xl\sharedStrings.xml")

$wbRels = Get-Rels (Join-Path $extract "xl\_rels\workbook.xml.rels")
[xml]$wb = Get-Content (Join-Path $extract "xl\workbook.xml") -Encoding UTF8 -Raw
$nsmgrWb = New-Object System.Xml.XmlNamespaceManager($wb.NameTable)
$nsmgrWb.AddNamespace("m", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")

$visibleSheets = @()
foreach ($sh in $wb.SelectNodes("//m:sheets/m:sheet", $nsmgrWb)) {
  $state = $sh.GetAttribute("state")
  if ($state -eq "hidden" -or $state -eq "veryHidden") { continue }
  $rid = $sh.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
  $name = Clean-Text $sh.GetAttribute("name")
  $target = $wbRels[$rid]
  $visibleSheets += @{ Name = $name; File = (Join-Path $extract "xl\$($target -replace '/','\')") }
}

$script:idSeq = 0
function New-Slug([string]$s) {
  $script:idSeq++
  return ("s{0}" -f $script:idSeq)
}

function Is-ColorHead([string]$a) {
  $x = Clean-Text $a
  foreach ($h in @($T_WHITE, $T_GRAPHITE, $T_BLACK, $T_BLACK2, $T_CLEAR)) {
    if ($x.StartsWith($h)) { return $true }
  }
  return $false
}

function Parse-Standard($rows, $rowImages, $catId, $catName, [bool]$seals) {
  $sections = New-Object System.Collections.Generic.List[object]
  $current = $null
  $note = $null
  $orderedKeys = $rows.Keys | Sort-Object
  foreach ($r in $orderedKeys) {
    $c = $rows[$r]
    $a = if ($c.Contains("A")) { $c["A"] } else { "" }
    $b = if ($c.Contains("B")) { $c["B"] } else { "" }
    $cc = if ($c.Contains("C")) { $c["C"] } else { "" }
    $d = if ($c.Contains("D")) { $c["D"] } else { "" }

    if ($seals) {
      if ($a.StartsWith($T_PRICES) -and $a.Length -gt 8) { $note = $a; continue }
      if (Is-ColorHead $a) {
        $sid = "$catId-" + (New-Slug $a)
        $current = @{ id = $sid; title = $a; note = $null; items = (New-Object System.Collections.Generic.List[object]) }
        [void]$sections.Add($current)
        continue
      }
    }

    if (Is-HeaderText $a) { continue }

    $dLooksNumber = $false
    $tmp = 0.0
    if ($d -ne "") {
      $dLooksNumber = [double]::TryParse($d, [ref]$tmp)
    }
    $isSection = ($a -ne "" -and $b -eq "" -and $cc -eq "" -and -not $dLooksNumber)
    if ($isSection -and $a -ne "#VALUE!") {
      $sid = "$catId-" + (New-Slug $a)
      $current = @{ id = $sid; title = $a; note = $null; items = (New-Object System.Collections.Generic.List[object]) }
      [void]$sections.Add($current)
      continue
    }

    $name = $b
    $art = $cc
    if ($name -eq "#VALUE!") { $name = "" }
    if ($art -eq "#VALUE!") { $art = "" }
    if ($name -eq "" -and $art -eq "") { continue }

    $price = Parse-Price $d
    $photo = Image-ForRow $rowImages $r
    $item = @{
      name = $(if ($name) { $name } else { $art })
      article = $art
      price = $price
      photo = $photo
      color = $null
    }
    if ($null -eq $current) {
      $current = @{ id = $catId; title = $catName; note = $null; items = (New-Object System.Collections.Generic.List[object]) }
      [void]$sections.Add($current)
    }
    [void]$current.items.Add($item)
  }
  if ($note -and $sections.Count -gt 0) {
    foreach ($sec in $sections) { if (-not $sec.note) { $sec.note = $note } }
  }
  return $sections
}

function Parse-Connectors($rows, $rowImages, $catId, $catName) {
  $sections = New-Object System.Collections.Generic.List[object]
  $current = $null
  $lastArt = ""
  $lastMat = ""
  $lastPhoto = $null
  $orderedKeys = $rows.Keys | Sort-Object
  foreach ($r in $orderedKeys) {
    $c = $rows[$r]
    $a = if ($c.Contains("A")) { $c["A"] } else { "" }
    $ccol = if ($c.Contains("C")) { $c["C"] } else { "" }
    $d = if ($c.Contains("D")) { $c["D"] } else { "" }
    $e = if ($c.Contains("E")) { $c["E"] } else { "" }
    if (Is-HeaderText $a) { continue }
    if ($a -ne "" -and $a.ToLower().Contains($T_PIPES)) {
      $sid = "$catId-" + (New-Slug $a)
      $current = @{ id = $sid; title = $a; note = $null; items = (New-Object System.Collections.Generic.List[object]) }
      [void]$sections.Add($current)
      $lastArt = ""; $lastMat = ""; $lastPhoto = $null
      continue
    }
    $art = $a
    if ($art -eq "") { $art = $lastArt } else {
      $lastArt = $art
      $ph = Image-ForRow $rowImages $r
      if ($ph) { $lastPhoto = $ph }
    }
    if ($ccol -ne "") { $lastMat = $ccol }
    $color = $d
    if ($art -eq "") { continue }
    if ($color -eq "" -and $ccol -eq "") { continue }
    $price = Parse-Price $e
    $name = $art
    if ($color) { $name = "$art - $color" }
    $item = @{
      name = $name
      article = $art
      price = $price
      photo = $lastPhoto
      color = $(if ($color) { $color } else { $null })
      material = $(if ($lastMat) { $lastMat } else { $null })
    }
    if ($null -eq $current) {
      $current = @{ id = $catId; title = $catName; note = $null; items = (New-Object System.Collections.Generic.List[object]) }
      [void]$sections.Add($current)
    }
    [void]$current.items.Add($item)
  }
  return $sections
}

$catalog = @{
  title = $T_TITLE
  priceNote = "Prices from Excel + 1000 RUB"
  categories = (New-Object System.Collections.Generic.List[object])
}

foreach ($sh in $visibleSheets) {
  $name = $sh.Name
  $file = $sh.File
  $id = New-Slug $name
  Write-Host "Parsing $name"
  $rows = Get-SheetMap $file $strings
  $sheetDir = Split-Path $file
  $sheetRels = Join-Path $sheetDir ("_rels\" + [IO.Path]::GetFileName($file) + ".rels")
  $drawingRel = Get-DrawingForSheet $sheetRels
  $rowImages = @{}
  if ($drawingRel) {
    $drawingPath = [IO.Path]::GetFullPath((Join-Path $sheetDir ($drawingRel -replace '/','\')))
    $drawingRels = Join-Path (Split-Path $drawingPath) ("_rels\" + [IO.Path]::GetFileName($drawingPath) + ".rels")
    $rowImages = Get-RowImages $drawingPath $drawingRels
    Write-Host ("  image rows: {0}" -f $rowImages.Count)
  }
  if ($name.Contains($T_CONN)) {
    $sections = Parse-Connectors $rows $rowImages $id $name
  } elseif ($name.Contains($T_SEAL)) {
    $sections = Parse-Standard $rows $rowImages $id $name $true
  } else {
    $sections = Parse-Standard $rows $rowImages $id $name $false
  }
  $count = 0
  for ($si = 0; $si -lt $sections.Count; $si++) {
    $count += $sections[$si].items.Count
    Write-Host ("  section[{0}] {1} n={2}" -f $si, $sections[$si].title, $sections[$si].items.Count)
  }
  Write-Host "  items: $count sections: $($sections.Count) type=$($sections.GetType().Name)"
  [void]$catalog.categories.Add(@{
    id = $id
    title = $name
    sections = $sections
  })
}

function Escape-Json([string]$s) {
  if ($null -eq $s) { return "" }
  $s = $s.Replace('\','\\').Replace('"','\"').Replace("`r",'\r').Replace("`n",'\n').Replace("`t",'\t')
  return $s
}

function ConvertTo-JsonText($obj) {
  if ($null -eq $obj) { return "null" }
  $t = $obj.GetType()
  if ($obj -is [string]) { return '"' + (Escape-Json $obj) + '"' }
  if ($obj -is [bool]) { if ($obj) { return "true" } else { return "false" } }
  if ($obj -is [byte] -or $obj -is [int16] -or $obj -is [int] -or $obj -is [int64] -or $obj -is [decimal] -or $obj -is [double] -or $obj -is [float] -or $obj -is [long]) {
    return ([string]$obj)
  }
  if ($obj -is [System.Collections.IDictionary]) {
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($k in $obj.Keys) {
      $parts.Add(('"{0}":{1}' -f (Escape-Json ([string]$k)), (ConvertTo-JsonText $obj[$k])))
    }
    return '{' + ($parts -join ',') + '}'
  }
  if ($obj -is [System.Collections.IEnumerable]) {
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($el in $obj) { $parts.Add((ConvertTo-JsonText $el)) }
    return '[' + ($parts -join ',') + ']'
  }
  return '"' + (Escape-Json ([string]$obj)) + '"'
}

$jsonPath = Join-Path $dataOut "catalog.json"
$json = ConvertTo-JsonText $catalog
[IO.File]::WriteAllText($jsonPath, $json, [Text.UTF8Encoding]::new($false))
Write-Host "Wrote $jsonPath"
Write-Host ("Unique images: {0}" -f $copied.Count)
$imgFiles = Get-ChildItem $imgOut -ErrorAction SilentlyContinue
if ($imgFiles) {
  $sum = ($imgFiles | Measure-Object Length -Sum).Sum
  Write-Host ("Images MB: {0:N1}" -f ($sum/1MB))
}
