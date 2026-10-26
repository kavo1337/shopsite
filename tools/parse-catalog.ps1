$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Drawing

$xlsxPath = Get-ChildItem "$env:USERPROFILE\Downloads\*.xlsx" | Where-Object { $_.Length -gt 20MB } | Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty FullName
if (-not $xlsxPath) { throw "xlsx not found in Downloads" }
Write-Host "XLSX: $xlsxPath"
$root = "D:\Cursor\Project"
$outData = Join-Path $root "data"
$outImages = Join-Path $root "images"
$temp = Join-Path $env:TEMP "price-xlsx-extract"

if (Test-Path $temp) { Remove-Item $temp -Recurse -Force }
New-Item -ItemType Directory -Path $temp | Out-Null
New-Item -ItemType Directory -Path $outData -Force | Out-Null
New-Item -ItemType Directory -Path $outImages -Force | Out-Null
Get-ChildItem $outImages -File -ErrorAction SilentlyContinue | Remove-Item -Force

[System.IO.Compression.ZipFile]::ExtractToDirectory($xlsxPath, $temp)

function Get-Xml($relPath) {
  $full = Join-Path $temp $relPath
  $xml = New-Object System.Xml.XmlDocument
  $xml.PreserveWhitespace = $false
  $xml.Load($full)
  return $xml
}

function XPath($node, [string]$xpath) {
  return $node.SelectNodes($xpath)
}
function XPathOne($node, [string]$xpath) {
  return $node.SelectSingleNode($xpath)
}

function Normalize-Text([string]$s) {
  if ($null -eq $s) { return "" }
  $s = $s -replace "`r|`n", " "
  $s = $s -replace "\s+", " "
  return $s.Trim()
}

$ssXml = Get-Xml "xl\sharedStrings.xml"
$strings = New-Object System.Collections.Generic.List[string]
foreach ($si in $ssXml.SelectNodes("//*[local-name()='si']")) {
  $tNodes = $si.SelectNodes(".//*[local-name()='t']")
  $parts = New-Object System.Collections.Generic.List[string]
  foreach ($t in $tNodes) { $parts.Add([string]$t.InnerText) }
  $strings.Add(($parts -join ""))
}

function Cell-Value($c, $strings) {
  $t = $c.GetAttribute("t")
  if ($t -eq "s") {
    $v = $c.SelectSingleNode("*[local-name()='v']")
    if ($v) { return [string]$strings[[int]$v.InnerText] }
  }
  if ($t -eq "inlineStr") {
    $tnode = $c.SelectSingleNode(".//*[local-name()='t']")
    if ($tnode) { return [string]$tnode.InnerText }
  }
  if ($t -eq "str") {
    $v = $c.SelectSingleNode("*[local-name()='v']")
    if ($v) { return [string]$v.InnerText }
  }
  $v = $c.SelectSingleNode("*[local-name()='v']")
  if ($v) { return [string]$v.InnerText }
  return ""
}

function Col-From-Ref([string]$ref) {
  $letters = ($ref -replace "[0-9]", "")
  $n = 0
  foreach ($ch in $letters.ToCharArray()) {
    $n = $n * 26 + ([int][char]$ch - 64)
  }
  return $n
}

function Get-SheetRows($sheetRel) {
  $xml = Get-Xml $sheetRel
  $rows = @{}
  foreach ($row in $xml.SelectNodes("//*[local-name()='sheetData']/*[local-name()='row']")) {
    $ridx = [int]$row.GetAttribute("r")
    $map = @{}
    foreach ($c in $row.SelectNodes("*[local-name()='c']")) {
      $ref = $c.GetAttribute("r")
      $col = Col-From-Ref $ref
      $map[$col] = Normalize-Text (Cell-Value $c $strings)
    }
    $rows[$ridx] = $map
  }
  return $rows
}

function Get-DrawingMap($sheetRel) {
  $relsPath = $sheetRel -replace "worksheets\\", "worksheets\_rels\" -replace "worksheets/", "worksheets/_rels/"
  $relsPath = $relsPath + ".rels"
  $fullRels = Join-Path $temp ("xl\worksheets\_rels\" + [IO.Path]::GetFileName($sheetRel) + ".rels")
  $rowToImage = @{}
  if (-not (Test-Path $fullRels)) { return $rowToImage }

  $relsXml = Get-Xml ("xl\worksheets\_rels\" + [IO.Path]::GetFileName($sheetRel) + ".rels")
  $drawingTarget = $null
  foreach ($rel in $relsXml.SelectNodes("//*[local-name()='Relationship']")) {
    if ($rel.Type -like "*drawing") {
      $drawingTarget = ($rel.Target -replace "^\.\./", "xl/") -replace "/", "\"
    }
  }
  if (-not $drawingTarget) { return $rowToImage }

  $drawXml = Get-Xml $drawingTarget
  $drawDir = Split-Path $drawingTarget -Parent
  $drawRels = Join-Path $drawDir ("_rels\" + [IO.Path]::GetFileName($drawingTarget) + ".rels")
  $embedToFile = @{}
  if (Test-Path (Join-Path $temp $drawRels)) {
    $drels = Get-Xml $drawRels
    foreach ($rel in $drels.SelectNodes("//*[local-name()='Relationship']")) {
      if ($rel.Type -like "*image") {
        $target = ($rel.Target -replace "^\.\./", "xl/") -replace "/", "\"
        $embedToFile[$rel.Id] = $target
      }
    }
  }

  foreach ($anchor in $drawXml.SelectNodes("//*[local-name()='twoCellAnchor' or local-name()='oneCellAnchor']")) {
    $fromRow = $anchor.SelectSingleNode("*[local-name()='from']/*[local-name()='row']")
    if (-not $fromRow) { continue }
    $excelRow = [int]$fromRow.InnerText + 1
    $blip = $anchor.SelectSingleNode(".//*[local-name()='blip']")
    if (-not $blip) { continue }
    $embed = $blip.GetAttribute("embed")
    if (-not $embed) {
      $attr = $blip.Attributes | Where-Object { $_.LocalName -eq "embed" } | Select-Object -First 1
      if ($attr) { $embed = $attr.Value }
    }
    if ($embed -and $embedToFile.ContainsKey($embed) -and -not $rowToImage.ContainsKey($excelRow)) {
      $rowToImage[$excelRow] = $embedToFile[$embed]
    }
  }
  return $rowToImage
}

$usedImages = @{}
$imgCounter = 0

function Copy-ProductImage($mediaRel) {
  if (-not $mediaRel) { return $null }
  if ($usedImages.ContainsKey($mediaRel)) { return $usedImages[$mediaRel] }
  $src = Join-Path $temp $mediaRel
  if (-not (Test-Path $src)) { return $null }
  $script:imgCounter++
  $ext = [IO.Path]::GetExtension($src).ToLower()
  if ($ext -notin @(".jpg", ".jpeg", ".png")) { $ext = ".jpg" }
  $name = "p{0:D4}{1}" -f $script:imgCounter, $(if ($ext -eq ".png") { ".jpg" } else { $ext })
  $dest = Join-Path $outImages $name
  try {
    $img = [System.Drawing.Image]::FromFile($src)
    try {
      $max = 480
      $w = $img.Width
      $h = $img.Height
      if ($w -gt $max -or $h -gt $max) {
        if ($w -ge $h) { $nw = $max; $nh = [int]($h * $max / $w) }
        else { $nh = $max; $nw = [int]($w * $max / $h) }
      } else { $nw = $w; $nh = $h }
      $bmp = New-Object System.Drawing.Bitmap $nw, $nh
      $g = [System.Drawing.Graphics]::FromImage($bmp)
      $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
      $g.DrawImage($img, 0, 0, $nw, $nh)
      $g.Dispose()
      $codec = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq "image/jpeg" }
      $enc = [System.Drawing.Imaging.Encoder]::Quality
      $ep = New-Object System.Drawing.Imaging.EncoderParameters 1
      $ep.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter $enc, ([int64]72)
      $bmp.Save($dest, $codec, $ep)
      $bmp.Dispose()
    } finally { $img.Dispose() }
  } catch {
    Copy-Item $src $dest -Force
  }
  $web = "images/$name"
  $usedImages[$mediaRel] = $web
  return $web
}

function Parse-Price([string]$raw) {
  if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
  $n = 0.0
  if ([double]::TryParse($raw, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$n)) {
    return [int][Math]::Round($n + 1000)
  }
  return $null
}

function New-ItemObj($name, $article, $price, $image, $color, $material) {
  $obj = [ordered]@{
    name = $name
    article = $article
    price = $price
    image = $image
  }
  if ($color) { $obj.color = $color }
  if ($material) { $obj.material = $material }
  return $obj
}

function Get($map, $col) {
  if ($null -eq $map) { return "" }
  if ($map.ContainsKey($col)) { return [string]$map[$col] }
  return ""
}

function Is-HeaderRow([string]$a, [string]$b, [string]$c, [string]$d) {
  $price = Parse-Price $d
  if ($c -and $d -and $null -eq $price) { return $true }
  return $false
}

function Parse-Standard($rows, $drawMap, $sectionHints) {
  $sections = New-Object System.Collections.Generic.List[object]
  $currentName = ""
  $currentNote = ""
  $currentItems = New-Object System.Collections.Generic.List[object]
  $orderedKeys = $rows.Keys | Sort-Object

  function Flush-Section {
    param($force)
    if ($currentItems.Count -gt 0 -or ($currentName -and $force)) {
      $sec = [ordered]@{
        name = $(if ($currentName) { $currentName } else { "" })
        items = @($currentItems.ToArray())
      }
      if ($currentNote) { $sec.note = $currentNote }
      $sections.Add($sec)
      $script:currentItems = New-Object System.Collections.Generic.List[object]
    }
  }

  $currentItems = New-Object System.Collections.Generic.List[object]
  foreach ($r in $orderedKeys) {
    $map = $rows[$r]
    $a = Get $map 1
    $b = Get $map 2
    $c = Get $map 3
    $d = Get $map 4

    if (Is-HeaderRow $a $b $c $d) { continue }

    $dummy = 0.0
    $hasPrice = [double]::TryParse($d, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$dummy)
    $isSection = $false
    if ($a -and -not $c -and -not $hasPrice -and $a -ne "#VALUE!") { $isSection = $true }

    if ($isSection) {
      if ($a.Length -gt 28) {
        $currentNote = $a
        continue
      }
      if ($currentItems.Count -gt 0 -or $currentName) {
        $sec = [ordered]@{ name = $(if ($currentName) { $currentName } else { "" }); items = @($currentItems.ToArray()) }
        if ($currentNote) { $sec.note = $currentNote }
        $sections.Add($sec)
      }
      $currentName = $a
      $currentNote = ""
      $currentItems = New-Object System.Collections.Generic.List[object]
      continue
    }

    $name = $b
    $article = $c
    if ($name -eq "#VALUE!") { $name = "" }
    if (-not $name -and -not $article) { continue }
    if (-not $name) { $name = $article }

    $price = Parse-Price $d
    $imgPath = $null
    if ($drawMap.ContainsKey($r)) { $imgPath = Copy-ProductImage $drawMap[$r] }

    $currentItems.Add((New-ItemObj $name $article $price $imgPath $null $null))
  }

  if ($currentItems.Count -gt 0 -or $currentName) {
    $sec = [ordered]@{ name = $(if ($currentName) { $currentName } else { "" }); items = @($currentItems.ToArray()) }
    if ($currentNote) { $sec.note = $currentNote }
    $sections.Add($sec)
  }
  if ($sections.Count -eq 0) {
    $sections.Add([ordered]@{ name = ""; items = @() })
  }
  return ,$sections.ToArray()
}

function Parse-Connectors($rows, $drawMap) {
  $sections = New-Object System.Collections.Generic.List[object]
  $currentName = ""
  $currentItems = New-Object System.Collections.Generic.List[object]
  $article = ""
  $material = ""
  $groupImage = $null
  $orderedKeys = $rows.Keys | Sort-Object

  foreach ($r in $orderedKeys) {
    $map = $rows[$r]
    $a = Get $map 1
    $c = Get $map 3
    $d = Get $map 4
    $e = Get $map 5

    if ($a -and -not $d -and -not $e) {
      if ($currentItems.Count -gt 0 -or $currentName) {
        $sections.Add([ordered]@{ name = $currentName; items = @($currentItems.ToArray()) })
      }
      $currentName = $a
      $currentItems = New-Object System.Collections.Generic.List[object]
      $article = ""
      $material = ""
      $groupImage = $null
      continue
    }
    if (Is-HeaderRow $a (Get $map 2) $c $e) { continue }
    if (-not $a -and -not $d -and -not $e) { continue }

    if ($a) {
      $article = $a
      $material = $c
      if ($drawMap.ContainsKey($r)) { $groupImage = Copy-ProductImage $drawMap[$r] }
      else { $groupImage = $null }
    } elseif ($drawMap.ContainsKey($r) -and -not $groupImage) {
      $groupImage = Copy-ProductImage $drawMap[$r]
    }

    $color = $d
    if (-not $article -and -not $color) { continue }
    $price = Parse-Price $e
    $name = $(if ($color) { "$article, $color" } else { $article })
    $currentItems.Add((New-ItemObj $name $article $price $groupImage $color $material))
  }
  if ($currentItems.Count -gt 0 -or $currentName) {
    $sections.Add([ordered]@{ name = $currentName; items = @($currentItems.ToArray()) })
  }
  return ,$sections.ToArray()
}

$wb = Get-Xml "xl\workbook.xml"
$wbRels = Get-Xml "xl\_rels\workbook.xml.rels"
$ridToFile = @{}
foreach ($rel in $wbRels.SelectNodes("//*[local-name()='Relationship']")) {
  if ($rel.Type -like "*worksheet") {
    $ridToFile[$rel.Id] = ("xl\" + ($rel.Target -replace "/", "\"))
  }
}
$idBySheet = @{
  "sheet1.xml" = "petli"
  "sheet2.xml" = "razdvizhki"
  "sheet3.xml" = "connectors"
  "sheet4.xml" = "soediniteli"
  "sheet5.xml" = "truby"
  "sheet6.xml" = "napravlyayushchie"
  "sheet7.xml" = "ruchki"
  "sheet8.xml" = "uplotnitel"
  "sheet10.xml" = "roliki"
}
$categoriesMeta = @()
foreach ($sh in $wb.SelectNodes("//*[local-name()='sheet']")) {
  $state = $sh.GetAttribute("state")
  if ($state -eq "hidden") { continue }
  $rid = $sh.GetAttribute("id")
  if (-not $rid) { $rid = $sh.Attributes | Where-Object { $_.LocalName -eq "id" } | Select-Object -First 1 -ExpandProperty Value }
  $file = $ridToFile[$rid]
  $sheetFile = [IO.Path]::GetFileName($file)
  if (-not $idBySheet.ContainsKey($sheetFile)) { continue }
  $kind = "standard"
  if ($sheetFile -eq "sheet4.xml") { $kind = "connectors" }
  $categoriesMeta += @{ id = $idBySheet[$sheetFile]; name = (Normalize-Text $sh.GetAttribute("name")); file = $file; kind = $kind }
}

$categories = New-Object System.Collections.Generic.List[object]
foreach ($meta in $categoriesMeta) {
  Write-Host "Parsing $($meta.name)..."
  $rows = Get-SheetRows $meta.file
  $drawMap = Get-DrawingMap $meta.file
  Write-Host "  rows=$($rows.Count) imagesMapped=$($drawMap.Count)"
  if ($meta.kind -eq "connectors") {
    $sections = Parse-Connectors $rows $drawMap
  } else {
    $sections = Parse-Standard $rows $drawMap $null
  }
  $itemCount = 0
  foreach ($s in $sections) { $itemCount += $s.items.Count }
  Write-Host "  sections=$($sections.Count) items=$itemCount"
  $catObj = New-Object PSObject -Property @{
    id = $meta.id
    name = $meta.name
    sections = @($sections | ForEach-Object { $_ })
  }
  $categories.Add($catObj)
}

function Json-Escape([string]$s) {
  if ($null -eq $s) { return "" }
  $s = $s.Replace("\", "\\").Replace('"', '\"').Replace("`r", "\r").Replace("`n", "\n").Replace("`t", "\t")
  return $s
}

$sb = New-Object System.Text.StringBuilder
[void]$sb.AppendLine("{")
[void]$sb.AppendLine('  "title": "Price",')
[void]$sb.AppendLine('  "priceNote": "+1000",')
[void]$sb.AppendLine('  "categories": [')
for ($ci = 0; $ci -lt $categories.Count; $ci++) {
  $cat = $categories[$ci]
  [void]$sb.AppendLine("    {")
  [void]$sb.AppendLine(('      "id": "' + (Json-Escape $cat.id) + '",'))
  [void]$sb.AppendLine(('      "name": "' + (Json-Escape $cat.name) + '",'))
  [void]$sb.AppendLine('      "sections": [')
  $secs = @($cat.sections)
  for ($si = 0; $si -lt $secs.Count; $si++) {
    $sec = $secs[$si]
    [void]$sb.AppendLine("        {")
    [void]$sb.AppendLine(('          "name": "' + (Json-Escape ([string]$sec.name)) + '",'))
    if ($sec.note) {
      [void]$sb.AppendLine(('          "note": "' + (Json-Escape ([string]$sec.note)) + '",'))
    }
    [void]$sb.AppendLine('          "items": [')
    $items = @($sec.items)
    for ($ii = 0; $ii -lt $items.Count; $ii++) {
      $it = $items[$ii]
      $comma = ","
      if ($ii -eq $items.Count - 1) { $comma = "" }
      $priceJson = "null"
      if ($null -ne $it.price) { $priceJson = [string]$it.price }
      $img = ""
      if ($it.image) { $img = [string]$it.image }
      $line = '            {"name":"' + (Json-Escape ([string]$it.name)) + '","article":"' + (Json-Escape ([string]$it.article)) + '","price":' + $priceJson + ',"image":"' + (Json-Escape $img) + '"'
      if ($it.color) { $line += ',"color":"' + (Json-Escape ([string]$it.color)) + '"' }
      if ($it.material) { $line += ',"material":"' + (Json-Escape ([string]$it.material)) + '"' }
      $line += "}" + $comma
      [void]$sb.AppendLine($line)
    }
    [void]$sb.AppendLine("          ]")
    $scomma = ","
    if ($si -eq $secs.Count - 1) { $scomma = "" }
    [void]$sb.AppendLine("        }" + $scomma)
  }
  [void]$sb.AppendLine("      ]")
  $ccomma = ","
  if ($ci -eq $categories.Count - 1) { $ccomma = "" }
  [void]$sb.AppendLine("    }" + $ccomma)
}
[void]$sb.AppendLine("  ]")
[void]$sb.AppendLine("}")

$jsonPath = Join-Path $outData "catalog.json"
[IO.File]::WriteAllText($jsonPath, $sb.ToString(), [Text.UTF8Encoding]::new($false))
Write-Host "Wrote $jsonPath"
Write-Host "Images: $imgCounter"
Get-ChildItem $outImages | Measure-Object -Property Length -Sum | ForEach-Object { Write-Host ("Images size MB: {0:N1}" -f ($_.Sum / 1MB)) }
