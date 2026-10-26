# Builds data/catalog.json and compressed images from the Excel price list.
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Drawing

$src = "C:\Users\clost\Downloads\Новый прайс (2).xlsx"
$root = Split-Path -Parent $PSScriptRoot
$imgDir = Join-Path $root "images"
$dataDir = Join-Path $root "data"
New-Item -ItemType Directory -Force -Path $imgDir | Out-Null
New-Item -ItemType Directory -Force -Path $dataDir | Out-Null
Get-ChildItem $imgDir -File -ErrorAction SilentlyContinue | Remove-Item -Force

$nsMain = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
$nsRel = "http://schemas.openxmlformats.org/package/2006/relationships"
$nsDraw = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
$nsA = "http://schemas.openxmlformats.org/drawingml/2006/main"
$nsR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

function Get-EntryText($zip, $name) {
  $e = $zip.GetEntry($name)
  if (-not $e) { return $null }
  $sr = New-Object System.IO.StreamReader($e.Open())
  $t = $sr.ReadToEnd()
  $sr.Close()
  return $t
}

function Get-Xml($zip, $name) {
  $t = Get-EntryText $zip $name
  if (-not $t) { return $null }
  $xml = New-Object System.Xml.XmlDocument
  $xml.PreserveWhitespace = $false
  $xml.LoadXml($t)
  return $xml
}

function Get-SharedStrings($zip) {
  $xml = Get-Xml $zip "xl/sharedStrings.xml"
  $list = New-Object System.Collections.Generic.List[string]
  foreach ($si in $xml.SelectNodes("//*[local-name()='si']")) {
    $texts = $si.SelectNodes(".//*[local-name()='t']")
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($t in $texts) { $parts.Add([string]$t.InnerText) }
    $list.Add(($parts -join ""))
  }
  return $list
}

function Get-ChildByName($node, $name) {
  foreach ($c in $node.ChildNodes) {
    if ($c.LocalName -eq $name) { return $c }
  }
  return $null
}

function Get-CellValue($cell, $strings) {
  $t = $cell.GetAttribute("t")
  if ($t -eq "s") {
    $v = Get-ChildByName $cell "v"
    if ($v) { return [string]$strings[[int]$v.InnerText] }
  }
  elseif ($t -eq "inlineStr") {
    $is = Get-ChildByName $cell "is"
    if ($is) { return [string]$is.InnerText }
  }
  elseif ($t -eq "str") {
    $v = Get-ChildByName $cell "v"
    if ($v) { return [string]$v.InnerText }
  }
  else {
    $v = Get-ChildByName $cell "v"
    if ($v) { return [string]$v.InnerText }
  }
  return ""
}

function Get-ColLetter($ref) {
  return ($ref -replace "[0-9]", "")
}

function Clean-Text($s) {
  if ($null -eq $s) { return "" }
  $s = [string]$s
  $s = $s -replace [char]160, " "
  $s = $s.Trim()
  $s = $s -replace "\s+", " "
  return $s
}

function Parse-Price($raw) {
  $raw = Clean-Text $raw
  if (-not $raw) { return $null }
  if ($raw -notmatch "^[0-9]+([.,][0-9]+)?$") { return $null }
  return [decimal]($raw -replace ",", ".")
}

function Slug($s) {
  $s = Clean-Text $s
  $s = $s.ToLowerInvariant()
  $s = $s -replace "[^a-z0-9а-яё]+", "-"
  $s = $s.Trim("-")
  if (-not $s) { return "cat" }
  return $s
}

function Get-DrawingMap($zip, $sheetFile) {
  $map = @{}
  $relsName = $sheetFile -replace "worksheets/", "worksheets/_rels/" -replace "\.xml$", ".xml.rels"
  $relsXml = Get-Xml $zip $relsName
  if (-not $relsXml) { return $map }
  $drawingRel = $null
  foreach ($rel in $relsXml.SelectNodes("//*[local-name()='Relationship']")) {
    if ($rel.Type -like "*drawing") { $drawingRel = $rel; break }
  }
  if (-not $drawingRel) { return $map }

  $target = $drawingRel.Target.Replace("\", "/")
  $drawingPath = "xl/drawings/" + ($target -replace "^\.\./drawings/", "")
  $drawingXml = Get-Xml $zip $drawingPath
  if (-not $drawingXml) { return $map }

  $drawRelsPath = $drawingPath -replace "drawings/", "drawings/_rels/" -replace "\.xml$", ".xml.rels"
  $drawRels = Get-Xml $zip $drawRelsPath
  $embedToMedia = @{}
  if ($drawRels) {
    foreach ($rel in $drawRels.SelectNodes("//*[local-name()='Relationship']")) {
      $media = ($rel.Target -replace "^\.\./media/", "")
      $embedToMedia[$rel.Id] = $media
    }
  }

  foreach ($anchor in $drawingXml.SelectNodes("//*[local-name()='twoCellAnchor' or local-name()='oneCellAnchor']")) {
    $from = $null
    foreach ($ch in $anchor.ChildNodes) { if ($ch.LocalName -eq "from") { $from = $ch; break } }
    if (-not $from) { continue }
    $rowNode = $null
    foreach ($ch in $from.ChildNodes) { if ($ch.LocalName -eq "row") { $rowNode = $ch; break } }
    if (-not $rowNode) { continue }
    $excelRow = [int]$rowNode.InnerText + 1
    $embed = $anchor.SelectSingleNode(".//*[local-name()='blip']")
    if (-not $embed) { continue }
    $rid = $embed.GetAttribute("embed")
    if (-not $rid -and $embed.Attributes) {
      foreach ($attr in $embed.Attributes) {
        if ($attr.LocalName -eq "embed") { $rid = $attr.Value; break }
      }
    }
    if (-not $rid -or -not $embedToMedia.ContainsKey($rid)) { continue }
    if (-not $map.ContainsKey($excelRow)) {
      $map[$excelRow] = $embedToMedia[$rid]
    }
  }
  return $map
}

function Save-ResizedImage($zip, $mediaName, $outPath) {
  if (Test-Path $outPath) { return }
  $entry = $zip.GetEntry("xl/media/$mediaName")
  if (-not $entry) { return }
  $tmp = Join-Path $env:TEMP ("xlsx-" + [Guid]::NewGuid().ToString() + [IO.Path]::GetExtension($mediaName))
  $inStream = $entry.Open()
  $fs = [IO.File]::Create($tmp)
  $inStream.CopyTo($fs)
  $fs.Close()
  $inStream.Close()
  try {
    $img = [System.Drawing.Image]::FromFile($tmp)
    $max = 420
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
    $img.Dispose()
    $enc = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq "image/jpeg" }
    $ep = New-Object System.Drawing.Imaging.EncoderParameters 1
    $ep.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter ([System.Drawing.Imaging.Encoder]::Quality, [int64]72)
    $bmp.Save($outPath, $enc, $ep)
    $bmp.Dispose()
  }
  catch {
    Copy-Item $tmp $outPath -Force
  }
  finally {
    Remove-Item $tmp -Force -ErrorAction SilentlyContinue
  }
}

function Get-Rows($zip, $sheetFile, $strings) {
  $xml = Get-Xml $zip $sheetFile
  $rows = @()
  foreach ($row in $xml.SelectNodes("//*[local-name()='sheetData']/*[local-name()='row']")) {
    $idx = [int]$row.r
    $cells = @{}
    foreach ($c in $row.ChildNodes) {
      if ($c.LocalName -ne "c") { continue }
      $col = Get-ColLetter $c.r
      $cells[$col] = Get-CellValue $c $strings
    }
    $rows += [pscustomobject]@{ Index = $idx; Cells = $cells }
  }
  return $rows
}

function Is-HeaderRow($a, $b, $c, $d, $e) {
  $joined = ("$a $b $c $d $e").ToLowerInvariant()
  if ($joined -match "фото" -and $joined -match "артикул") { return $true }
  if ($a -eq "Артикул" -and $c -eq "Материал") { return $true }
  return $false
}

$zip = [System.IO.Compression.ZipFile]::OpenRead($src)
$strings = Get-SharedStrings $zip

$categories = @(
  @{ id = "petli"; title = "Петли"; file = "xl/worksheets/sheet1.xml"; kind = "standard" },
  @{ id = "razdvizhki"; title = "Раздввижки"; file = "xl/worksheets/sheet2.xml"; kind = "standard" },
  @{ id = "connectors"; title = "Коннектора"; file = "xl/worksheets/sheet3.xml"; kind = "standard" },
  @{ id = "soediniteli"; title = "Соеденители для труб"; file = "xl/worksheets/sheet4.xml"; kind = "connectors" },
  @{ id = "truby"; title = "Трубы-П Профиль"; file = "xl/worksheets/sheet5.xml"; kind = "standard" },
  @{ id = "napravlyayushchie"; title = "Направляющие нижние"; file = "xl/worksheets/sheet6.xml"; kind = "standard" },
  @{ id = "ruchki"; title = "Ручки"; file = "xl/worksheets/sheet7.xml"; kind = "standard" },
  @{ id = "uplotniteli"; title = "Уплотнительный профиль"; file = "xl/worksheets/sheet8.xml"; kind = "standard" },
  @{ id = "roliki"; title = "Ролики"; file = "xl/worksheets/sheet10.xml"; kind = "standard" }
)

$usedMedia = @{}
$outCats = New-Object System.Collections.ArrayList
$script:itemSeq = 0

foreach ($cat in $categories) {
  Write-Host "Parsing $($cat.title)..."
  $rows = Get-Rows $zip $cat.file $strings
  $drawMap = Get-DrawingMap $zip $cat.file
  $sections = New-Object System.Collections.ArrayList
  $currentItems = New-Object System.Collections.ArrayList
  $currentTitle = ""
  $currentNote = ""
  $currentId = $cat.id
  $carryArticle = ""
  $carryMaterial = ""
  $carryPhotoMedia = ""

  foreach ($row in $rows) {
    $A = Clean-Text $row.Cells["A"]
    $B = Clean-Text $row.Cells["B"]
    $C = Clean-Text $row.Cells["C"]
    $D = Clean-Text $row.Cells["D"]
    $E = Clean-Text $row.Cells["E"]
    $I = Clean-Text $row.Cells["I"]

    if (Is-HeaderRow $A $B $C $D $E) { continue }
    if ($A -eq "#VALUE!") { $A = "" }

    $rowPhotoMedia = $null
    if ($drawMap.ContainsKey($row.Index)) { $rowPhotoMedia = $drawMap[$row.Index] }

    if ($cat.kind -eq "connectors") {
      if ($A -and -not $C -and -not $D -and -not (Parse-Price $E)) {
        if ($currentItems.Count -gt 0 -or $currentTitle) {
          $title = $currentTitle
          if (-not $title) { $title = $cat.title }
          [void]$sections.Add([ordered]@{
            id = $currentId
            title = $title
            note = $currentNote
            items = @($currentItems.ToArray())
          })
          $currentItems.Clear()
        }
        $currentTitle = $A
        $currentId = $cat.id + "-" + (Slug $A)
        $currentNote = ""
        $carryArticle = ""; $carryMaterial = ""; $carryPhotoMedia = ""
        continue
      }
      if ($A) { $carryArticle = $A }
      if ($C) { $carryMaterial = $C }
      if ($rowPhotoMedia) { $carryPhotoMedia = $rowPhotoMedia }
      $color = $D
      $price = Parse-Price $E
      $article = $carryArticle
      if (-not $article -and -not $color) { continue }
      if (-not $color -and -not $price) { continue }
      $script:itemSeq++
      $photo = $null
      if ($carryPhotoMedia) {
        $photoName = ("p{0:D4}.jpg" -f $script:itemSeq)
        $usedMedia[$photoName] = $carryPhotoMedia
        $photo = "images/$photoName"
      }
      $nameParts = @()
      if ($article) { $nameParts += $article }
      if ($color) { $nameParts += $color }
      $name = ($nameParts -join ", ")
      $sitePrice = $null
      if ($null -ne $price) { $sitePrice = [int][math]::Round($price + 1000) }
      [void]$currentItems.Add([ordered]@{
        name = $name
        article = $article
        color = $color
        material = $carryMaterial
        price = $sitePrice
        photo = $photo
      })
      continue
    }

    if ($A -match "^Цены актуальны") {
      $currentNote = $A
      continue
    }

    $looksLikeSection = $false
    if ($A -and -not $C -and -not (Parse-Price $D) -and -not (Parse-Price $B)) { $looksLikeSection = $true }
    if ($looksLikeSection) {
      if ($currentItems.Count -gt 0 -or $currentTitle) {
        $title = $currentTitle
        if (-not $title) { $title = $cat.title }
        [void]$sections.Add([ordered]@{
          id = $currentId
          title = $title
          note = $currentNote
          items = @($currentItems.ToArray())
        })
        $currentItems.Clear()
      }
      $currentTitle = $A
      $currentId = $cat.id + "-" + (Slug $A)
      $currentNote = ""
      continue
    }

    $name = $B
    $article = $C
    $price = Parse-Price $D
    if (-not $name -and $article) { $name = $article }
    if (-not $name -and -not $article) { continue }
    if ($name -in @("Описание", "Описание/Ссылка")) { continue }

    $script:itemSeq++
    $photo = $null
    if ($rowPhotoMedia) {
      $photoName = ("p{0:D4}.jpg" -f $script:itemSeq)
      $usedMedia[$photoName] = $rowPhotoMedia
      $photo = "images/$photoName"
    }
    $sitePrice = $null
    if ($null -ne $price) { $sitePrice = [int][math]::Round($price + 1000) }
    [void]$currentItems.Add([ordered]@{
      name = $name
      article = $article
      color = ""
      material = ""
      note = $I
      price = $sitePrice
      photo = $photo
    })
  }

  if ($currentItems.Count -gt 0 -or $currentTitle) {
    $title = $currentTitle
    if (-not $title) { $title = $cat.title }
    [void]$sections.Add([ordered]@{
      id = $currentId
      title = $title
      note = $currentNote
      items = @($currentItems.ToArray())
    })
  }

  $count = 0
  foreach ($s in $sections) { $count += $s.items.Count }
  Write-Host "  sections=$($sections.Count) items=$count"
  [void]$outCats.Add([ordered]@{
    id = $cat.id
    title = $cat.title
    sections = @($sections)
  })
}

Write-Host "Exporting $($usedMedia.Count) images..."
foreach ($kv in $usedMedia.GetEnumerator()) {
  $outPath = Join-Path $imgDir $kv.Key
  Save-ResizedImage $zip $kv.Value $outPath
}

$zip.Dispose()

function Get-PhotoFamily($article) {
  $a = ([string]$article).ToUpperInvariant()
  $a = $a -replace "\s+", ""
  $a = $a -replace "/.*$", ""
  $a = $a -replace "[СC]R$", ""
  $a = $a -replace "(BLK|BL|STN|GFM|BGD|GD|BRZ|WM|PC|MC|GF|GDM|GDB|GRF|PSS|KW|BKL)$", ""
  $a = $a -replace "[-_]+$", ""
  return $a
}

function Fill-SectionPhotos($items) {
  $last = $null
  foreach ($it in $items) {
    if ($it.photo) { $last = $it.photo }
    elseif ($last) { $it.photo = $last }
  }
  $next = $null
  for ($i = $items.Count - 1; $i -ge 0; $i--) {
    $it = $items[$i]
    if ($it.photo) { $next = $it.photo }
    elseif ($next) { $it.photo = $next }
  }
}

foreach ($cat in $outCats) {
  foreach ($sec in $cat.sections) {
    Fill-SectionPhotos $sec.items
  }
}

$familyPhoto = @{}
$slidePhoto = $null
$anyPhoto = $null
foreach ($cat in $outCats) {
  foreach ($sec in $cat.sections) {
    foreach ($it in $sec.items) {
      if (-not $it.photo) { continue }
      if (-not $anyPhoto) { $anyPhoto = $it.photo }
      $nm = ([string]$it.name).ToLowerInvariant()
      if (-not $slidePhoto -and $nm -match "раздвиж") { $slidePhoto = $it.photo }
      $fam = Get-PhotoFamily $it.article
      if ($fam -and -not $familyPhoto.ContainsKey($fam)) { $familyPhoto[$fam] = $it.photo }
    }
  }
}

foreach ($cat in $outCats) {
  $catPhoto = $null
  foreach ($sec in $cat.sections) {
    foreach ($it in $sec.items) { if ($it.photo) { $catPhoto = $it.photo; break } }
    if ($catPhoto) { break }
  }
  foreach ($sec in $cat.sections) {
    foreach ($it in $sec.items) {
      if ($it.photo) { continue }
      $fam = Get-PhotoFamily $it.article
      $nm = ([string]$it.name).ToLowerInvariant()
      $art = ([string]$it.article).ToUpperInvariant()
      if ($fam -and $familyPhoto.ContainsKey($fam)) { $it.photo = $familyPhoto[$fam] }
      elseif ($slidePhoto -and ($nm -match "раздвиж" -or $art -match "^RS")) { $it.photo = $slidePhoto }
      elseif ($catPhoto) { $it.photo = $catPhoto }
      elseif ($anyPhoto) { $it.photo = $anyPhoto }
    }
  }
}

$catalog = [ordered]@{
  title = "Прайс-лист"
  priceMarkup = 1000
  categories = @($outCats)
}
$json = $catalog | ConvertTo-Json -Depth 10 -Compress:$false
[IO.File]::WriteAllText((Join-Path $dataDir "catalog.json"), $json, [Text.UTF8Encoding]::new($false))
$jsDir = Join-Path $root "js"
New-Item -ItemType Directory -Force -Path $jsDir | Out-Null
$js = "window.CATALOG = " + $json + ";`n"
[IO.File]::WriteAllText((Join-Path $jsDir "catalog-data.js"), $js, [Text.UTF8Encoding]::new($false))
Write-Host "Done. Items total=$($script:itemSeq)"
