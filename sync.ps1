# 벤츠ACC 랜딩페이지 - 구글시트 자동 동기화 스크립트
# 매일 실행되어 products.json / images / index.html의 상품 데이터를 최신 시트 내용으로 갱신합니다.
# 실패 시 기존 페이지 파일은 건드리지 않고 로그만 남깁니다.

$ErrorActionPreference = "Stop"

$SHEET_ID   = "1BSnR47aHPb5kLnHhjIK3ry7qVlBKWmDtXQHy3ME0qYY"
$SHEET_NAME = "컬렉션목록"
$ProjDir    = "c:\Users\ADMIN\Desktop\Js\클로드코워크\벤츠ACC"
$ImgDir     = Join-Path $ProjDir "images"
$TmpDir     = Join-Path $env:TEMP "acc_sync_tmp"
$LogFile    = Join-Path $ProjDir "sync_log.txt"

function Write-Log($msg) {
    $line = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $msg
    Add-Content -LiteralPath $LogFile -Value $line -Encoding utf8
    Write-Host $line
}

function Get-CellValue($cellNode, $ns, $sharedStrings) {
    if (-not $cellNode) { return $null }
    $t = $cellNode.GetAttribute("t")
    $v = $cellNode.SelectSingleNode("s:v", $ns)
    if (-not $v) { return $null }
    if ($t -eq "s") { return $sharedStrings[[int]$v.InnerText] }
    return $v.InnerText
}

function Read-ZipEntryText($zip, $entryName) {
    $entry = $zip.GetEntry($entryName)
    if (-not $entry) { return $null }
    $stream = $entry.Open()
    $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
    $text = $reader.ReadToEnd()
    $reader.Close(); $stream.Close()
    return $text
}

try {
    New-Item -ItemType Directory -Force -Path $TmpDir | Out-Null
    Write-Log "동기화 시작"

    # 1) CSV (텍스트 필드: 부품번호/부품명/수량/가격/카테고리) 다운로드
    $csvUrl  = "https://docs.google.com/spreadsheets/d/$SHEET_ID/gviz/tq?tqx=out:csv&sheet=$([uri]::EscapeDataString($SHEET_NAME))"
    $csvPath = Join-Path $TmpDir "collection_list.csv"
    Invoke-WebRequest -Uri $csvUrl -OutFile $csvPath -UseBasicParsing

    $rows = Import-Csv -LiteralPath $csvPath
    if ($rows.Count -lt 100) { throw "CSV 데이터가 비정상적으로 적습니다 ($($rows.Count)행) - 시트 접근 실패 의심" }

    $textByPart = @{}
    foreach ($r in $rows) {
        $part = $r.PartNUMBER
        if (-not $part) { continue }
        $qty   = [int]($r.Quantity -replace '[^0-9\-]', '')
        $price = [int]((($r.'판매가 vat포함') -replace '[^0-9]', ''))
        $textByPart[$part] = [PSCustomObject]@{
            part     = $part
            name     = $r.Description.Trim()
            qty      = $qty
            price    = $price
            category = $r.'종류'.Trim()
        }
    }
    Write-Log "CSV 파싱 완료: $($textByPart.Count)개 상품 텍스트 데이터"

    # 2) XLSX (이미지 매칭용) 다운로드
    $xlsxUrl  = "https://docs.google.com/spreadsheets/d/$SHEET_ID/export?format=xlsx"
    $xlsxPath = Join-Path $TmpDir "sheet.xlsx"
    Invoke-WebRequest -Uri $xlsxUrl -OutFile $xlsxPath -UseBasicParsing

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($xlsxPath)

    [xml]$ssXml = Read-ZipEntryText $zip "xl/sharedStrings.xml"
    $ns = New-Object System.Xml.XmlNamespaceManager($ssXml.NameTable)
    $ns.AddNamespace("s", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
    $siNodes = $ssXml.SelectNodes("//s:si", $ns)
    $sharedStrings = New-Object string[] $siNodes.Count
    for ($i = 0; $i -lt $siNodes.Count; $i++) {
        $texts = $siNodes[$i].SelectNodes(".//s:t", $ns) | ForEach-Object { $_.InnerText }
        $sharedStrings[$i] = ($texts -join "")
    }

    # workbook.xml -> sheet name -> rId -> sheetN.xml
    [xml]$wbXml = Read-ZipEntryText $zip "xl/workbook.xml"
    $nsWb = New-Object System.Xml.XmlNamespaceManager($wbXml.NameTable)
    $nsWb.AddNamespace("s", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
    $nsWb.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
    $sheetNode = $wbXml.SelectSingleNode("//s:sheets/s:sheet[@name='$SHEET_NAME']", $nsWb)
    if (-not $sheetNode) { throw "워크북에서 '$SHEET_NAME' 시트를 찾을 수 없습니다" }
    $rId = $sheetNode.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")

    [xml]$wbRels = Read-ZipEntryText $zip "xl/_rels/workbook.xml.rels"
    $relTarget = ($wbRels.Relationships.Relationship | Where-Object { $_.Id -eq $rId }).Target
    $sheetFile = "xl/" + $relTarget

    [xml]$sheetXml = Read-ZipEntryText $zip $sheetFile
    $rowsXml = $sheetXml.SelectNodes("//s:sheetData/s:row", $ns)

    $partToRow = @{}
    foreach ($row in $rowsXml) {
        $rNum = [int]$row.GetAttribute("r")
        if ($rNum -eq 1) { continue }
        if ($row.GetAttribute("hidden") -eq "1") { continue }
        $aCell = $row.SelectSingleNode("s:c[@r='A$rNum']", $ns)
        $aVal = Get-CellValue $aCell $ns $sharedStrings
        if (-not $aVal) { continue }
        $partToRow[$aVal] = $rNum
    }
    Write-Log "시트 노출 행: $($partToRow.Count)개"

    # 해당 시트의 drawing 관계 찾기
    $sheetBaseName = [System.IO.Path]::GetFileName($sheetFile)
    $sheetRelsPath = "xl/worksheets/_rels/$sheetBaseName.rels"
    [xml]$sheetRels = Read-ZipEntryText $zip $sheetRelsPath
    $drawingRel = $sheetRels.Relationships.Relationship | Where-Object { $_.Type -like "*drawing*" }
    $drawingFile = "xl/" + ($drawingRel.Target -replace '^\.\./', '')

    [xml]$drawingXml = Read-ZipEntryText $zip $drawingFile
    $drawingDir = [System.IO.Path]::GetDirectoryName($drawingFile) -replace '\\', '/'
    $drawingBaseName = [System.IO.Path]::GetFileName($drawingFile)
    [xml]$drawingRels = Read-ZipEntryText $zip "$drawingDir/_rels/$drawingBaseName.rels"
    $relMap = @{}
    foreach ($rel in $drawingRels.Relationships.Relationship) { $relMap[$rel.Id] = $rel.Target }

    $nsD = New-Object System.Xml.XmlNamespaceManager($drawingXml.NameTable)
    $nsD.AddNamespace("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")
    $nsD.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")
    $anchors = $drawingXml.SelectNodes("//xdr:oneCellAnchor", $nsD)
    $xdrRowToImage = @{}
    foreach ($a in $anchors) {
        $xrow = [int]$a.SelectSingleNode("xdr:from/xdr:row", $nsD).InnerText
        $embedId = $a.SelectSingleNode(".//a:blip", $nsD).GetAttribute("embed", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
        $xdrRowToImage[$xrow] = $relMap[$embedId]
    }
    Write-Log "이미지 앵커: $($xdrRowToImage.Count)개"

    $mediaDirInZip = "xl/media"

    # 3) 텍스트 + 이미지 결합, 새 이미지 파일로 추출
    New-Item -ItemType Directory -Force -Path $ImgDir | Out-Null
    $newProducts = @()
    $usedImageFiles = New-Object System.Collections.Generic.HashSet[string]
    $matched = 0

    foreach ($part in $textByPart.Keys) {
        $t = $textByPart[$part]
        $imgField = $null
        if ($partToRow.ContainsKey($part)) {
            $sheetRow = $partToRow[$part]
            $xdrRow = $sheetRow - 1
            if ($xdrRowToImage.ContainsKey($xdrRow)) {
                $mediaRel = $xdrRowToImage[$xdrRow] -replace '^\.\./', ''
                $mediaEntryName = if ($mediaRel -match '^xl/') { $mediaRel } else { "xl/$mediaRel" }
                $mediaEntry = $zip.GetEntry($mediaEntryName)
                if ($mediaEntry) {
                    $ext = [System.IO.Path]::GetExtension($mediaEntry.Name)
                    $safePart = ($part -replace '[^A-Za-z0-9]+', '-').Trim('-')
                    $outName = "$safePart$ext"
                    $outPath = Join-Path $ImgDir $outName
                    $es = $mediaEntry.Open()
                    $fs = [System.IO.File]::Create($outPath)
                    $es.CopyTo($fs)
                    $fs.Close(); $es.Close()
                    $imgField = "images/$outName"
                    [void]$usedImageFiles.Add($outName)
                    $matched++
                }
            }
        }
        $newProducts += [PSCustomObject]@{
            part     = $t.part
            name     = $t.name
            qty      = $t.qty
            price    = $t.price
            category = $t.category
            image    = $imgField
        }
    }
    $zip.Dispose()
    Write-Log "이미지 매칭: $matched / $($newProducts.Count)"

    # 4) 더 이상 쓰이지 않는 이미지 파일 정리
    $existingFiles = Get-ChildItem -LiteralPath $ImgDir -File
    $removed = 0
    foreach ($f in $existingFiles) {
        if (-not $usedImageFiles.Contains($f.Name)) {
            Remove-Item -LiteralPath $f.FullName -Force
            $removed++
        }
    }
    if ($removed -gt 0) { Write-Log "미사용 이미지 $removed 개 정리" }

    # 5) products.json 저장
    $newProducts = $newProducts | Sort-Object { [array]::IndexOf(($rows.PartNUMBER), $_.part) }
    $productsJsonPath = Join-Path $ProjDir "products.json"
    $newProducts | ConvertTo-Json -Depth 3 | Out-File -LiteralPath $productsJsonPath -Encoding utf8

    # 6) index.html 갱신 (PRODUCTS 배열 + 동기화 배지)
    $indexPath = Join-Path $ProjDir "index.html"
    $htmlLines = Get-Content -LiteralPath $indexPath -Encoding UTF8
    $compactJson = ($newProducts | ConvertTo-Json -Depth 3 -Compress)

    $productsLineIdx = -1
    for ($i = 0; $i -lt $htmlLines.Count; $i++) {
        if ($htmlLines[$i] -match 'const\s+PRODUCTS\s*=') { $productsLineIdx = $i; break }
    }
    if ($productsLineIdx -lt 0) { throw "index.html에서 'const PRODUCTS =' 라인을 찾을 수 없습니다" }
    $htmlLines[$productsLineIdx] = "const PRODUCTS = $compactJson;"

    $syncTimeText = Get-Date -Format "yyyy-MM-dd HH:mm"
    for ($i = 0; $i -lt $htmlLines.Count; $i++) {
        if ($htmlLines[$i] -match '마지막 동기화') {
            $htmlLines[$i] = $htmlLines[$i] -replace '마지막 동기화 \d{4}-\d{2}-\d{2} \d{2}:\d{2}', "마지막 동기화 $syncTimeText"
        }
    }

    Set-Content -LiteralPath $indexPath -Value $htmlLines -Encoding utf8
    Write-Log "index.html 갱신 완료 (동기화 시각: $syncTimeText)"
    Write-Log "동기화 성공: 총 $($newProducts.Count)개 상품, 이미지 $matched 개"

    # 7) GitHub Pages(모바일용 실제 배포 사이트)에 반영
    try {
        Push-Location $ProjDir
        $env:PATH += ";C:\Program Files\GitHub CLI"
        git add -A 2>&1 | Out-Null
        $changed = git status --porcelain
        if ($changed) {
            git commit -m "자동 동기화: $syncTimeText" 2>&1 | Out-Null
            git push origin master 2>&1 | Out-Null
            Write-Log "GitHub Pages에 배포 완료"
        } else {
            Write-Log "GitHub에 반영할 변경사항 없음"
        }
        Pop-Location
    } catch {
        Write-Log "GitHub 배포 실패(로컬 파일은 정상 갱신됨): $($_.Exception.Message)"
        Pop-Location
    }
}
catch {
    Write-Log "동기화 실패: $($_.Exception.Message)"
    exit 1
}
finally {
    if (Test-Path $TmpDir) { Remove-Item -LiteralPath $TmpDir -Recurse -Force -ErrorAction SilentlyContinue }
}
