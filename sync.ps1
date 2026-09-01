# 벤츠ACC 랜딩페이지 - 구글시트 자동 동기화 스크립트
# 매일 17:00에 실행되어 products.json / images / index.html의 상품 데이터를 최신 시트 내용으로 갱신합니다.
# 실패 시 기존 페이지 파일은 건드리지 않고 로그만 남깁니다.
# 시트에서 파싱된 상품 데이터가 비정상적으로 적게 나오는 경우(원인 불문) 재시도 예약 파일을
# 남기고, 18:00에 BenzACC_RetrySync 작업이 -RetryCheck로 이 스크립트를 다시 호출한다.
# 재시도 예약이 없으면 -RetryCheck 호출은 네트워크를 건드리지 않고 즉시 종료한다.

param([switch]$RetryCheck)

$ErrorActionPreference = "Stop"

$SHEET_ID   = "1BSnR47aHPb5kLnHhjIK3ry7qVlBKWmDtXQHy3ME0qYY"
$SHEET_NAME = "컬렉션목록"
$ProjDir    = "c:\Users\ADMIN\Desktop\Js\클로드코워크\벤츠ACC"
$ImgDir     = Join-Path $ProjDir "images"
$TmpDir     = Join-Path $env:TEMP "acc_sync_tmp"
$LogFile    = Join-Path $ProjDir "sync_log.txt"
$GiftPath   = Join-Path $ProjDir "gift.html"
$RetryMarkerPath = Join-Path $ProjDir ".sync_retry_needed"

function Write-Log($msg) {
    $line = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $msg
    Add-Content -LiteralPath $LogFile -Value $line -Encoding utf8
    Write-Host $line
}

if ($RetryCheck) {
    if (-not (Test-Path $RetryMarkerPath)) {
        Write-Log "18:00 재시도 확인: 재시도 예약 없음 (17:00 동기화 정상) - 종료"
        exit 0
    }
    Write-Log "18:00 재시도 확인: 재시도 예약 발견 - 재동기화 시작"
    Remove-Item -LiteralPath $RetryMarkerPath -Force
}

# 5 STAR GIFT 증정 후보 부품번호 (구글시트 "5star A&C LIST" 탭에서 큐레이션, 2026-08-12 기준).
# 이 목록 자체는 매일 바뀌지 않음 - 실제 증정 페이지 노출 여부는 아래에서
# "메인 컬렉션(products.json)에 남아있고 재고(qty)가 0보다 큰가"로 매번 재판정한다.
# 메인 컬렉션에서 품목이 사라지거나 재고가 0이 되면 증정 페이지에서도 자동으로 사라진다.
$GIFT_PART_NUMBERS = @(
    "MB6 6 04 1523","MB6 6 04 1524","MB6 6 95 3744","MB6 6 95 3307","MB6 6 95 9846","MB6 7 87 1175","MB6 6 95 9725",
    "MB6 6 95 9728","MB6 6 95 9376","MB6 6 95 9377","MB6 7 99 7018","MB6 6 95 9352","MB6 6 95 9341","MB6 6 95 9821",
    "MB6 6 96 0576","MB6 6 95 9924","MB6 6 95 8404","MB6 6 95 3714","MB6 6 95 3713","MB6 6 95 3358","MB6 6 95 9715",
    "MB6 7 87 2867","MB6 6 04 2026","MB6 6 95 5083","MB6 6 95 5015","MQALKRBV9591501","MB6 6 95 9654","MA177 810 81 03",
    "MB6 6 45 0405","MQALKRBN9782401","MQALKRBN9782801","MQALKRBM1390201","MQALKRBN9786201","MNBU3125 601","MQALKRBU2643201",
    "MB6 6 95 5770","MB6 6 95 9442","MQALKRB81201134","MNBV9403 101","MNBV9403 001","MB6 6 45 0461","MB6 6 95 4769",
    "MB6 7 87 2008","MB6 6 95 9385","MB6 6 95 3636/64","MB6 6 47 2005","MB6 6 95 9925","MB6 6 04 1558","MB6 6 95 4774",
    "MQALKRB81202037","MQALKRB81202039","MQALKRB81202041","MB6 6 95 3230","MB6 6 95 4738","MB6 6 95 3148","MB6 6 95 3409",
    "MB6 7 87 1618","MB6 6 04 1563","MQALKRB81201164","MB6 6 95 3317","MB6 6 95 3318","MQALKRBB1613101","MB6 6 95 9789",
    "MQALKRB81201188","MQALKRB81201163"
)

# 증정 전용 재고 - 메인 컬렉션 리스트(index.html)에 노출하지 않을 부품번호.
# 시트에는 그대로 남아있고 재고 수량도 정상 반영되지만, 판매용 리스트에는 나오지 않는다.
# 우산 MB UMBRELLA 133CM / MB UMBREALLA 105 CM (2026-08-19 기준, 증정용으로 지정).
$MAIN_LIST_EXCLUDE_PART_NUMBERS = @(
    "MQALKRQ10023067","MQALKRQ10023066"
)

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

    # 1) XLSX 다운로드 - 텍스트 데이터(부품번호/부품명/수량/가격/카테고리)와 이미지 매칭을
    # 하나의 소스에서 함께 처리한다.
    # 주의: gviz CSV 내보내기는 시트에 걸린 실시간 공유 필터(데이터 > 필터 만들기) 상태를
    # 그대로 반영해서, 누군가 시트를 필터링해둔 동안에는 필터링된 일부 행만 내려온다
    # (2026-08-20 확인: 우산 재고 반영 중 "키링" 필터가 걸려있어 305개 중 33개만 수신됨).
    # 반면 XLSX 내보내기는 필터 상태와 무관하게 항상 전체 데이터를 담고 있어, 필터를
    # 해제하지 않아도 안정적으로 동기화할 수 있다.
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

    # 헤더 행(1행)에서 컬럼명 -> 열 문자 매핑 (컬럼 순서가 바뀌어도 안전하게 동작)
    $headerRow = $rowsXml | Where-Object { $_.GetAttribute("r") -eq "1" } | Select-Object -First 1
    if (-not $headerRow) { throw "시트에서 헤더 행(1행)을 찾을 수 없습니다" }
    $colByHeader = @{}
    foreach ($c in $headerRow.SelectNodes("s:c", $ns)) {
        $colLetter = ($c.GetAttribute("r") -replace '[0-9]+$', '')
        $headerVal = Get-CellValue $c $ns $sharedStrings
        if ($headerVal) { $colByHeader[$headerVal.Trim()] = $colLetter }
    }
    # 헤더 이름이 바뀌어도(예: "종류" -> "구분") 깨지지 않도록 후보 이름 중 하나만 있으면 되게 처리.
    $HEADER_ALIASES = @{
        PART     = @("PartNUMBER")
        NAME     = @("Description")
        QTY      = @("Quantity")
        PRICE    = @("판매가 vat포함")
        CATEGORY = @("종류", "구분")
    }
    $col = @{}
    foreach ($key in $HEADER_ALIASES.Keys) {
        $found = $HEADER_ALIASES[$key] | Where-Object { $colByHeader.ContainsKey($_) } | Select-Object -First 1
        if (-not $found) { throw "시트에서 '$key' 컬럼을 찾을 수 없습니다 (시도한 헤더명: $($HEADER_ALIASES[$key] -join ', '))" }
        $col[$key] = $colByHeader[$found]
    }

    $textByPart = @{}
    $partToRow  = @{}
    $partOrder  = @()
    foreach ($row in $rowsXml) {
        $rNum = [int]$row.GetAttribute("r")
        if ($rNum -eq 1) { continue }
        if ($row.GetAttribute("hidden") -eq "1") { continue }

        $partCell = $row.SelectSingleNode("s:c[@r='$($col['PART'])$rNum']", $ns)
        $part = Get-CellValue $partCell $ns $sharedStrings
        if (-not $part) { continue }

        $nameCell  = $row.SelectSingleNode("s:c[@r='$($col['NAME'])$rNum']", $ns)
        $name = ((Get-CellValue $nameCell $ns $sharedStrings) -replace '^\s+|\s+$', '')
        # 부품번호만 채워지고 상품명이 비어있는 행은 아직 등록이 끝나지 않은 미완성 항목이므로 제외한다.
        if (-not $name) { continue }

        $qtyCell   = $row.SelectSingleNode("s:c[@r='$($col['QTY'])$rNum']", $ns)
        $priceCell = $row.SelectSingleNode("s:c[@r='$($col['PRICE'])$rNum']", $ns)
        $catCell   = $row.SelectSingleNode("s:c[@r='$($col['CATEGORY'])$rNum']", $ns)

        $qtyRaw   = Get-CellValue $qtyCell $ns $sharedStrings
        $priceRaw = Get-CellValue $priceCell $ns $sharedStrings

        $textByPart[$part] = [PSCustomObject]@{
            part     = $part
            name     = $name
            qty      = [int]($qtyRaw -replace '[^0-9\-]', '')
            price    = [int]($priceRaw -replace '[^0-9]', '')
            category = ((Get-CellValue $catCell $ns $sharedStrings) -replace '^\s+|\s+$', '')
        }
        $partToRow[$part] = $rNum
        $partOrder += $part
    }
    if ($textByPart.Count -lt 100) { throw "SHEET_ROWS_LOW: 시트에서 파싱된 상품 데이터가 비정상적으로 적습니다 ($($textByPart.Count)개) - 시트 접근 실패 의심" }
    Write-Log "시트 파싱 완료: $($textByPart.Count)개 상품 텍스트 데이터, 노출 행 $($partToRow.Count)개"

    # 2) 해당 시트의 drawing 관계 찾기
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
        if ($MAIN_LIST_EXCLUDE_PART_NUMBERS -contains $part) { continue }
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
    $newProducts = $newProducts | Sort-Object { [array]::IndexOf($partOrder, $_.part) }
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

    # 6.5) gift.html 갱신 (GIFTS 배열) - 5 STAR GIFT 후보 중 메인 컬렉션에 남아있고
    #      재고(qty)가 0보다 큰 항목만 노출. 메인 컬렉션에서 빠지거나 품절되면 자동으로 사라짐.
    $productByPart = @{}
    foreach ($p in $newProducts) { $productByPart[$p.part.Trim()] = $p }

    $giftProducts = @()
    $giftExcluded = 0
    foreach ($gp in $GIFT_PART_NUMBERS) {
        if ($productByPart.ContainsKey($gp) -and $productByPart[$gp].qty -gt 0) {
            $mp = $productByPart[$gp]
            $giftProducts += [PSCustomObject]@{ part = $mp.part; name = $mp.name; category = $mp.category; image = $mp.image }
        } else {
            $giftExcluded++
        }
    }

    if (Test-Path $GiftPath) {
        $giftLines = Get-Content -LiteralPath $GiftPath -Encoding UTF8
        $giftCompactJson = ($giftProducts | ConvertTo-Json -Depth 3 -Compress)

        $giftsLineIdx = -1
        for ($i = 0; $i -lt $giftLines.Count; $i++) {
            if ($giftLines[$i] -match 'const\s+GIFTS\s*=') { $giftsLineIdx = $i; break }
        }
        if ($giftsLineIdx -lt 0) {
            Write-Log "gift.html에서 'const GIFTS =' 라인을 찾을 수 없어 증정 페이지는 갱신하지 않았습니다"
        } else {
            $giftLines[$giftsLineIdx] = "const GIFTS = $giftCompactJson;"
            Set-Content -LiteralPath $GiftPath -Value $giftLines -Encoding utf8
            Write-Log "gift.html 갱신 완료: 증정 후보 $($GIFT_PART_NUMBERS.Count)개 중 $($giftProducts.Count)개 노출 (메인 컬렉션에 없거나 품절되어 제외 $giftExcluded 개)"
        }
    } else {
        Write-Log "gift.html이 없어 증정 페이지는 갱신하지 않았습니다"
    }

    # 7) GitHub Pages / Vercel(모바일용 실제 배포 사이트)에 반영
    # 주의: git/vercel은 성공 시에도 진행 메시지를 stderr로 출력하는 경우가 많음.
    # ErrorActionPreference=Stop 상태에서 stderr를 그대로 파이프하면 성공한 호출도
    # 예외로 오판될 수 있으므로, 이 블록에서는 Continue로 바꾸고 $LASTEXITCODE로만 판단한다.
    $prevEAP = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    try {
        Push-Location $ProjDir
        $env:PATH += ";C:\Program Files\GitHub CLI;C:\Program Files\nodejs;$env:APPDATA\npm"
        git add -A | Out-Null
        $changed = git status --porcelain
        if ($changed) {
            git commit -m "자동 동기화: $syncTimeText" | Out-Null
            git push origin master
            if ($LASTEXITCODE -ne 0) { throw "git push 실패 (exit code $LASTEXITCODE)" }
            Write-Log "GitHub Pages에 배포 완료"

            vercel --yes --prod
            if ($LASTEXITCODE -ne 0) { throw "vercel 배포 실패 (exit code $LASTEXITCODE)" }
            Write-Log "Vercel에 배포 완료"
        } else {
            Write-Log "반영할 변경사항 없음 (GitHub/Vercel 재배포 생략)"
        }
        Pop-Location
    } catch {
        Write-Log "배포 실패(로컬 파일은 정상 갱신됨): $($_.Exception.Message)"
        Pop-Location
    } finally {
        $ErrorActionPreference = $prevEAP
    }
}
catch {
    Write-Log "동기화 실패: $($_.Exception.Message)"
    if ($_.Exception.Message -like "SHEET_ROWS_LOW:*") {
        New-Item -ItemType File -Force -Path $RetryMarkerPath | Out-Null
        Write-Log "시트 파싱 데이터 부족으로 인한 실패 - 18:00에 재시도 예약됨"
    }
    exit 1
}
finally {
    if (Test-Path $TmpDir) { Remove-Item -LiteralPath $TmpDir -Recurse -Force -ErrorAction SilentlyContinue }
}
