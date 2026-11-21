<#
  AccessVsc.TestReport.ps1
  実施したテストの項目と結果を Excel（.xlsx）に出力する。Excel がなくても作れる（Open XML を直接書く）。
    シート：概要 / 単体テスト / 結合テスト / システムテスト
  元にするもの：
    .accessvsc/test-results/last-test.json      最後に実行したテスト（access.ps1 test）
    .accessvsc/test-results/last-snapshot.json  最後の帳票スナップショット（システムテスト：帳票の見た目）
    modules/*.bas                                テストのコード（確認内容＝直前のコメント、手順、期待結果＝Assert）
  区分の決め方：
    1. テストの直前のコメント、またはモジュール先頭のコメントに「区分: 単体 / 結合 / システム」（@category unit / integration / system も可）
    2. 書いていなければ、UI モジュールでフォームを操作するテストは結合テスト、それ以外は単体テスト
#>

$script:TestCategories = @('単体テスト', '結合テスト', 'システムテスト')

function Get-TestCategoryTag([string[]]$CommentLines) {
    foreach ($c in $CommentLines) {
        if ($c -match '区分\s*[:：]\s*(単体|結合|システム)') { return "$($Matches[1])テスト" }
        if ($c -match '@category\s+(unit|integration|system)') { return @{ unit = '単体テスト'; integration = '結合テスト'; system = 'システムテスト' }[$Matches[1]] }
    }
    return $null
}

# カンマで区切る（括弧と文字列の中は区切らない）
function Split-VbaArgs([string]$Text) {
    $out = New-Object System.Collections.Generic.List[string]
    $depth = 0; $inStr = $false; $cur = New-Object System.Text.StringBuilder
    foreach ($ch in $Text.ToCharArray()) {
        if ($ch -eq '"') { $inStr = -not $inStr }
        elseif (-not $inStr -and $ch -eq '(') { $depth++ }
        elseif (-not $inStr -and $ch -eq ')') { $depth-- }
        elseif (-not $inStr -and $depth -eq 0 -and $ch -eq ',') { $out.Add($cur.ToString().Trim()); [void]$cur.Clear(); continue }
        [void]$cur.Append($ch)
    }
    if ($cur.Length) { $out.Add($cur.ToString().Trim()) }
    return , $out.ToArray()
}
function Get-Unquoted([string]$s) { if ($s -match '^"(.*)"$') { return $Matches[1] -replace '""', '"' } return $s }
# 「Name 引数」または「Name(引数)」の引数部分（外側の括弧は、全体を囲んでいるときだけ外す）
function Get-CallArgs([string]$Rest) {
    $r = $Rest.Trim()
    if ($r.StartsWith('(') -and $r.EndsWith(')')) {
        $depth = 0; $whole = $true
        for ($i = 0; $i -lt $r.Length; $i++) {
            if ($r[$i] -eq '(') { $depth++ } elseif ($r[$i] -eq ')') { $depth--; if ($depth -eq 0 -and $i -lt $r.Length - 1) { $whole = $false; break } }
        }
        if ($whole) { $r = $r.Substring(1, $r.Length - 2) }
    }
    return , (Split-VbaArgs $r)
}

# 手順を読みやすく（UI モジュールの操作は日本語に）
function ConvertTo-ReadableStep([string]$Line) {
    $l = $Line.Trim()
    if ($l -match '^UI\.(\w+)(.*)$') {
        $a = Get-CallArgs $Matches[2]
        switch ($Matches[1]) {
            'OpenForm' { return "フォーム「$(Get-Unquoted $a[0])」を開く$(if ($a.Count -gt 1) { "（条件 $(Get-Unquoted $a[1])）" })" }
            'TypeText' { return "$(Get-Unquoted $a[0]) に「$(Get-Unquoted $a[1])」を入力する" }
            'Click' { return "$(Get-Unquoted $a[0]) を押す" }
            'PressKey' { return "キー $($a[0] -replace '^UI_KEY_', '') を押す" }
            'Focus' { return "$(Get-Unquoted $a[0]) にフォーカスを移す" }
            'CloseForm' { return "フォーム$(if ($a.Count -and $a[0]) { "「$(Get-Unquoted $a[0])」" })を閉じる" }
            'UseForm' { return "操作の対象をフォーム「$(Get-Unquoted $a[0])」にする" }
        }
    }
    return $l
}

# 期待結果を読みやすく（Assert）
function ConvertTo-ReadableExpect([string]$Line) {
    $l = $Line.Trim()
    if ($l -notmatch '^Assert\.(\w+)(.*)$') { return $null }
    $kind = $Matches[1]; $a = Get-CallArgs $Matches[2]
    $msg = ''
    switch ($kind) {
        'AreEqual' { $s = "$($a[1]) が $($a[0])"; if ($a.Count -gt 2) { $msg = Get-Unquoted $a[2] } }
        'AreNotEqual' { $s = "$($a[1]) が $($a[0]) ではない"; if ($a.Count -gt 2) { $msg = Get-Unquoted $a[2] } }
        'AreClose' { $s = "$($a[1]) が $($a[0])（誤差 $(if ($a.Count -gt 2) { $a[2] } else { '0.000001' }) 以内）"; if ($a.Count -gt 3) { $msg = Get-Unquoted $a[3] } }
        'IsTrue' { $s = "$($a[0]) が成り立つ"; if ($a.Count -gt 1) { $msg = Get-Unquoted $a[1] } }
        'IsFalse' { $s = "$($a[0]) が成り立たない"; if ($a.Count -gt 1) { $msg = Get-Unquoted $a[1] } }
        'Contains' { $s = "$($a[0]) に $($a[1]) が含まれる"; if ($a.Count -gt 2) { $msg = Get-Unquoted $a[2] } }
        'IsNullValue' { $s = "$($a[0]) が Null" }
        'IsNotNullValue' { $s = "$($a[0]) が Null ではない" }
        'IsNothing' { $s = "$($a[0]) が Nothing" }
        'IsNotNothing' { $s = "$($a[0]) が Nothing ではない" }
        'Fail' { $s = '（ここに来たら失敗）' }
        default { $s = $l }
    }
    if ($msg) { $s += "（$msg）" }
    return $s
}

# テストのコードから、確認内容・手順・期待結果・区分を読み取る
function Get-TestCaseDetail($Case) {
    $lines = @((Read-SourceText $Case.file) -split "`r?`n")
    $start = [int]$Case.line - 1
    $doc = New-Object System.Collections.Generic.List[string]
    for ($k = $start - 1; $k -ge 0; $k--) {
        $c = $lines[$k].Trim()
        if ($c -match "^'") { $doc.Insert(0, ($c -replace "^'+\s?", '')) } else { break }
    }
    $header = New-Object System.Collections.Generic.List[string]
    foreach ($l in $lines) {
        if ($l -match '^\s*(Public|Private|Sub|Function)\b') { break }
        if ($l.Trim() -match "^'") { $header.Add(($l.Trim() -replace "^'+\s?", '')) }
    }
    $steps = New-Object System.Collections.Generic.List[string]
    $expects = New-Object System.Collections.Generic.List[string]
    $usesUi = $false
    for ($k = $start + 1; $k -lt $lines.Count; $k++) {
        $l = $lines[$k].Trim()
        if ($l -match '^End\s+(Sub|Function)\b') { break }
        if (-not $l -or $l -match "^'" -or $l -match '^(Dim|Const|Static)\s') { continue }
        if ($l -match '^UI\.') { $usesUi = $true }
        $e = ConvertTo-ReadableExpect $l
        if ($e) { $expects.Add($e) } else { $steps.Add((ConvertTo-ReadableStep $l)) }
    }
    $docText = ($doc | Where-Object { $_ -notmatch '区分\s*[:：]|@category' }) -join "`n"
    $category = Get-TestCategoryTag $doc.ToArray()
    if (-not $category) { $category = Get-TestCategoryTag $header.ToArray() }
    if (-not $category) { $category = if ($usesUi) { '結合テスト' } else { '単体テスト' } }
    return [ordered]@{ category = $category; purpose = $docText; steps = $steps.ToArray(); expects = $expects.ToArray() }
}

# ------------------------------------------------------------
# .xlsx を書く（共有文字列を使わず、セルに文字を直接入れる）
# ------------------------------------------------------------
function ConvertTo-XmlText([string]$s) {
    if ($null -eq $s) { return '' }
    $s = $s -replace '[\x00-\x08\x0B\x0C\x0E-\x1F]', ''
    return [System.Security.SecurityElement]::Escape($s)
}
function Get-ColumnName([int]$i) {
    $n = ''; $i++
    while ($i -gt 0) { $m = ($i - 1) % 26; $n = [char](65 + $m) + $n; $i = [Math]::Floor(($i - $m) / 26) }
    return $n
}
# Sheet: @{ name; widths = @(数値…); rows = @( @{ cells = @(値…); styles = @(番号…) or style = 番号 } ); freezeRows; filterRow（見出しの行番号 1 始まり）}
# スタイル番号：0 標準 / 1 見出し / 2 本文（枠・折り返し）/ 3 合格 / 4 不合格 / 5 タイトル / 6 本文（数値・中央）/ 7 注意（黄）
function New-SheetXml($Sheet) {
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
    if ($Sheet.freezeRows) { [void]$sb.Append("<sheetViews><sheetView workbookViewId=""0""><pane ySplit=""$($Sheet.freezeRows)"" topLeftCell=""A$($Sheet.freezeRows + 1)"" activePane=""bottomLeft"" state=""frozen""/></sheetView></sheetViews>") }
    else { [void]$sb.Append('<sheetViews><sheetView workbookViewId="0"/></sheetViews>') }
    [void]$sb.Append('<cols>')
    for ($i = 0; $i -lt $Sheet.widths.Count; $i++) { [void]$sb.Append("<col min=""$($i + 1)"" max=""$($i + 1)"" width=""$($Sheet.widths[$i])"" customWidth=""1""/>") }
    [void]$sb.Append('</cols><sheetData>')
    $r = 0; $maxCol = 0
    foreach ($row in $Sheet.rows) {
        $r++
        [void]$sb.Append("<row r=""$r"">")
        $cells = @($row.cells)
        for ($c = 0; $c -lt $cells.Count; $c++) {
            $v = $cells[$c]
            $style = if ($row.styles) { [int]@($row.styles)[[Math]::Min($c, @($row.styles).Count - 1)] } elseif ($null -ne $row.style) { [int]$row.style } else { 0 }
            $ref = (Get-ColumnName $c) + $r
            if ($null -eq $v -or "$v" -eq '') { [void]$sb.Append("<c r=""$ref"" s=""$style""/>") }
            elseif ($v -is [int] -or $v -is [long] -or $v -is [double] -or $v -is [decimal]) { [void]$sb.Append("<c r=""$ref"" s=""$style""><v>$([System.Convert]::ToString($v, [System.Globalization.CultureInfo]::InvariantCulture))</v></c>") }
            else { [void]$sb.Append("<c r=""$ref"" s=""$style"" t=""inlineStr""><is><t xml:space=""preserve"">$(ConvertTo-XmlText ([string]$v))</t></is></c>") }
        }
        if ($cells.Count -gt $maxCol) { $maxCol = $cells.Count }
        [void]$sb.Append('</row>')
    }
    [void]$sb.Append('</sheetData>')
    if ($Sheet.filterRow -and $r -gt $Sheet.filterRow) { [void]$sb.Append("<autoFilter ref=""A$($Sheet.filterRow):$(Get-ColumnName ($maxCol - 1))$r""/>") }
    [void]$sb.Append('<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/><pageSetup paperSize="9" orientation="landscape" fitToHeight="0"/></worksheet>')
    return $sb.ToString()
}
function New-XlsxFile([string]$Path, $Sheets) {
    Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
    $dir = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
    if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }
    $u8 = New-Object System.Text.UTF8Encoding($false)
    $zip = [System.IO.Compression.ZipFile]::Open($Path, 'Create')
    try {
        $add = { param($name, $text) $e = $zip.CreateEntry($name); $w = New-Object System.IO.StreamWriter($e.Open(), $u8); $w.Write($text); $w.Dispose() }
        $over = ($Sheets | ForEach-Object -Begin { $i = 0 } -Process { $i++; "<Override PartName=""/xl/worksheets/sheet$i.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>" }) -join ''
        & $add '[Content_Types].xml' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' + $over + '</Types>')
        & $add '_rels/.rels' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
        $sheetsXml = ''; $relsXml = ''
        for ($i = 0; $i -lt $Sheets.Count; $i++) {
            $sheetsXml += "<sheet name=""$(ConvertTo-XmlText $Sheets[$i].name)"" sheetId=""$($i + 1)"" r:id=""rId$($i + 1)""/>"
            $relsXml += "<Relationship Id=""rId$($i + 1)"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""worksheets/sheet$($i + 1).xml""/>"
        }
        $relsXml += "<Relationship Id=""rId$($Sheets.Count + 1)"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"" Target=""styles.xml""/>"
        # オートフィルター用の名前（Excel が必要とする）
        $names = ''
        for ($i = 0; $i -lt $Sheets.Count; $i++) {
            $s = $Sheets[$i]
            if ($s.filterRow -and @($s.rows).Count -gt $s.filterRow) {
                $maxCol = (@($s.rows | ForEach-Object { @($_.cells).Count }) | Measure-Object -Maximum).Maximum
                $names += "<definedName name=""_xlnm._FilterDatabase"" localSheetId=""$i"" hidden=""1"">'$($s.name -replace "'", "''")'!`$A`$$($s.filterRow):`$$(Get-ColumnName ($maxCol - 1))`$$(@($s.rows).Count)</definedName>"
            }
        }
        & $add 'xl/workbook.xml' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' + $sheetsXml + '</sheets>' + $(if ($names) { "<definedNames>$names</definedNames>" }) + '</workbook>')
        & $add 'xl/_rels/workbook.xml.rels' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + $relsXml + '</Relationships>')
        & $add 'xl/styles.xml' @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="3"><font><sz val="10"/><name val="Meiryo UI"/><family val="3"/><charset val="128"/></font><font><b/><sz val="10"/><name val="Meiryo UI"/><family val="3"/><charset val="128"/></font><font><b/><sz val="14"/><name val="Meiryo UI"/><family val="3"/><charset val="128"/></font></fonts>
<fills count="6"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFDDE5F2"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFE2EFDA"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFCE4D6"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FFA6A6A6"/></left><right style="thin"><color rgb="FFA6A6A6"/></right><top style="thin"><color rgb="FFA6A6A6"/></top><bottom style="thin"><color rgb="FFA6A6A6"/></bottom><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="8">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
<xf numFmtId="0" fontId="1" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="top"/></xf>
<xf numFmtId="0" fontId="1" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="top"/></xf>
<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="top"/></xf>
<xf numFmtId="0" fontId="1" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="top"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="標準" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>
'@
        for ($i = 0; $i -lt $Sheets.Count; $i++) { & $add "xl/worksheets/sheet$($i + 1).xml" (New-SheetXml $Sheets[$i]) }
    }
    finally { $zip.Dispose() }
}

# ------------------------------------------------------------
# 入口
# ------------------------------------------------------------
function Invoke-TestReport([string]$Db, [string]$Src, [string]$OutPath) {
    $resDir = Join-Path (Split-Path -Parent $Db) '.accessvsc\test-results'
    $lastTest = Join-Path $resDir 'last-test.json'
    $lastSnap = Join-Path $resDir 'last-snapshot.json'
    $test = if (Test-Path -LiteralPath $lastTest) { (Read-SourceText $lastTest) | ConvertFrom-Json } else { $null }
    $snap = if (Test-Path -LiteralPath $lastSnap) { (Read-SourceText $lastSnap) | ConvertFrom-Json } else { $null }
    if (-not $test -and -not $snap) { throw 'No test results yet. Run the tests first (access.ps1 test, or test-report -Run).' }

    $cases = @{}
    foreach ($c in @(Get-TestCases $Src)) { $cases[$c.id] = $c }
    $rows = @{}; foreach ($cat in $script:TestCategories) { $rows[$cat] = New-Object System.Collections.Generic.List[object] }
    $ranAt = if ($test) { [string]$test.ranAt } else { [string]$snap.ranAt }
    foreach ($t in @($(if ($test -and $test.result) { $test.result.tests } else { @() }))) {
        $c = $cases[[string]$t.id]
        $detail = if ($c) { Get-TestCaseDetail $c } else { [ordered]@{ category = '単体テスト'; purpose = ''; steps = @(); expects = @() } }
        $mark = switch ([string]$t.status) { 'passed' { '○' } 'failed' { '×' } default { 'エラー' } }
        $where = if ($t.location -and $t.location.file) { "$([System.IO.Path]::GetFileName([string]$t.location.file)):$($t.location.fileLine)" } elseif ($c) { "$([System.IO.Path]::GetFileName([string]$c.file)):$($c.line)" } else { '' }
        $rows[$detail.category].Add([ordered]@{
                id = [string]$t.id; purpose = $(if ($detail.purpose) { $detail.purpose } else { [string]$t.name })
                steps = (@($detail.steps | ForEach-Object -Begin { $n = 0 } -Process { $n++; "$n. $_" }) -join "`n")
                expects = (@($detail.expects) -join "`n"); mark = $mark; status = [string]$t.status
                actual = [string]$t.message; ms = [int]$t.ms; where = $where
            })
    }
    foreach ($r in @($(if ($snap -and $snap.result) { $snap.result.reports } else { @() }))) {
        $mark = switch ([string]$r.status) { 'same' { '○' } 'updated' { '○' } 'new' { '要確認' } 'changed' { '×' } default { 'エラー' } }
        $jp = @{ same = '同じ'; changed = '変化あり'; new = '基準なし'; updated = '基準を更新'; added = 'ページが増えた'; removed = 'ページが減った'; error = 'エラー' }
        $pages = @($r.pages | Where-Object { $_ -and $_.status -ne 'same' } | ForEach-Object { "$($_.page) ページ：$(if ($jp.ContainsKey([string]$_.status)) { $jp[[string]$_.status] } else { $_.status })$(if ($_.diffPercent) { "（$($_.diffPercent)% のピクセルが変化）" })" }) -join "`n"
        $rows['システムテスト'].Add([ordered]@{
                id = "帳票.$($r.name)"; purpose = "帳票 $($r.name) の見た目が、承認済みの基準と同じであること"
                steps = "1. 帳票を PDF に出力し、ページごとに画像にする`n2. 基準画像（snapshots）とピクセル単位で比べる"
                expects = '差がない（基準と同じ）'; mark = $mark; status = [string]$r.status
                actual = ((@($r.message, $pages) | Where-Object { $_ }) -join "`n"); ms = $null; where = "snapshots/$($r.name)"
            })
    }

    # 概要
    $dbName = [System.IO.Path]::GetFileName($Db)
    $summaryRows = New-Object System.Collections.Generic.List[object]
    $summaryRows.Add(@{ cells = @('テスト結果報告書'); style = 5 })
    $summaryRows.Add(@{ cells = @() })
    $summaryRows.Add(@{ cells = @('データベース', $dbName); styles = @(1, 2) })
    $summaryRows.Add(@{ cells = @('実施日時', ($ranAt -replace 'T', ' ')); styles = @(1, 2) })
    $summaryRows.Add(@{ cells = @('作成日時', (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')); styles = @(1, 2) })
    $summaryRows.Add(@{ cells = @('作成', 'Access VSC'); styles = @(1, 2) })
    $summaryRows.Add(@{ cells = @() })
    $summaryRows.Add(@{ cells = @('区分', '件数', '合格', '不合格', 'エラー・要確認', '合格率'); style = 1 })
    $tot = @{ n = 0; ok = 0; ng = 0; er = 0 }
    foreach ($cat in $script:TestCategories) {
        $list = $rows[$cat]
        $ok = @($list | Where-Object { $_.mark -eq '○' }).Count; $ng = @($list | Where-Object { $_.mark -eq '×' }).Count; $er = $list.Count - $ok - $ng
        $tot.n += $list.Count; $tot.ok += $ok; $tot.ng += $ng; $tot.er += $er
        $rate = if ($list.Count) { "{0:0.0}%" -f (100.0 * $ok / $list.Count) } else { '—' }
        $summaryRows.Add(@{ cells = @($cat, $list.Count, $ok, $ng, $er, $rate); styles = @(2, 6, 6, 6, 6, 6) })
    }
    $summaryRows.Add(@{ cells = @('合計', $tot.n, $tot.ok, $tot.ng, $tot.er, $(if ($tot.n) { "{0:0.0}%" -f (100.0 * $tot.ok / $tot.n) } else { '—' })); styles = @(1, 6, 6, 6, 6, 6) })
    $summaryRows.Add(@{ cells = @() })
    $summaryRows.Add(@{ cells = @('区分の決め方：テストの直前（またはモジュールの先頭）のコメントに「区分: 単体 / 結合 / システム」と書くと、その区分になります。書いていない場合、UI モジュールでフォームを操作するテストは結合テスト、それ以外は単体テストです。システムテストには、帳票スナップショット（見た目の比較）の結果も載ります。') })
    $sheets = New-Object System.Collections.Generic.List[object]
    $sheets.Add(@{ name = '概要'; widths = @(18, 30, 10, 10, 14, 10); rows = $summaryRows.ToArray() })

    foreach ($cat in $script:TestCategories) {
        $list = $rows[$cat]
        $r = New-Object System.Collections.Generic.List[object]
        $r.Add(@{ cells = @("$cat（$dbName）"); style = 5 })
        $r.Add(@{ cells = @("実施日時 $($ranAt -replace 'T', ' ')　／　$($list.Count) 件") })
        $r.Add(@{ cells = @('No', 'テスト ID', '確認内容', '手順', '期待結果', '結果', '実際の値・メッセージ', '時間(ms)', '場所'); style = 1 })
        $n = 0
        foreach ($x in $list) {
            $n++
            $ms = $(if ($null -ne $x.ms) { [int]$x.ms } else { '' })
            $markStyle = switch ($x.mark) { '○' { 3 } '×' { 4 } default { 7 } }
            $r.Add(@{ cells = @($n, $x.id, $x.purpose, $x.steps, $x.expects, $x.mark, $x.actual, $ms, $x.where); styles = @(6, 2, 2, 2, 2, $markStyle, 2, 6, 2) })
        }
        if (-not $list.Count) { $r.Add(@{ cells = @('', '（該当するテストはありません）'); styles = @(2, 2) }) }
        $sheets.Add(@{ name = $cat; widths = @(5, 30, 32, 40, 36, 7, 34, 9, 18); rows = $r.ToArray(); freezeRows = 3; filterRow = 3 })
    }

    if (-not $OutPath) {
        $base = [System.IO.Path]::GetFileNameWithoutExtension($Db)
        $OutPath = Join-Path (Split-Path -Parent $Db) ("test-reports\テスト結果_{0}_{1}.xlsx" -f $base, (Get-Date).ToString('yyyyMMdd_HHmm'))
    }
    New-XlsxFile $OutPath $sheets.ToArray()
    $counts = [ordered]@{}
    foreach ($cat in $script:TestCategories) { $counts[$cat] = $rows[$cat].Count }
    return [ordered]@{ path = $OutPath; ranAt = $ranAt; counts = $counts; passed = $tot.ok; failed = $tot.ng; errors = $tot.er; total = $tot.n }
}
