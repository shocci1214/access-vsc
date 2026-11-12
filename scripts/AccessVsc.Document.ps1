<#
  AccessVsc.Document.ps1
  取り出したソースフォルダから、データベースの構造ドキュメント（Markdown）を作る。Access は使わない。
    index.md     概要・ER 図（Mermaid）・オブジェクト一覧
    tables.md    テーブル：項目・キー・インデックス・リレーション・使っているところ
    queries.md   クエリ：SQL・元にしているもの・使っているところ
    forms.md     フォーム：元データ・コントロール・イベント・コードの手続き・使っているもの・使っているところ
    reports.md   レポート：元データ・グループ・コントロール・使っているもの・使っているところ
    modules.md   VBA：手続きの一覧（説明コメントつき）・使っているもの・呼ばれているところ
  各オブジェクトに「説明」欄があり、人や AI が書いた説明は、作り直しても残る。
  対応表（使っている・使われている）は、名前が出てくるかどうかで調べる（コメントは除く。動的に組み立てた名前は見つからない）。
#>

$script:DocPlaceholder = '_（説明：未記入。ここに書いた内容は、ドキュメントを作り直しても残ります）_'

# ------------------------------------------------------------
# 読み取り
# ------------------------------------------------------------
# VBA のコメントだけを取り除く（文字列は残す。テーブル名などは文字列の中に出てくるため）
function Remove-VbaComment([string]$Line) {
    $inStr = $false
    for ($i = 0; $i -lt $Line.Length; $i++) {
        $ch = $Line[$i]
        if ($ch -eq '"') { $inStr = -not $inStr; continue }
        if (-not $inStr -and $ch -eq "'") { return $Line.Substring(0, $i) }
    }
    if ($Line -match '^\s*Rem\b') { return '' }
    return $Line
}

# VBA の手続き（Sub / Function / Property）
function Get-VbaProcedures([string[]]$Lines, [int]$Offset = 0) {
    $out = New-Object System.Collections.Generic.List[object]
    $re = [regex]'(?i)^\s*(?:(Public|Private|Friend|Global)\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_\u0080-\uFFFF][\w\u0080-\uFFFF]*)\s*\((.*)$'
    for ($i = $Offset; $i -lt $Lines.Count; $i++) {
        $m = $re.Match($Lines[$i])
        if (-not $m.Success) { continue }
        $sig = $Lines[$i].Trim()
        $j = $i
        while ($Lines[$j] -match '\s_\s*$' -and $j + 1 -lt $Lines.Count) { $j++; $sig = ($sig -replace '\s_\s*$', ' ') + $Lines[$j].Trim() }
        $doc = New-Object System.Collections.Generic.List[string]
        for ($k = $i - 1; $k -ge $Offset; $k--) {
            $c = $Lines[$k].Trim()
            if ($c -match "^'") { $doc.Insert(0, ($c -replace "^'+\s?", '')) } else { break }
        }
        $out.Add([ordered]@{
                name  = $m.Groups[3].Value
                scope = $(if ($m.Groups[1].Value) { (Get-Culture).TextInfo.ToTitleCase($m.Groups[1].Value.ToLower()) } else { 'Public' })
                kind  = ($m.Groups[2].Value -replace '\s+', ' ')
                signature = ($sig -replace '\s+', ' ')
                doc   = ($doc -join ' ')
                line  = $i + 1
            })
    }
    return , $out.ToArray()
}

# SaveAsText のプロパティ行（Key ="値" / Key =値）。続きの行（"..." だけの行）もつなぐ
function ConvertFrom-AccessText([string[]]$Lines) {
    # 戻り値：@{ props = 先頭のオブジェクトのプロパティ; controls = コントロールの一覧; sections; groups; codeStart }
    $result = [ordered]@{ props = [ordered]@{}; controls = New-Object System.Collections.Generic.List[object]; sections = New-Object System.Collections.Generic.List[string]; groups = New-Object System.Collections.Generic.List[object]; codeStart = -1 }
    $stack = New-Object System.Collections.Generic.List[object]
    $sectionTypes = @('Section', 'PageHeader', 'PageFooter', 'FormHeader', 'FormFooter', 'BreakHeader', 'BreakFooter')
    $lastKey = $null; $lastObj = $null
    for ($i = 0; $i -lt $Lines.Count; $i++) {
        $line = $Lines[$i]
        $t = $line.Trim()
        if ($t -match '^CodeBehind(Form|Report)$') { $result.codeStart = $i + 1; break }
        if ($t -match '^[A-Za-z]\w* = Begin$') {
            # バイナリのプロパティ（GUID・PrtDevMode など）：対応する End まで読み飛ばす
            $depth = 1
            while ($depth -gt 0 -and ++$i -lt $Lines.Count) {
                $u = $Lines[$i].Trim()
                if ($u -match '= Begin$' -or $u -eq 'Begin') { $depth++ } elseif ($u -eq 'End') { $depth-- }
            }
            $lastKey = $null
            continue
        }
        if ($t -match '^Begin(?:\s+(\w+))?$') {
            $obj = [ordered]@{ type = $Matches[1]; props = [ordered]@{} }
            $stack.Add($obj); $lastKey = $null
            continue
        }
        if ($t -eq 'End') {
            if ($stack.Count) {
                $obj = $stack[$stack.Count - 1]; $stack.RemoveAt($stack.Count - 1)
                if ($obj.type -in $sectionTypes -and $obj.props.Name) { $result.sections.Add([string]$obj.props.Name) }
                elseif ($obj.type -eq 'BreakLevel') { $result.groups.Add([ordered]@{ field = $obj.props.ControlSource; header = [bool]$obj.props.GroupHeader; footer = [bool]$obj.props.GroupFooter }) }
                elseif ($obj.type -and $obj.props.Name -and $stack.Count -ge 2) {
                    # 一番外側（フォーム／レポート）でも、セクションでもない、名前のあるもの＝コントロール
                    $sec = $null
                    foreach ($s in $stack) { if ($s.type -in $sectionTypes) { $sec = $s.props.Name } }
                    $obj.section = $sec
                    $result.controls.Add($obj)
                }
            }
            $lastKey = $null
            continue
        }
        if ($t -match '^"(.*)"$' -and $lastKey -and $lastObj) {
            $lastObj.props[$lastKey] = [string]$lastObj.props[$lastKey] + ($Matches[1] -replace '\\"', '"')
            continue
        }
        if ($t -match '^([A-Za-z]\w*)\s*=\s*(.*)$' -and $stack.Count) {
            $key = $Matches[1]; $val = $Matches[2].Trim()
            if ($val -match '^"(.*)"$') { $val = $Matches[1] -replace '\\"', '"' }
            elseif ($val -eq 'NotDefault') { $val = $true }
            $obj = $stack[$stack.Count - 1]
            $obj.props[$key] = $val
            if ($stack.Count -eq 1) { $result.props[$key] = $val }
            $lastKey = $key; $lastObj = $obj
        }
    }
    return $result
}

# イベントのプロパティ（OnClick・AfterUpdate など）を一覧に
function Get-AccessEvents($Props, [string]$Owner) {
    $out = New-Object System.Collections.Generic.List[object]
    foreach ($k in @($Props.Keys)) {
        if ($k -notmatch '^(On[A-Z]\w*|BeforeUpdate|AfterUpdate|BeforeInsert|AfterInsert|BeforeDelConfirm|AfterDelConfirm)$') { continue }
        $v = [string]$Props[$k]
        if (-not $v) { continue }
        $ev = $k -replace '^On', ''
        $kind = if ($v -eq '[Event Procedure]') { 'procedure' } elseif ($v.StartsWith('=')) { 'expression' } else { 'macro' }
        $out.Add([ordered]@{ event = $ev; value = $v; kind = $kind; handler = $(if ($kind -eq 'procedure') { "${Owner}_$ev" } else { $null }) })
    }
    return , $out.ToArray()
}

# 名前が「単語として」出てくるか（日本語の名前にも対応）
function Test-NameUsed([string]$Text, [string]$Name, [bool]$Call = $false) {
    # Call = $true\uFF1A\u95A2\u6570\u306E\u547C\u3073\u51FA\u3057\uFF08\u540D\u524D\u306E\u3059\u3050\u5F8C\u306B\u300C(\u300D\uFF09\u3060\u3051\u3092\u6570\u3048\u308B\uFF08SQL \u3084\u5F0F\u306E\u5217\u540D\u3068\u533A\u5225\u3059\u308B\u305F\u3081\uFF09
    if (-not $Text -or -not $Name) { return $false }
    $n = [regex]::Escape($Name)
    $tail = if ($Call) { '\s*\(' } else { '(?![\w\u0080-\uFFFF])' }
    if ([regex]::IsMatch($Text, "(?<![\w\u0080-\uFFFF])$n$tail", 'IgnoreCase')) { return $true }
    return (-not $Call) -and [regex]::IsMatch($Text, "\[$n\]", 'IgnoreCase')
}

function Read-DatabaseStructure([string]$Src) {
    $db = [ordered]@{ tables = @(); queries = @(); forms = @(); reports = @(); modules = @(); macros = @(); relationships = @() }
    $tdir = Join-Path $Src 'tables'
    if (Test-Path -LiteralPath $tdir) {
        $db.tables = @(Get-ChildItem -LiteralPath $tdir -Filter '*.json' -File | Sort-Object Name | ForEach-Object { (Read-SourceText $_.FullName) | ConvertFrom-Json })
    }
    $rel = Join-Path $Src 'relationships.json'
    if (Test-Path -LiteralPath $rel) { try { $db.relationships = @(((Read-SourceText $rel) | ConvertFrom-Json).relationships) } catch {} }
    foreach ($rel in @(Get-SourceFiles $Src)) {
        $info = Get-ObjectInfoFromRelPath $rel
        if (-not $info) { continue }
        $full = Join-Path $Src ($rel.Replace('/', '\'))
        switch ($info.kind) {
            'queries' {
                $sqlLines = @((Read-SourceText $full) -split "`r?`n" | Where-Object { $_ -notmatch '^--\s*accessvsc:' })
                $db.queries += [ordered]@{ name = $info.name; file = $rel; sql = ($sqlLines -join "`n").Trim() }
            }
            'modules' {
                $lines = [string[]]@((Read-SourceText $full) -split "`r?`n")
                $db.modules += [ordered]@{ name = $info.name; file = $rel; isClass = ($info.ext -eq '.cls'); procs = (Get-VbaProcedures $lines); code = (($lines | ForEach-Object { Remove-VbaComment $_ }) -join "`n") }
            }
            { $_ -in @('forms', 'reports') } {
                $lines = [string[]]@((Read-SourceText $full) -split "`r?`n")
                $parsed = ConvertFrom-AccessText $lines
                $owner = if ($info.kind -eq 'forms') { 'Form' } else { 'Report' }
                $code = ''; $procs = @()
                if ($parsed.codeStart -ge 0) {
                    $procs = Get-VbaProcedures $lines $parsed.codeStart
                    $code = (($lines | Select-Object -Skip $parsed.codeStart | ForEach-Object { Remove-VbaComment $_ }) -join "`n")
                }
                $controls = @($parsed.controls | ForEach-Object {
                        [ordered]@{
                            name = [string]$_.props.Name; type = [string]$_.type; section = $_.section
                            source = $(if ($_.props.ControlSource) { [string]$_.props.ControlSource } elseif ($_.props.SourceObject) { [string]$_.props.SourceObject } else { $null })
                            rowSource = $(if ($_.props.RowSource) { [string]$_.props.RowSource } else { $null })
                            caption = $(if ($_.props.Caption) { [string]$_.props.Caption } else { $null })
                            events = (Get-AccessEvents $_.props ([string]$_.props.Name))
                        } })
                $item = [ordered]@{
                    name = $info.name; file = $rel; recordSource = [string]$parsed.props.RecordSource; caption = [string]$parsed.props.Caption
                    events = (Get-AccessEvents $parsed.props $owner); controls = $controls; sections = $parsed.sections.ToArray(); groups = $parsed.groups.ToArray()
                    procs = $procs; code = $code
                }
                if ($info.kind -eq 'forms') { $db.forms += $item } else { $db.reports += $item }
            }
            'macros' { $db.macros += [ordered]@{ name = $info.name; file = $rel } }
        }
    }
    return $db
}

# ------------------------------------------------------------
# 対応表（使っている・使われている）
# ------------------------------------------------------------
function Get-DatabaseXref($Db) {
    $targets = New-Object System.Collections.Generic.List[object]
    foreach ($t in $Db.tables) { $targets.Add(@{ kind = 'table'; name = [string]$t.name }) }
    foreach ($q in $Db.queries) { $targets.Add(@{ kind = 'query'; name = $q.name }) }
    foreach ($f in $Db.forms) { $targets.Add(@{ kind = 'form'; name = $f.name }) }
    foreach ($r in $Db.reports) { $targets.Add(@{ kind = 'report'; name = $r.name }) }
    foreach ($m in $Db.modules) {
        foreach ($p in $m.procs) { if ($p.scope -ne 'Private') { $targets.Add(@{ kind = 'procedure'; name = $p.name; module = $m.name }) } }
    }
    # 使う側：そのオブジェクトの中で、名前を探す文字列
    $users = New-Object System.Collections.Generic.List[object]
    # expr：SQL・コントロールの式（手続きは「名前(」だけ数える） / code：VBA のコード
    foreach ($q in $Db.queries) { $users.Add(@{ kind = 'query'; name = $q.name; expr = $q.sql; code = '' }) }
    foreach ($m in $Db.modules) {
        # 自分の手続きの宣言行は除く（定義を「使っている」と数えないため）
        $body = ($m.code -split "`n" | Where-Object { $_ -notmatch '^\s*(?:(?:Public|Private|Friend|Global)\s+)?(?:Static\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s' }) -join "`n"
        $users.Add(@{ kind = 'module'; name = $m.name; expr = ''; code = $body })
    }
    foreach ($kind in 'form', 'report') {
        foreach ($o in @($(if ($kind -eq 'form') { $Db.forms } else { $Db.reports }))) {
            $parts = New-Object System.Collections.Generic.List[string]
            $parts.Add($o.recordSource)
            foreach ($c in $o.controls) {
                if ($c.source) { $parts.Add($c.source) }
                if ($c.rowSource) { $parts.Add($c.rowSource) }
                foreach ($e in $c.events) { if ($e.kind -ne 'procedure') { $parts.Add($e.value) } }
            }
            foreach ($e in $o.events) { if ($e.kind -ne 'procedure') { $parts.Add($e.value) } }
            $body = ($o.code -split "`n" | Where-Object { $_ -notmatch '^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s' }) -join "`n"
            $users.Add(@{ kind = $kind; name = $o.name; expr = ($parts -join "`n"); code = $body })
        }
    }
    $uses = @{}      # "kind:name" -> 使っているもの [ @{kind; name} ]
    $usedBy = @{}    # "kind:name" -> 使われているところ
    foreach ($u in $users) {
        $key = "$($u.kind):$($u.name)"
        $uses[$key] = New-Object System.Collections.Generic.List[object]
        foreach ($t in $targets) {
            if ($t.kind -eq $u.kind -and $t.name -eq $u.name) { continue }
            if ($t.kind -eq 'procedure' -and $u.kind -eq 'module' -and $t.module -eq $u.name) { continue }   # 同じモジュール内の呼び出しは数えない
            if ($t.kind -eq 'form' -and $u.kind -eq 'form' -and $t.name -eq $u.name) { continue }
            $isProc = $t.kind -eq 'procedure'
            if ((Test-NameUsed $u.code $t.name) -or (Test-NameUsed $u.expr $t.name $isProc)) {
                $uses[$key].Add($t)
                $tk = "$($t.kind):$($t.name)"
                if (-not $usedBy.ContainsKey($tk)) { $usedBy[$tk] = New-Object System.Collections.Generic.List[object] }
                $usedBy[$tk].Add(@{ kind = $u.kind; name = $u.name })
            }
        }
    }
    return @{ uses = $uses; usedBy = $usedBy }
}

# ------------------------------------------------------------
# Markdown
# ------------------------------------------------------------
$script:KindLabel = @{ table = 'テーブル'; query = 'クエリ'; form = 'フォーム'; report = 'レポート'; module = 'モジュール'; procedure = '手続き' }
$script:KindFile = @{ table = 'tables.md'; query = 'queries.md'; form = 'forms.md'; report = 'reports.md'; module = 'modules.md'; procedure = 'modules.md' }

function ConvertTo-MdAnchor([string]$Kind, [string]$Name) {
    # 見出しの id（日本語の名前でも壊れないよう、記号だけ置き換える）
    return ("$Kind-" + ($Name -replace '[\s\.\(\)\[\]\{\}#&?/\\:;,!"''<>|=+*%@^~`$]', '-')).ToLowerInvariant()
}
function ConvertTo-MdCell($v) {
    if ($null -eq $v) { return '' }
    return ([string]$v -replace '\|', '\|' -replace "`r?`n", ' ')
}
function Format-MdLink($t) {
    $kind = $t.kind
    $label = if ($kind -eq 'procedure' -and $t.module) { "$($t.module).$($t.name)" } else { $t.name }
    $anchor = if ($kind -eq 'procedure') { ConvertTo-MdAnchor 'module' $t.module } else { ConvertTo-MdAnchor $kind $t.name }
    return "[$label]($($script:KindFile[$kind])#$anchor)"
}
function Format-MdRefs($List, [string]$Empty = '—') {
    if ($null -eq $List -or $List.Count -eq 0) { return $Empty }
    $groups = @($List | Group-Object { $_.kind } | Sort-Object Name)
    return (($groups | ForEach-Object { "$($script:KindLabel[$_.Name])：" + ((@($_.Group | Sort-Object { $_.name } -Unique | ForEach-Object { Format-MdLink $_ })) -join '、') }) -join ' ／ ')
}

# 作り直しても残す「説明」欄
function Read-DocNotes([string]$Dir) {
    $notes = @{}
    if (-not (Test-Path -LiteralPath $Dir)) { return $notes }
    foreach ($f in Get-ChildItem -LiteralPath $Dir -Filter '*.md' -File) {
        $text = Read-SourceText $f.FullName
        foreach ($m in [regex]::Matches($text, '(?s)<!-- notes:(?<key>[^\r\n>]+?) -->\r?\n(?<body>.*?)\r?\n<!-- /notes -->')) {
            $body = $m.Groups['body'].Value.Trim()
            if ($body -and $body -ne $script:DocPlaceholder) { $notes[$m.Groups['key'].Value] = $body }
        }
    }
    return $notes
}
function Format-DocNotes($Notes, [string]$Key) {
    $body = if ($Notes.ContainsKey($Key)) { $Notes[$Key] } else { $script:DocPlaceholder }
    return "<!-- notes:$Key -->`n$body`n<!-- /notes -->"
}

function ConvertTo-MermaidId([string]$Name) {
    if ($Name -match '^[A-Za-z_][A-Za-z0-9_]*$') { return $Name }
    return 'T_' + (Get-StringSha1 $Name).Substring(0, 8)
}

function New-DocIndex($Db, $Xref, $Notes, [string]$DbName) {
    $sb = New-Object System.Text.StringBuilder
    $w = { param($s) [void]$sb.AppendLine($s) }
    & $w "# $DbName の構造"
    & $w ''
    & $w "> Access VSC が、取り出したソース（テキスト）から自動で作ったドキュメントです（$((Get-Date).ToString('yyyy-MM-dd HH:mm'))）。"
    & $w '> 「説明」欄に書いた内容は、作り直しても残ります。対応表（使っている・使われている）は名前が出てくるかどうかで調べているため、コードで組み立てた名前などは見つかりません。'
    & $w ''
    & $w '## 概要'
    & $w ''
    & $w (Format-DocNotes $Notes 'database')
    & $w ''
    & $w '| 種類 | 数 | 一覧 |'
    & $w '|---|---|---|'
    & $w "| テーブル | $(@($Db.tables).Count) | [tables.md](tables.md) |"
    & $w "| クエリ | $(@($Db.queries).Count) | [queries.md](queries.md) |"
    & $w "| フォーム | $(@($Db.forms).Count) | [forms.md](forms.md) |"
    & $w "| レポート | $(@($Db.reports).Count) | [reports.md](reports.md) |"
    & $w "| モジュール | $(@($Db.modules).Count)（手続き $(@($Db.modules | ForEach-Object { $_.procs } | Where-Object { $_ }).Count)） | [modules.md](modules.md) |"
    if (@($Db.macros).Count) { & $w "| マクロ | $(@($Db.macros).Count) | $((@($Db.macros | ForEach-Object { $_.name })) -join '、') |" }
    & $w ''
    if (@($Db.tables).Count) {
        & $w '## テーブルの関係（ER 図）'
        & $w ''
        & $w '```mermaid'
        & $w 'erDiagram'
        foreach ($t in $Db.tables) {
            $id = ConvertTo-MermaidId $t.name
            $label = if ($id -ne $t.name) { "$id[""$($t.name)""]" } else { $id }
            $pk = @($t.indexes | Where-Object { $_.primary } | ForEach-Object { $_.fields })
            $fk = @($Db.relationships | Where-Object { $_.foreignTable -eq $t.name } | ForEach-Object { $_.fields | ForEach-Object { $_.foreignField } })
            & $w "    $label {"
            $n = 0
            foreach ($f in $t.fields) {
                $n++
                $fname = if ($f.name -match '^[A-Za-z_][A-Za-z0-9_]*$') { $f.name } else { "f$n" }
                $mark = @(); if ($pk -contains $f.name) { $mark += 'PK' }; if ($fk -contains $f.name) { $mark += 'FK' }
                $comment = if ($fname -ne $f.name) { " ""$($f.name -replace '"', '')""" } else { '' }
                & $w ("        {0} {1}{2}{3}" -f ($f.type -replace '[^A-Za-z0-9_]', ''), $fname, $(if ($mark) { ' ' + ($mark -join ',') } else { '' }), $comment)
            }
            & $w '    }'
        }
        foreach ($r in $Db.relationships) {
            $map = (@($r.fields | ForEach-Object { "$($_.field) → $($_.foreignField)" })) -join ', '
            & $w ("    {0} ||--o{{ {1} : ""{2}""" -f (ConvertTo-MermaidId $r.table), (ConvertTo-MermaidId $r.foreignTable), ($map -replace '"', ''))
        }
        & $w '```'
        & $w ''
    }
    & $w '## オブジェクト一覧'
    & $w ''
    & $w '| 種類 | 名前 | 元データ・内容 | 使われているところ |'
    & $w '|---|---|---|---|'
    foreach ($t in $Db.tables) { & $w "| テーブル | $(Format-MdLink @{ kind = 'table'; name = $t.name }) | $(if ($t.linked) { 'リンクテーブル' } else { "$($t.recordCount) 件" }) | $(Format-MdRefs $Xref.usedBy["table:$($t.name)"]) |" }
    foreach ($q in $Db.queries) { & $w "| クエリ | $(Format-MdLink @{ kind = 'query'; name = $q.name }) | $(Format-MdRefs $Xref.uses["query:$($q.name)"]) | $(Format-MdRefs $Xref.usedBy["query:$($q.name)"]) |" }
    foreach ($f in $Db.forms) { & $w "| フォーム | $(Format-MdLink @{ kind = 'form'; name = $f.name }) | $(ConvertTo-MdCell $(if ($f.recordSource) { $f.recordSource } else { '（非連結）' })) | $(Format-MdRefs $Xref.usedBy["form:$($f.name)"]) |" }
    foreach ($r in $Db.reports) { & $w "| レポート | $(Format-MdLink @{ kind = 'report'; name = $r.name }) | $(ConvertTo-MdCell $r.recordSource) | $(Format-MdRefs $Xref.usedBy["report:$($r.name)"]) |" }
    foreach ($m in $Db.modules) { & $w "| モジュール | $(Format-MdLink @{ kind = 'module'; name = $m.name }) | 手続き $(@($m.procs).Count) | — |" }
    return $sb.ToString()
}

function New-DocTables($Db, $Xref, $Notes) {
    $sb = New-Object System.Text.StringBuilder
    $w = { param($s) [void]$sb.AppendLine($s) }
    & $w '# テーブル'
    & $w ''
    & $w '[← 概要](index.md)'
    foreach ($t in $Db.tables) {
        & $w ''
        & $w "<a id=""$(ConvertTo-MdAnchor 'table' $t.name)""></a>"
        & $w "## $($t.name)"
        & $w ''
        & $w (Format-DocNotes $Notes "table:$($t.name)")
        & $w ''
        $kind = if ($t.linked) { "リンクテーブル$(if ($t.sourceTable) { "（元：$($t.sourceTable)）" })" } else { "$($t.recordCount) 件" }
        & $w "- 種類：$kind"
        & $w "- 使われているところ：$(Format-MdRefs $Xref.usedBy["table:$($t.name)"] 'なし（見つかりません）')"
        $refs = @($Db.relationships | Where-Object { $_.foreignTable -eq $t.name } | ForEach-Object { "$((@($_.fields | ForEach-Object { $_.foreignField })) -join ', ') → $(Format-MdLink @{ kind = 'table'; name = $_.table })" })
        $refBy = @($Db.relationships | Where-Object { $_.table -eq $t.name } | ForEach-Object { "$(Format-MdLink @{ kind = 'table'; name = $_.foreignTable })（$((@($_.fields | ForEach-Object { $_.foreignField })) -join ', ')）" })
        if ($refs) { & $w "- 参照している（外部キー）：$($refs -join '、')" }
        if ($refBy) { & $w "- 参照されている：$($refBy -join '、')" }
        & $w ''
        $pk = @($t.indexes | Where-Object { $_.primary } | ForEach-Object { $_.fields })
        $fk = @($Db.relationships | Where-Object { $_.foreignTable -eq $t.name } | ForEach-Object { $_.fields | ForEach-Object { $_.foreignField } })
        & $w '| 項目 | 型 | サイズ | 必須 | キー | 備考 |'
        & $w '|---|---|---|---|---|---|'
        foreach ($f in $t.fields) {
            $key = @(); if ($pk -contains $f.name) { $key += '主キー' }; if ($fk -contains $f.name) { $key += '外部キー' }
            $note = @(); if ($f.autoNumber) { $note += 'オートナンバー' }; if ($f.description) { $note += $f.description }; if ($null -ne $f.defaultValue -and "$($f.defaultValue)") { $note += "既定値 $($f.defaultValue)" }
            & $w "| $(ConvertTo-MdCell $f.name) | $($f.type) | $($f.size) | $(if ($f.required) { '○' }) | $($key -join '・') | $(ConvertTo-MdCell ($note -join '、')) |"
        }
        $idx = @($t.indexes | Where-Object { -not $_.primary })
        if ($idx.Count) {
            & $w ''
            & $w "インデックス：$((@($idx | ForEach-Object { "$($_.name)（$((@($_.fields)) -join ', ')$(if ($_.unique) { '・重複なし' })）" })) -join '、')"
        }
    }
    return $sb.ToString()
}

function New-DocQueries($Db, $Xref, $Notes) {
    $sb = New-Object System.Text.StringBuilder
    $w = { param($s) [void]$sb.AppendLine($s) }
    & $w '# クエリ'
    & $w ''
    & $w '[← 概要](index.md)'
    foreach ($q in $Db.queries) {
        & $w ''
        & $w "<a id=""$(ConvertTo-MdAnchor 'query' $q.name)""></a>"
        & $w "## $($q.name)"
        & $w ''
        & $w (Format-DocNotes $Notes "query:$($q.name)")
        & $w ''
        $type = if ($q.sql -match '^\s*(INSERT|UPDATE|DELETE|SELECT[\s\S]*\bINTO\b)') { '更新系（アクション クエリ）' } elseif ($q.sql -match '^\s*TRANSFORM') { 'クロス集計' } else { '選択' }
        & $w "- 種類：$type"
        & $w "- 元にしているもの：$(Format-MdRefs $Xref.uses["query:$($q.name)"])"
        & $w "- 使われているところ：$(Format-MdRefs $Xref.usedBy["query:$($q.name)"] 'なし（見つかりません）')"
        & $w "- ファイル：``$($q.file)``"
        & $w ''
        & $w '```sql'
        & $w $q.sql
        & $w '```'
    }
    return $sb.ToString()
}

function New-DocFormsReports($Items, [string]$Kind, $Xref, $Notes) {
    $label = if ($Kind -eq 'form') { 'フォーム' } else { 'レポート' }
    $sb = New-Object System.Text.StringBuilder
    $w = { param($s) [void]$sb.AppendLine($s) }
    & $w "# $label"
    & $w ''
    & $w '[← 概要](index.md)'
    foreach ($o in $Items) {
        & $w ''
        & $w "<a id=""$(ConvertTo-MdAnchor $Kind $o.name)""></a>"
        & $w "## $($o.name)$(if ($o.caption) { "（$($o.caption)）" })"
        & $w ''
        & $w (Format-DocNotes $Notes "${Kind}:$($o.name)")
        & $w ''
        & $w "- 元データ：$(if ($o.recordSource) { '`' + (ConvertTo-MdCell $o.recordSource) + '`' } else { '（非連結）' })"
        if ($o.groups.Count) { & $w "- グループ：$((@($o.groups | ForEach-Object { "$($_.field)$(if ($_.header -and $_.footer) { '（ヘッダー・フッター）' } elseif ($_.header) { '（ヘッダー）' } elseif ($_.footer) { '（フッター）' })" })) -join ' → ')" }
        & $w "- 使っているもの：$(Format-MdRefs $Xref.uses["${Kind}:$($o.name)"])"
        & $w "- 使われているところ：$(Format-MdRefs $Xref.usedBy["${Kind}:$($o.name)"] 'なし（見つかりません）')"
        & $w "- ファイル：``$($o.file)``"
        $evs = @($o.events | ForEach-Object { "$($_.event)：$(if ($_.kind -eq 'procedure') { '`' + $_.handler + '`' } else { '`' + (ConvertTo-MdCell $_.value) + '`' })" })
        if ($evs.Count) { & $w "- $($label)のイベント：$($evs -join '、')" }
        $ctls = @($o.controls | Where-Object { $_.type -ne 'Label' -or $_.events.Count })
        if ($ctls.Count) {
            & $w ''
            & $w '| コントロール | 種類 | セクション | 中身（コントロールソース・値集合ソース） | イベント |'
            & $w '|---|---|---|---|---|'
            foreach ($c in $ctls) {
                $src = @(); if ($c.source) { $src += "``$(ConvertTo-MdCell $c.source)``" }; if ($c.rowSource) { $src += "値集合 ``$(ConvertTo-MdCell $c.rowSource)``" }
                if (-not $src -and $c.caption) { $src += "「$(ConvertTo-MdCell $c.caption)」" }
                $ev = @($c.events | ForEach-Object { "$($_.event)$(if ($_.kind -eq 'procedure') { '' } else { '=' + (ConvertTo-MdCell $_.value) })" })
                & $w "| $($c.name) | $($c.type) | $(ConvertTo-MdCell $c.section) | $($src -join '<br>') | $($ev -join '、') |"
            }
        }
        if (@($o.procs).Count) {
            & $w ''
            # どのコントロールのどのイベントで動くか
            $handlers = @{}
            foreach ($e in $o.events) { if ($e.handler) { $handlers[$e.handler.ToLower()] = "$($label)の $($e.event)" } }
            foreach ($c in $o.controls) { foreach ($e in $c.events) { if ($e.handler) { $handlers[$e.handler.ToLower()] = "$($c.name) の $($e.event)" } } }
            & $w '| コードの手続き | 動くとき | 説明 |'
            & $w '|---|---|---|'
            foreach ($p in $o.procs) {
                $when = if ($handlers.ContainsKey($p.name.ToLower())) { $handlers[$p.name.ToLower()] } else { '' }
                & $w "| ``$(ConvertTo-MdCell $p.signature)`` | $when | $(ConvertTo-MdCell $p.doc) |"
            }
        }
    }
    return $sb.ToString()
}

function New-DocModules($Db, $Xref, $Notes) {
    $sb = New-Object System.Text.StringBuilder
    $w = { param($s) [void]$sb.AppendLine($s) }
    & $w '# VBA モジュール'
    & $w ''
    & $w '[← 概要](index.md)'
    foreach ($m in $Db.modules) {
        & $w ''
        & $w "<a id=""$(ConvertTo-MdAnchor 'module' $m.name)""></a>"
        & $w "## $($m.name)$(if ($m.isClass) { '（クラス）' })"
        & $w ''
        & $w (Format-DocNotes $Notes "module:$($m.name)")
        & $w ''
        & $w "- 使っているもの：$(Format-MdRefs $Xref.uses["module:$($m.name)"])"
        & $w "- ファイル：``$($m.file)``"
        if (@($m.procs).Count) {
            & $w ''
            & $w '| 手続き | 種類 | 説明 | 呼ばれているところ |'
            & $w '|---|---|---|---|'
            foreach ($p in $m.procs) {
                # 同じモジュールのほかの手続きから呼ばれているか（自分の本体＝戻り値の代入などは除く）
                $codeLines = $m.code -split "`n"
                $endAt = $codeLines.Count - 1
                for ($k = $p.line; $k -lt $codeLines.Count; $k++) { if ($codeLines[$k] -match '^\s*End\s+(Sub|Function|Property)\b') { $endAt = $k; break } }
                $rest = @(for ($k = 0; $k -lt $codeLines.Count; $k++) { if ($k -lt $p.line - 1 -or $k -gt $endAt) { $codeLines[$k] } })
                $inModule = Test-NameUsed (($rest | Where-Object { $_ -notmatch '^\s*(?:(?:Public|Private|Friend|Global)\s+)?(?:Static\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s' }) -join "`n") $p.name
                $others = if ($p.scope -eq 'Private') { $null } else { $Xref.usedBy["procedure:$($p.name)"] }
                $calledBy = if ($null -ne $others -and $others.Count) { Format-MdRefs $others } else { '' }
                if ($inModule) { $calledBy = (@($calledBy, "$($m.name) の中") | Where-Object { $_ }) -join ' ／ ' }
                if (-not $calledBy) { $calledBy = 'なし（見つかりません）' }
                & $w "| ``$(ConvertTo-MdCell $p.signature)`` | $($p.scope) $($p.kind) | $(ConvertTo-MdCell $p.doc) | $calledBy |"
            }
        }
    }
    return $sb.ToString()
}

# 入口：ソースフォルダからドキュメントを作る
function Invoke-SourceDocument([string]$Src, [string]$OutDir, [string]$DbName) {
    if (-not $OutDir) { $OutDir = ($Src -replace '\.src$', '') + '.docs' }
    if (-not (Test-Path -LiteralPath $OutDir)) { [void](New-Item -ItemType Directory -Path $OutDir -Force) }
    $notes = Read-DocNotes $OutDir
    $db = Read-DatabaseStructure $Src
    $xref = Get-DatabaseXref $db
    if (-not $DbName) { $DbName = [System.IO.Path]::GetFileName(($Src -replace '\.src$', '')) }
    $files = [ordered]@{
        'index.md'   = (New-DocIndex $db $xref $notes $DbName)
        'tables.md'  = (New-DocTables $db $xref $notes)
        'queries.md' = (New-DocQueries $db $xref $notes)
        'forms.md'   = (New-DocFormsReports $db.forms 'form' $xref $notes)
        'reports.md' = (New-DocFormsReports $db.reports 'report' $xref $notes)
        'modules.md' = (New-DocModules $db $xref $notes)
    }
    $written = New-Object System.Collections.Generic.List[string]
    foreach ($k in $files.Keys) {
        $path = Join-Path $OutDir $k
        Write-SourceText $path ($files[$k] -replace "`r?`n", "`r`n")
        $written.Add($path)
    }
    $keptNotes = @($notes.Keys)
    return [ordered]@{
        outDir  = $OutDir
        index   = (Join-Path $OutDir 'index.md')
        files   = $written.ToArray()
        counts  = [ordered]@{ tables = @($db.tables).Count; queries = @($db.queries).Count; forms = @($db.forms).Count; reports = @($db.reports).Count; modules = @($db.modules).Count; macros = @($db.macros).Count }
        notesKept = $keptNotes.Count
        emptyNotes = @(@('database') + @($db.tables | ForEach-Object { "table:$($_.name)" }) + @($db.queries | ForEach-Object { "query:$($_.name)" }) + @($db.forms | ForEach-Object { "form:$($_.name)" }) + @($db.reports | ForEach-Object { "report:$($_.name)" }) + @($db.modules | ForEach-Object { "module:$($_.name)" }) | Where-Object { $keptNotes -notcontains $_ })
    }
}
