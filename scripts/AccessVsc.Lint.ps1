<#
  AccessVsc.Lint.ps1
  Access を起動せずにソースフォルダをチェックする（CI のクラウド実行環境や commit 前フックで使える）。

  ルール:
    sjis               (error)   Shift-JIS で表せない文字（Access に取り込むと ? になる）
    option-explicit    (warning) Option Explicit がない
    stop-statement     (warning) Stop 文が残っている（実行が止まる）
    layout-width       (warning) コントロールの右端がレポート/フォームの幅を超えている（取り込むと幅が自動で広がる）
    layout-section     (warning) コントロールの下端がセクションの高さを超えている
    empty-sql          (error)   クエリの SQL が空
    readonly-edited    (warning) 読み取り専用のスナップショット（tables/*.json）が編集されている
    implicit-variant   (warning) 変数・引数の型（As ...）が書かれていない（暗黙の Variant）
    implicit-variant-return (info) Function / Property Get の戻り値の型が書かれていない
    unused-variable    (warning) 手続きの中で宣言したのに使っていない変数
    broken-reference   (warning) VBA の参照設定が「参照不可」になっている（references.json）
    block-structure    (warning) If / For / Do / With / Select などのブロックの対応が取れない（End If の書き忘れなど）
    format             (info)    インデントが整っていない（access.ps1 format / エディタの Shift+Alt+F で直せる）
#>

$script:SectionTypes = @('Section', 'PageHeader', 'PageFooter', 'FormHeader', 'FormFooter', 'BreakHeader', 'BreakFooter')

function New-LintProblem([string]$File, [int]$Line, [string]$Severity, [string]$Rule, [string]$Message) {
    return [ordered]@{ file = $File; line = $Line; severity = $Severity; rule = $Rule; message = $Message }
}

function Test-VbaCode([string]$Rel, [string[]]$Lines, [int]$Offset, $Problems) {
    # $Offset: ファイル内でコードが始まる行（0 始まり）
    $code = @($Lines | Select-Object -Skip $Offset)
    $hasCode = @($code | Where-Object { $_ -match '^\s*(Public|Private|Friend|Sub|Function|Property|Dim)\b' }).Count -gt 0
    if ($hasCode -and -not (@($code | Where-Object { $_ -match '^\s*Option\s+Explicit\b' }).Count)) {
        $Problems.Add((New-LintProblem $Rel ($Offset + 1) 'warning' 'option-explicit' 'Option Explicit がありません。変数の書き間違いがコンパイルで見つからなくなります。'))
    }
    for ($i = 0; $i -lt $code.Count; $i++) {
        if ($code[$i] -match '^\s*Stop\s*(''.*)?$') {
            $Problems.Add((New-LintProblem $Rel ($Offset + $i + 1) 'warning' 'stop-statement' 'Stop 文が残っています。実行がそこで止まります。'))
        }
    }
    Test-VbaDeclarations $Rel $code $Offset $Problems
}

# 文字列リテラルとコメントを空白にする（位置はそのまま）
function Remove-VbaNoise([string]$Line) {
    $sb = New-Object System.Text.StringBuilder
    $inStr = $false
    for ($i = 0; $i -lt $Line.Length; $i++) {
        $ch = $Line[$i]
        if ($inStr) {
            if ($ch -eq '"') {
                if ($i + 1 -lt $Line.Length -and $Line[$i + 1] -eq '"') { [void]$sb.Append('  '); $i++; continue }
                $inStr = $false; [void]$sb.Append('"'); continue
            }
            [void]$sb.Append(' '); continue
        }
        if ($ch -eq '"') { $inStr = $true; [void]$sb.Append('"'); continue }
        if ($ch -eq "'") { [void]$sb.Append(' ', $Line.Length - $i); break }
        [void]$sb.Append($ch)
    }
    $s = $sb.ToString()
    if ($s -match '^\s*Rem\b') { return (' ' * $Line.Length) }
    return $s
}

function Split-TopLevel([string]$Text) {
    $out = New-Object System.Collections.Generic.List[string]
    $depth = 0; $cur = ''
    foreach ($ch in $Text.ToCharArray()) {
        if ($ch -eq '(') { $depth++ } elseif ($ch -eq ')') { $depth-- }
        if ($ch -eq ',' -and $depth -eq 0) { $out.Add($cur.Trim()); $cur = ''; continue }
        $cur += $ch
    }
    if ($cur.Trim()) { $out.Add($cur.Trim()) }
    return , $out.ToArray()
}

# 型の書き忘れ（暗黙の Variant）・使っていない変数
function Test-VbaDeclarations([string]$Rel, [string[]]$Code, [int]$Offset, $Problems) {
    # 行継続（ _）をつないだ論理行
    $logical = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt $Code.Count; $i++) {
        $start = $i
        $s = Remove-VbaNoise $Code[$i]
        while ($s -match '\s_\s*$' -and $i + 1 -lt $Code.Count) { $i++; $s = ($s -replace '\s_\s*$', ' ') + (Remove-VbaNoise $Code[$i]).Trim() }
        [void]$logical.Add(@{ line = $start; text = $s })
    }
    $id = '[\p{L}_][\p{L}\p{Nd}_]*'
    $hasType = { param($part) ($part -match '(?i)\bAs\s+\S') -or ($part -match '^[\p{L}_][\p{L}\p{Nd}_]*[%&!#@$]') }
    $inType = $false
    $proc = $null
    foreach ($l in $logical) {
        $t = $l.text
        $lineNo = $Offset + $l.line + 1
        if ($t -match '(?i)^\s*(Public\s+|Private\s+)?(Type|Enum)\s') { $inType = $true; continue }
        if ($inType) { if ($t -match '(?i)^\s*End\s+(Type|Enum)\b') { $inType = $false }; continue }

        $m = [regex]::Match($t, "(?i)^\s*(?:(?:Public|Private|Friend|Global)\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+($id[%&!#@$]?)\s*\((.*)$")
        if ($m.Success) {
            $proc = @{ name = $m.Groups[2].Value; start = $l.line; vars = New-Object System.Collections.ArrayList; body = New-Object System.Collections.ArrayList }
            $rest = $m.Groups[3].Value
            $depth = 1; $close = -1
            for ($k = 0; $k -lt $rest.Length; $k++) { if ($rest[$k] -eq '(') { $depth++ } elseif ($rest[$k] -eq ')') { $depth--; if ($depth -eq 0) { $close = $k; break } } }
            $params = if ($close -ge 0) { $rest.Substring(0, $close) } else { $rest }
            foreach ($p in (Split-TopLevel $params)) {
                if (-not $p) { continue }
                $clean = $p -replace '(?i)^(Optional\s+)?(ByVal\s+|ByRef\s+)?', ''
                if ($clean -match '(?i)^ParamArray\b') { continue }
                $clean = ($clean -split '=')[0].Trim()
                if (-not (& $hasType $clean)) {
                    $pn = [regex]::Match($clean, "^$id").Value
                    $Problems.Add((New-LintProblem $Rel $lineNo 'warning' 'implicit-variant' "引数 $pn の型が書かれていません（暗黙の Variant）。As Long などを付けると、間違った値を渡したときにコンパイルで気づけます。"))
                }
            }
            if ($m.Groups[1].Value -match '(?i)^(Function|Property\s+Get)$' -and $m.Groups[2].Value -notmatch '[%&!#@$]$') {
                $after = if ($close -ge 0) { $rest.Substring($close + 1) } else { '' }
                if ($after -notmatch '(?i)^\s*As\s+\S') {
                    $Problems.Add((New-LintProblem $Rel $lineNo 'info' 'implicit-variant-return' "$($m.Groups[2].Value) の戻り値の型が書かれていません（暗黙の Variant）。"))
                }
            }
            continue
        }
        if ($proc -and $t -match '(?i)^\s*End\s+(Sub|Function|Property)\b') {
            # 使っていない変数：宣言した行以外に名前が出てこない
            foreach ($v in $proc.vars) {
                $re = '(?i)(?<![\p{L}\p{Nd}_])' + [regex]::Escape($v.name) + '(?![\p{L}\p{Nd}_])'
                $used = $false
                foreach ($b in $proc.body) {
                    # 宣言と同じ行（Dim i As Long: For i = ...）は 2 回以上出てきたら使っている
                    $n = [regex]::Matches($b.text, $re).Count
                    if (($b.line -ne $v.line -and $n -gt 0) -or ($b.line -eq $v.line -and $n -gt 1)) { $used = $true; break }
                }
                if (-not $used) {
                    $Problems.Add((New-LintProblem $Rel ($Offset + $v.line + 1) 'warning' 'unused-variable' "変数 $($v.name) は宣言されていますが使われていません。"))
                }
            }
            $proc = $null
            continue
        }
        if ($proc) { [void]$proc.body.Add(@{ line = $l.line; text = $t }) }

        # 変数の宣言（モジュール先頭・手続きの中）
        $d = [regex]::Match($t, '(?i)^\s*(Dim|Static|Private|Public|Global)\s+(?!Const\b|Sub\b|Function\b|Property\b|Declare\b|Type\b|Enum\b|Event\b|WithEvents\b)(.*)$')
        if (-not $d.Success) { $d = [regex]::Match($t, '(?i)^\s*(Dim|Private|Public)\s+WithEvents\s+(.*)$') }
        if ($d.Success) {
            foreach ($part in (Split-TopLevel $d.Groups[2].Value)) {
                if (-not $part) { continue }
                $nm = [regex]::Match($part, "^$id").Value
                if (-not $nm) { continue }
                if (-not (& $hasType $part)) {
                    $Problems.Add((New-LintProblem $Rel $lineNo 'warning' 'implicit-variant' "変数 $nm の型が書かれていません（暗黙の Variant）。Dim a, b As Long と書くと a は Variant になる点にも注意。"))
                }
                if ($proc -and $d.Groups[1].Value -match '(?i)^(Dim|Static)$') { [void]$proc.vars.Add(@{ name = $nm; line = $l.line }) }
            }
        }
    }
}

function Test-ShiftJis([string]$Rel, [string[]]$Lines, $Problems) {
    $enc = [System.Text.Encoding]::GetEncoding(932)
    for ($i = 0; $i -lt $Lines.Count; $i++) {
        $l = $Lines[$i]
        if ($l -match '^[\x00-\x7F]*$') { continue }
        if ($enc.GetString($enc.GetBytes($l)) -cne $l) {
            # 絵文字などのサロゲートペアを 1 文字として扱う
            $chars = @([regex]::Matches($l, '[\uD800-\uDBFF][\uDC00-\uDFFF]|.') | ForEach-Object { $_.Value })
            $bad = @($chars | Where-Object { $enc.GetString($enc.GetBytes($_)) -cne $_ } | Select-Object -Unique) -join ' '
            $Problems.Add((New-LintProblem $Rel ($i + 1) 'error' 'sjis' "Shift-JIS で表せない文字があります（$bad）。Access に取り込むと ? に変わります。"))
        }
    }
}

function Test-ObjectLayoutText([string]$Rel, [string[]]$Lines, $Problems) {
    # SaveAsText の Begin/End 構造をたどり、はみ出しを調べる（単位 twip）
    $stack = New-Object System.Collections.ArrayList
    $width = $null
    $binary = $false
    for ($i = 0; $i -lt $Lines.Count; $i++) {
        $t = $Lines[$i].Trim()
        if ($t -match '^CodeBehind(Form|Report)$') { break }
        if ($binary) { if ($t -eq 'End') { $binary = $false }; continue }
        if ($t -match '^Begin(\s+(\w+))?$') {
            [void]$stack.Add(@{ type = $Matches[2]; props = @{}; line = $i + 1 })
            continue
        }
        if ($t -eq 'End') {
            if ($stack.Count -eq 0) { continue }
            $b = $stack[$stack.Count - 1]
            $stack.RemoveAt($stack.Count - 1)
            if ($b.type -in @('Report', 'Form')) { continue }
            if (-not $b.props.ContainsKey('Name') -or -not $b.props.ContainsKey('Width')) { continue }
            # 親をたどってセクションを探す（見つからなければ既定値ブロックなので対象外）
            $section = $null
            for ($k = $stack.Count - 1; $k -ge 0; $k--) { if ($script:SectionTypes -contains $stack[$k].type) { $section = $stack[$k]; break } }
            if (-not $section) { continue }
            $name = $b.props['Name'].Trim('"')
            $left = 0; if ($b.props.ContainsKey('Left')) { $left = [int]$b.props['Left'] }
            $top = 0; if ($b.props.ContainsKey('Top')) { $top = [int]$b.props['Top'] }
            $w = [int]$b.props['Width']
            $h = 0; if ($b.props.ContainsKey('Height')) { $h = [int]$b.props['Height'] }
            if ($width -and ($left + $w) -gt $width) {
                $Problems.Add((New-LintProblem $Rel $b.line 'warning' 'layout-width' ("{0} の右端 {1:0.00}cm が幅 {2:0.00}cm を超えています（取り込むと幅が自動で広がり、用紙に入らなくなることがあります）。" -f $name, (($left + $w) / 567), ($width / 567))))
            }
            if ($section.props.ContainsKey('Height') -and ($top + $h) -gt [int]$section.props['Height']) {
                $Problems.Add((New-LintProblem $Rel $b.line 'warning' 'layout-section' ("{0} の下端 {1:0.00}cm がセクションの高さ {2:0.00}cm を超えています。" -f $name, (($top + $h) / 567), ([int]$section.props['Height'] / 567))))
            }
            continue
        }
        if ($t -match '^(\w+)\s*=\s*(.*)$') {
            $key = $Matches[1]; $val = $Matches[2]
            if ($val -eq 'Begin') { $binary = $true; continue }   # PrtMip = Begin ... End などのバイナリ
            if ($stack.Count -gt 0) {
                $cur = $stack[$stack.Count - 1]
                if (-not $cur.props.ContainsKey($key)) { $cur.props[$key] = $val }
                if ($stack.Count -eq 1 -and $key -eq 'Width' -and $cur.type -in @('Report', 'Form')) { $width = [int]$val }
            }
        }
    }
}

function Test-VbaFormat([string]$Rel, [string[]]$Lines, $Problems) {
    # 整形の規則（AccessVsc.Format.ps1）と比べる
    if (-not (Get-Command Format-VbaText -ErrorAction SilentlyContinue)) { return }
    $new = Format-VbaText ($Lines -join "`n") $Rel
    if ($null -eq $new) {
        $Problems.Add((New-LintProblem $Rel 1 'warning' 'block-structure' 'If / For / Do / With / Select Case / Sub などのブロックの対応が取れません（End If・Next・Loop・End Sub などの書き忘れ）。コンパイルエラーになります。'))
        return
    }
    $nl = $new -split "`n"
    for ($i = 0; $i -lt [Math]::Min($nl.Count, $Lines.Count); $i++) {
        if ($nl[$i] -cne $Lines[$i]) {
            $Problems.Add((New-LintProblem $Rel ($i + 1) 'info' 'format' 'インデントが整っていません（ここから下）。整形（エディタの Shift+Alt+F、または access.ps1 format）で直せます。'))
            return
        }
    }
}

function Invoke-SourceLint([string]$Src, [string[]]$Only) {
    $problems = New-Object System.Collections.Generic.List[object]
    $files = @(Get-SourceFiles $Src)
    if ($Only) { $files = @($files | Where-Object { $Only -contains $_ }) }
    $status = $null
    foreach ($rel in $files) {
        $info = Get-ObjectInfoFromRelPath $rel
        if (-not $info) { continue }
        $full = Join-Path $Src ($rel.Replace('/', '\'))
        $lines = @((Read-SourceText $full) -split "`r?`n")
        switch ($info.kind) {
            'modules' {
                Test-ShiftJis $rel $lines $problems
                $off = 0
                while ($off -lt $lines.Count -and $lines[$off] -match '^(Attribute VB_|VERSION |BEGIN|\s+MultiUse|END$)') { $off++ }
                Test-VbaCode $rel $lines $off $problems
                Test-VbaFormat $rel $lines $problems
            }
            { $_ -in @('forms', 'reports') } {
                Test-ShiftJis $rel $lines $problems
                Test-ObjectLayoutText $rel $lines $problems
                $cb = [Array]::FindIndex([string[]]$lines, [Predicate[string]] { param($l) $l.Trim() -match '^CodeBehind(Form|Report)$' })
                if ($cb -ge 0) { Test-VbaCode $rel $lines ($cb + 1) $problems; Test-VbaFormat $rel $lines $problems }
            }
            'macros' { Test-ShiftJis $rel $lines $problems }
            'queries' {
                Test-ShiftJis $rel $lines $problems
                $sql = (@($lines | Where-Object { $_ -notmatch '^--\s*accessvsc:' }) -join "`n").Trim()
                if (-not $sql) { $problems.Add((New-LintProblem $rel 1 'error' 'empty-sql' 'SQL が空です。')) }
            }
            default {
                if ($info.kind -eq 'references') {
                    try {
                        $refs = (Read-SourceText $full | ConvertFrom-Json).references
                        foreach ($r in @($refs)) {
                            if ($r.isBroken) {
                                $problems.Add((New-LintProblem $rel 1 'warning' 'broken-reference' ("VBA の参照設定「{0}」が参照不可です（{1}）。この PC にないライブラリを参照していると、コンパイルできません。Access の［ツール］→［参照設定］で外すか、ライブラリを入れてください。" -f $r.name, $r.fullPath)))
                            }
                        }
                    }
                    catch {}
                }
                if ($info.readOnly) {
                    if (-not $status) { $status = Get-SyncStatus $Src }
                    if (@($status.readOnlyChanged) -contains $rel) {
                        $problems.Add((New-LintProblem $rel 1 'warning' 'readonly-edited' '読み取り専用のスナップショットです。編集しても DB は変わりません（テーブル変更は SQL の DDL で行い、エクスポートし直してください）。'))
                    }
                }
            }
        }
    }
    $errors = @($problems | Where-Object { $_.severity -eq 'error' }).Count
    $infos = @($problems | Where-Object { $_.severity -eq 'info' }).Count
    return [ordered]@{
        files        = $files.Count
        errorCount   = $errors
        warningCount = $problems.Count - $errors - $infos
        infoCount    = $infos
        problems     = $problems.ToArray()
    }
}
