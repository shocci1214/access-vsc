<#
  AccessVsc.Format.ps1
  VBA の整形（インデントと行末の空白だけ）。vba-format.js（エディタの Shift+Alt+F）と同じ規則の PowerShell 版。
  Access も DB も使わない。ブロックの対応が取れないファイル（End If の書き忘れなど）は変えずに報告する。
  テスト: test/format/*.in.bas → *.out.bas（JS 版と共通。test/Test-Format.ps1）
#>

$script:FmtId = '[A-Za-z_\u0080-￿][\w\u0080-￿]*'
$script:FmtOpt = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
$script:FmtProc = New-Object regex ("^(?:(?:Public|Private|Friend|Global)\s+)?(?:Static\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+$($script:FmtId)\s*\("), $script:FmtOpt
$script:FmtType = New-Object regex ("^(?:(?:Public|Private)\s+)?(?:Type|Enum)\s+$($script:FmtId)$"), $script:FmtOpt
$script:FmtLabel = New-Object regex ("^($($script:FmtId)|\d+):$")

function Get-VbaBlockKind([string]$S) {
    # $S: コメントと文字列を除き、前後の空白を取った論理行
    if ($S -match '^End\s+(Sub|Function|Property)\b') { return @{ close = 'proc' } }
    if ($S -match '^End\s+If\b') { return @{ close = 'if' } }
    if ($S -match '^End\s+Select\b') { return @{ close = 'select' } }
    if ($S -match '^End\s+With\b') { return @{ close = 'with' } }
    if ($S -match '^End\s+(Type|Enum)\b') { return @{ close = 'type' } }
    if ($S -match '^Next(\s|$)') {
        $rest = ($S -replace '^Next\s*', '').Trim()
        return @{ close = 'for'; count = $(if ($rest) { $rest.Split(',').Count } else { 1 }) }
    }
    if ($S -match '^Loop(\s|$)') { return @{ close = 'do' } }
    if ($S -match '^Wend$') { return @{ close = 'while' } }
    if ($S -match '^Else$') { return @{ middle = 'if' } }
    if ($S -match '^ElseIf\b.*\bThen$') { return @{ middle = 'if' } }
    if ($S -match '^Case(\s|$)') { return @{ kase = $true } }
    if ($script:FmtProc.IsMatch($S)) { return @{ open = 'proc' } }
    if ($script:FmtType.IsMatch($S)) { return @{ open = 'type' } }
    if ($S -match '^If\b.*\bThen$') { return @{ open = 'if' } }
    if ($S -match '^Select\s+Case\b') { return @{ open = 'select' } }
    if ($S -match '^For(\s|$)') { return @{ open = 'for' } }
    if ($S -match '^Do(\s|$)') { return @{ open = 'do' } }
    if ($S -match '^While\s') { return @{ open = 'while' } }
    if ($S -match '^With\s') { return @{ open = 'with' } }
    return @{}
}

function Get-VbaCodeStart([string[]]$Lines, [string]$File) {
    if ($File -match '\.txt$') {
        for ($i = 0; $i -lt $Lines.Count; $i++) { if ($Lines[$i].Trim() -match '^CodeBehind(Form|Report)$') { return $i + 1 } }
        return $Lines.Count
    }
    $start = 0
    for ($i = 0; $i -lt [Math]::Min($Lines.Count, 40); $i++) { if ($Lines[$i] -match '^Attribute VB_') { $start = $i + 1 } }
    return $start
}

function Format-VbaText([string]$Text, [string]$File = '', [string]$Indent = '    ') {
    # 整形した全文を返す。ブロックの対応が取れなければ $null
    $eol = if ($Text.Contains("`r`n")) { "`r`n" } else { "`n" }
    $lines = [string[]]($Text -split "`r?`n")
    $start = Get-VbaCodeStart $lines $File
    $out = New-Object System.Collections.Generic.List[string]
    for ($i = 0; $i -lt $start; $i++) { $out.Add($lines[$i]) }
    $stack = New-Object System.Collections.Generic.List[string]
    $top = { if ($stack.Count) { $stack[$stack.Count - 1] } else { '' } }
    $width = { param($ws) $ws.Replace("`t", '    ').Length }
    for ($i = $start; $i -lt $lines.Count; $i++) {
        $raw = $lines[$i] -replace '[ \t]+$', ''
        $stripped = Remove-VbaNoise $raw
        $s0 = $stripped.Trim()
        if (-not $raw.Trim()) { $out.Add(''); continue }
        $end = $i
        $logical = $stripped
        while ((Remove-VbaNoise $lines[$end]) -match '\s_\s*$' -and $end + 1 -lt $lines.Count) {
            $end++
            $logical = ($logical -replace '\s_\s*$', ' ') + ' ' + (Remove-VbaNoise $lines[$end]).Trim()
        }
        $s = $logical.Trim()
        $keep = ($s0 -match '^#') -or ($s0 -match '^Attribute\s')
        $lab = $script:FmtLabel.Match($s)
        if ($keep) { $depth = $null }
        elseif (-not $s0) { $depth = $stack.Count }
        elseif ($lab.Success -and $lab.Groups[1].Value -notmatch '^(Else|Case|Resume|GoTo|Exit|End|Next|Loop|Wend)$') { $depth = -1 }
        else {
            $c = Get-VbaBlockKind $s
            if ($c.close) {
                if ($c.close -eq 'select' -and (& $top) -eq 'case') { $stack.RemoveAt($stack.Count - 1) }
                $n = if ($c.count) { $c.count } else { 1 }
                for ($k = 0; $k -lt $n; $k++) {
                    if ((& $top) -ne $c.close) { return $null }
                    $stack.RemoveAt($stack.Count - 1)
                }
                $depth = $stack.Count
            }
            elseif ($c.middle) {
                if ((& $top) -ne 'if') { return $null }
                $depth = $stack.Count - 1
            }
            elseif ($c.kase) {
                if ((& $top) -eq 'case') { $stack.RemoveAt($stack.Count - 1) }
                if ((& $top) -ne 'select') { return $null }
                $depth = $stack.Count
                $stack.Add('case')
            }
            else {
                $depth = $stack.Count
                if ($c.open) { $stack.Add($c.open) }
            }
        }
        $oldIndent = ([regex]'^[ \t]*').Match($raw).Value
        $newIndent = if ($null -eq $depth) { $oldIndent } elseif ($depth -lt 0) { '' } else { $Indent * $depth }
        $out.Add($newIndent + $raw.Substring($oldIndent.Length))
        $delta = (& $width $newIndent) - (& $width $oldIndent)
        for ($j = $i + 1; $j -le $end; $j++) {
            $r = $lines[$j] -replace '[ \t]+$', ''
            $ws = ([regex]'^[ \t]*').Match($r).Value
            $w = [Math]::Max(0, (& $width $ws) + $delta)
            $pad = if ($Indent -eq "`t") { ("`t" * [Math]::Floor($w / 4)) + (' ' * ($w % 4)) } else { ' ' * $w }
            $out.Add($pad + $r.Substring($ws.Length))
        }
        $i = $end
    }
    if ($stack.Count) { return $null }
    return ($out -join $eol)
}

function Invoke-SourceFormat([string]$Src, [string[]]$Only, [bool]$Check) {
    # VBA を含むファイル（modules/*.bas|cls、forms・reports の .txt）を整形する。$Check のときは書き込まずに一覧だけ返す
    $changed = New-Object System.Collections.Generic.List[string]
    $skipped = New-Object System.Collections.Generic.List[object]
    $files = @(Get-SourceFiles $Src | Where-Object { $_ -match '^(modules/.+\.(bas|cls)|(forms|reports)/.+\.txt)$' })
    if ($Only) {
        $want = @($Only | ForEach-Object { $_.Replace('\', '/') })
        $files = @($files | Where-Object { $want -contains $_ })
    }
    foreach ($rel in $files) {
        $full = Join-Path $Src ($rel.Replace('/', '\'))
        $bytes = [System.IO.File]::ReadAllBytes($full)
        $bom = $bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF
        $text = Read-SourceText $full
        $new = Format-VbaText $text $rel
        if ($null -eq $new) { $skipped.Add([ordered]@{ file = $rel; reason = 'ブロックの対応が取れません（End If・Next・End Sub などの書き忘れ）。このファイルは変えていません。' }); continue }
        if ($new -ceq $text) { continue }
        $changed.Add($rel)
        if (-not $Check) { [System.IO.File]::WriteAllText($full, $new, (New-Object System.Text.UTF8Encoding($bom))) }
    }
    return [ordered]@{ files = $files.Count; changed = [string[]]$changed.ToArray(); skipped = [object[]]$skipped.ToArray(); check = $Check }
}
