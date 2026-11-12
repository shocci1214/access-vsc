<#
  AccessVsc.Ops.ps1
  Access を COM で起動して実際の操作を行う部分。
  Runner が用意した STA ランスペース（別スレッド）の中で実行される。
  ダイアログで止まった場合は Runner 側の見張り役が対処する。
#>

$ErrorActionPreference = 'Stop'

# DAO のフィールド型番号 -> 名前
$script:DaoTypes = @{
    1 = 'Boolean'; 2 = 'Byte'; 3 = 'Integer'; 4 = 'Long'; 5 = 'Currency'; 6 = 'Single'; 7 = 'Double'
    8 = 'Date'; 9 = 'Binary'; 10 = 'Text'; 11 = 'OLEObject'; 12 = 'Memo'; 15 = 'GUID'; 16 = 'BigInt'
    17 = 'VarBinary'; 18 = 'Char'; 19 = 'Numeric'; 20 = 'Decimal'; 21 = 'Float'; 22 = 'Time'; 23 = 'TimeStamp'
    101 = 'Attachment'; 102 = 'ComplexByte'; 103 = 'ComplexInteger'; 104 = 'ComplexLong'; 105 = 'ComplexSingle'
    106 = 'ComplexDouble'; 107 = 'ComplexGUID'; 108 = 'ComplexDecimal'; 109 = 'ComplexText'
}
$script:ClassHeader = @(
    'Attribute VB_GlobalNameSpace = False'
    'Attribute VB_Creatable = False'
    'Attribute VB_PredeclaredId = False'
    'Attribute VB_Exposed = False'
)

function Get-InnerMessage($err) {
    $e = $err
    if ($e -is [System.Management.Automation.ErrorRecord]) { $e = $e.Exception }
    while ($e.InnerException) { $e = $e.InnerException }
    $msg = $e.Message
    if ($e -is [System.Runtime.InteropServices.COMException]) { $msg = "$msg (0x{0:X8})" -f $e.HResult }
    return $msg
}

function ConvertTo-PlainValue($v) {
    # COM / DAO の値を JSON にできる形へ
    if ($null -eq $v -or $v -is [System.DBNull]) { return $null }
    if ($v -is [datetime]) {
        if ($v.TimeOfDay.Ticks -eq 0) { return $v.ToString('yyyy-MM-dd') }
        return $v.ToString('yyyy-MM-ddTHH:mm:ss')
    }
    if ($v -is [byte[]]) { return "<binary $($v.Length) bytes>" }
    if ($v -is [System.__ComObject]) { return '<object>' }
    if ($v -is [array]) {
        if ($v.Rank -eq 2) {
            $rows = New-Object System.Collections.ArrayList
            for ($r = $v.GetLowerBound(1); $r -le $v.GetUpperBound(1); $r++) {
                $row = @()
                for ($c = $v.GetLowerBound(0); $c -le $v.GetUpperBound(0); $c++) { $row += , (ConvertTo-PlainValue $v[$c, $r]) }
                [void]$rows.Add($row)
            }
            return , $rows.ToArray()
        }
        $list = New-Object System.Collections.ArrayList
        foreach ($x in $v) { [void]$list.Add((ConvertTo-PlainValue $x)) }
        return , $list.ToArray()
    }
    if ($v -is [decimal]) { return [double]$v }
    return $v
}

function Hide-Password([string]$Connect) {
    if (-not $Connect) { return $Connect }
    return [regex]::Replace($Connect, '(?i)\b(PWD|Password)=[^;]*', '$1=***')
}

function New-TempDir {
    $d = Join-Path ([System.IO.Path]::GetTempPath()) ('accessvsc_' + [guid]::NewGuid().ToString('N'))
    [void](New-Item -ItemType Directory -Path $d -Force)
    return $d
}

function Get-VbComponentTypes($app) {
    $types = @{}
    try {
        foreach ($c in $app.VBE.ActiveVBProject.VBComponents) { $types[$c.Name] = [int]$c.Type }
    }
    catch {}
    return $types
}

# ============================================================
# コンパイル
#   1) 全モジュールをコンパイルして保存（非表示ならダイアログは出ない）
#   2) 失敗したら VBE の「コンパイル」ボタンを実行し、カーソル位置からエラー箇所を得る
# ============================================================
function Invoke-Compile($app, [bool]$Save = $true) {
    # Save = $false のときは保存しない（実行前の確認など、読み取りだけの操作で DB を書き換えないため）
    if ($Save) { try { $app.RunCommand(126) } catch {} }   # acCmdCompileAndSaveAllModules
    else { try { $app.RunCommand(125) } catch {} }        # acCmdCompileAllModules
    if ($app.IsCompiled) { return [ordered]@{ compiled = $true; error = $null } }

    $err = [ordered]@{ module = $null; line = $null; column = $null; code = $null; message = 'Compile error' }
    try {
        $ctl = $app.VBE.CommandBars.FindControl(1, 578)   # [デバッグ] - [コンパイル]
        if ($ctl) { $ctl.Execute() }
        $pane = $app.VBE.ActiveCodePane
        if ($pane) {
            $sl = 0; $sc = 0; $el = 0; $ec = 0
            $pane.GetSelection([ref]$sl, [ref]$sc, [ref]$el, [ref]$ec)
            $err.module = $pane.CodeModule.Parent.Name
            $err.line = $sl
            $err.column = $sc
            if ($sl -ge 1) { $err.code = $pane.CodeModule.Lines($sl, 1) }
            if ($sc -ge 1 -and $ec -gt $sc -and $err.code) {
                $len = [Math]::Min($ec - $sc, $err.code.Length - $sc + 1)
                if ($len -gt 0) { $err.message = "Compile error near '" + $err.code.Substring($sc - 1, $len) + "'" }
            }
        }
    }
    catch { $err.message = 'Compile error (location unavailable): ' + (Get-InnerMessage $_) }
    return [ordered]@{ compiled = $false; error = $err }
}

# ============================================================
# エクスポート（Access -> ソースフォルダ）
# ============================================================
function Export-TableSchema($db, $td) {
    $fields = New-Object System.Collections.ArrayList
    foreach ($f in $td.Fields) {
        $desc = $null
        try { $desc = $f.Properties.Item('Description').Value } catch {}
        $t = [int]$f.Type
        $fo = [ordered]@{
            name     = $f.Name
            type     = $(if ($script:DaoTypes.ContainsKey($t)) { $script:DaoTypes[$t] } else { "Type$t" })
            size     = [int]$f.Size
            required = [bool]$f.Required
        }
        if (($f.Attributes -band 16) -ne 0) { $fo.autoNumber = $true }
        if ($t -eq 10 -or $t -eq 12) { $fo.allowZeroLength = [bool]$f.AllowZeroLength }
        if ($f.DefaultValue) { $fo.defaultValue = [string]$f.DefaultValue }
        if ($f.ValidationRule) { $fo.validationRule = [string]$f.ValidationRule }
        if ($desc) { $fo.description = [string]$desc }
        [void]$fields.Add($fo)
    }
    $indexes = New-Object System.Collections.ArrayList
    try {
        foreach ($ix in $td.Indexes) {
            $names = @(); foreach ($ixf in $ix.Fields) { $names += $ixf.Name }
            [void]$indexes.Add([ordered]@{ name = $ix.Name; primary = [bool]$ix.Primary; unique = [bool]$ix.Unique; fields = $names })
        }
    }
    catch {}
    $o = [ordered]@{
        '$note' = 'Read-only schema snapshot generated by Access VSC. Editing this file does not change the database.'
        name    = $td.Name
    }
    if ($td.Connect) {
        $o.linked = $true
        $o.connect = Hide-Password $td.Connect
        $o.sourceTable = $td.SourceTableName
    }
    else {
        $o.linked = $false
        try { $o.recordCount = [int]$td.RecordCount } catch {}
    }
    $desc = $null
    try { $desc = $td.Properties.Item('Description').Value } catch {}
    if ($desc) { $o.description = [string]$desc }
    $o.fields = $fields.ToArray()
    $o.indexes = $indexes.ToArray()
    return $o
}

function Invoke-ExportOp($app, $P) {
    $src = $P.src
    $tmp = New-TempDir
    $files = [ordered]@{}            # relPath -> @{ text; encoding }
    $failed = New-Object System.Collections.Generic.List[string]
    $warnings = New-Object System.Collections.Generic.List[string]
    $kinds = if ($P.kinds) { @($P.kinds) } else { @('modules', 'forms', 'reports', 'macros', 'queries', 'tables') }

    function Save-One($acType, $name, $relPath) {
        $t = Join-Path $tmp ([guid]::NewGuid().ToString('N') + '.txt')
        try {
            $app.SaveAsText($acType, $name, $t)
            $r = Read-TextAuto $t
            $files[$relPath] = @{ text = $r.text; encoding = $r.encoding }
        }
        catch {
            $failed.Add($relPath)
            $warnings.Add("Export failed: $relPath : " + (Get-InnerMessage $_))
        }
    }

    try {
        if ($kinds -contains 'modules') {
            $vbTypes = Get-VbComponentTypes $app
            foreach ($m in $app.CurrentProject.AllModules) {
                $ext = if ($vbTypes[$m.Name] -eq 2) { '.cls' } else { '.bas' }
                Save-One 5 $m.Name ('modules/' + (ConvertTo-SafeFileName $m.Name) + $ext)
            }
            # VBA の参照設定（[ツール]→[参照設定]）。見るだけのスナップショット
            try {
                $refs = New-Object System.Collections.ArrayList
                foreach ($r in $app.References) {
                    $o = [ordered]@{ name = [string]$r.Name; builtIn = [bool]$r.BuiltIn; isBroken = [bool]$r.IsBroken }
                    try { $o.guid = [string]$r.Guid; $o.major = [int]$r.Major; $o.minor = [int]$r.Minor } catch {}
                    try { $o.fullPath = [string]$r.FullPath } catch { $o.fullPath = $null }
                    [void]$refs.Add($o)
                }
                $files[$script:ReferencesFile] = @{
                    text     = (ConvertTo-PrettyJson ([ordered]@{ '$note' = 'Read-only snapshot of VBA references (Tools > References) generated by Access VSC.'; references = $refs.ToArray() }))
                    encoding = 'schema'
                }
            }
            catch { $warnings.Add('Could not read VBA references: ' + (Get-InnerMessage $_)) }
        }
        if ($kinds -contains 'forms') {
            foreach ($o in $app.CurrentProject.AllForms) { Save-One 2 $o.Name ('forms/' + (ConvertTo-SafeFileName $o.Name) + '.txt') }
        }
        if ($kinds -contains 'reports') {
            foreach ($o in $app.CurrentProject.AllReports) { Save-One 3 $o.Name ('reports/' + (ConvertTo-SafeFileName $o.Name) + '.txt') }
        }
        if ($kinds -contains 'macros') {
            foreach ($o in $app.CurrentProject.AllMacros) { Save-One 4 $o.Name ('macros/' + (ConvertTo-SafeFileName $o.Name) + '.txt') }
        }
        $db = $app.CurrentDb()
        if ($kinds -contains 'queries') {
            foreach ($qd in $db.QueryDefs) {
                if ($qd.Name.StartsWith('~')) { continue }
                $rel = 'queries/' + (ConvertTo-SafeFileName $qd.Name) + '.sql'
                try {
                    $sql = [string]$qd.SQL
                    $header = ''
                    if ([int]$qd.Type -eq 112 -or $qd.Connect) {
                        $meta = [ordered]@{ type = [int]$qd.Type; connect = (Hide-Password $qd.Connect); returnsRecords = [bool]$qd.ReturnsRecords }
                        $header = '-- accessvsc: ' + (ConvertTo-Json -InputObject $meta -Compress) + "`r`n"
                    }
                    $files[$rel] = @{ text = $header + $sql; encoding = 'dao' }
                }
                catch { $failed.Add($rel); $warnings.Add("Export failed: $rel : " + (Get-InnerMessage $_)) }
            }
        }
        if ($kinds -contains 'tables') {
            foreach ($td in $db.TableDefs) {
                if ($td.Name -match '^(MSys|~)' ) { continue }
                $rel = 'tables/' + (ConvertTo-SafeFileName $td.Name) + '.json'
                try {
                    $files[$rel] = @{ text = (ConvertTo-PrettyJson (Export-TableSchema $db $td)); encoding = 'schema' }
                }
                catch { $failed.Add($rel); $warnings.Add("Export failed: $rel : " + (Get-InnerMessage $_)) }
            }
            $rels = New-Object System.Collections.ArrayList
            foreach ($r in $db.Relations) {
                if ($r.Name -match '^MSys') { continue }
                $fl = @(); foreach ($f in $r.Fields) { $fl += [ordered]@{ field = $f.Name; foreignField = $f.ForeignName } }
                [void]$rels.Add([ordered]@{
                        name = $r.Name; table = $r.Table; foreignTable = $r.ForeignTable
                        enforceIntegrity = (($r.Attributes -band 2) -eq 0)
                        cascadeUpdate = (($r.Attributes -band 256) -ne 0); cascadeDelete = (($r.Attributes -band 4096) -ne 0)
                        fields = $fl
                    })
            }
            $files[$script:RelationshipsFile] = @{
                text = (ConvertTo-PrettyJson ([ordered]@{ '$note' = 'Read-only snapshot generated by Access VSC.'; relationships = $rels.ToArray() }))
                encoding = 'schema'
            }
        }
    }
    finally {
        Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
    }

    # --- ファイルへ書き出し（内容が同じなら書かない） ---
    $manifest = Read-Manifest $src
    $written = New-Object System.Collections.Generic.List[string]
    $unchanged = 0
    foreach ($rel in $files.Keys) {
        $full = Join-Path $src ($rel.Replace('/', '\'))
        $text = $files[$rel].text
        $same = $false
        if (Test-Path -LiteralPath $full) { $same = ((Read-SourceText $full) -ceq $text) }
        if (-not $same) { Write-SourceText $full $text; $written.Add($rel) } else { $unchanged++ }
        $manifest.files[$rel] = @{ hash = (Get-FileSha256 $full); encoding = $files[$rel].encoding }
    }
    # --- DB から消えたオブジェクトのファイルを削除 ---
    $deleted = New-Object System.Collections.Generic.List[string]
    foreach ($rel in @(Get-SourceFiles $src)) {
        $info = Get-ObjectInfoFromRelPath $rel
        if (-not $info) { continue }
        $kindOfRel = switch ($info.kind) { 'relationships' { 'tables' } 'references' { 'modules' } default { $info.kind } }
        if ($kinds -notcontains $kindOfRel) { continue }
        if ($files.Contains($rel) -or $failed.Contains($rel)) { continue }
        Remove-Item -LiteralPath (Join-Path $src ($rel.Replace('/', '\'))) -Force
        [void]$manifest.files.Remove($rel)
        $deleted.Add($rel)
    }
    foreach ($rel in @($manifest.files.Keys)) {
        if (-not (Test-Path -LiteralPath (Join-Path $src ($rel.Replace('/', '\'))))) { [void]$manifest.files.Remove($rel) }
    }
    $manifest.database = $P.db
    $manifest.exportedAt = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')
    Save-Manifest $src $manifest

    return [ordered]@{
        ok       = ($failed.Count -eq 0)
        result   = [ordered]@{
            objects   = $files.Count
            written   = @($written)
            unchanged = $unchanged
            deleted   = @($deleted)
            failed    = @($failed)
        }
        warnings = @($warnings)
        error    = $(if ($failed.Count) { "$($failed.Count) object(s) could not be exported" } else { $null })
    }
}

# ============================================================
# インポート（ソースフォルダ -> Access）
# ============================================================
function Convert-ModuleSource([string]$Text, [bool]$IsClass) {
    # VBE 形式の .bas/.cls も受け付けるよう、LoadFromText 用に整える
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.AddRange([string[]]($Text -split "`r?`n"))
    $i = 0
    if ($lines.Count -gt 0 -and $lines[0] -match '^VERSION\s') {
        $lines.RemoveAt(0)
        if ($lines.Count -gt 0 -and $lines[0].Trim() -eq 'BEGIN') {
            while ($lines.Count -gt 0 -and $lines[0].Trim() -ne 'END') { $lines.RemoveAt(0) }
            if ($lines.Count -gt 0) { $lines.RemoveAt(0) }
        }
    }
    $attrs = New-Object System.Collections.Generic.List[string]
    while ($lines.Count -gt 0 -and $lines[0] -match '^Attribute VB_') {
        if ($lines[0] -notmatch '^Attribute VB_Name\s*=') { $attrs.Add($lines[0]) }
        $lines.RemoveAt(0)
    }
    $head = New-Object System.Collections.Generic.List[string]
    if ($IsClass) {
        foreach ($h in $script:ClassHeader) {
            $key = ($h -split '=')[0].Trim()
            $given = $attrs | Where-Object { $_ -like "$key*" } | Select-Object -First 1
            $head.Add($(if ($given) { $given } else { $h }))
        }
        foreach ($a in $attrs) { if (-not ($head -contains $a) -and $a -notmatch 'VB_(GlobalNameSpace|Creatable|PredeclaredId|Exposed)') { $head.Add($a) } }
    }
    $all = @($head) + @($lines)
    return ($all -join "`r`n")
}

function Write-ForLoad([string]$Path, [string]$Text, [string]$Encoding, $Warnings, [string]$RelPath) {
    $Text = $Text -replace "`r?`n", "`r`n"
    $enc = Get-TextEncoding $Encoding
    if ($Encoding -eq 'ansi') {
        $round = $enc.GetString($enc.GetBytes($Text))
        if ($round -cne $Text) {
            $a = $Text -split "`r`n"; $b = $round -split "`r`n"
            for ($k = 0; $k -lt [Math]::Min($a.Count, $b.Count); $k++) {
                if ($a[$k] -cne $b[$k]) { $Warnings.Add("$RelPath line $($k + 1): contains characters that cannot be stored in VBA (ANSI/Shift-JIS); they become '?'"); break }
            }
        }
    }
    [System.IO.File]::WriteAllText($Path, $Text, $enc)
}

function Import-Query($db, [string]$Name, [string]$Text) {
    $meta = $null
    $lines = @($Text -split "`r?`n")
    $body = New-Object System.Collections.Generic.List[string]
    foreach ($l in $lines) {
        if ($l -match '^--\s*accessvsc:\s*(\{.*\})\s*$') { $meta = $Matches[1] | ConvertFrom-Json; continue }
        $body.Add($l)
    }
    $sql = ($body -join "`r`n").Trim()
    if (-not $sql) { throw 'SQL is empty' }
    $qd = $null
    foreach ($q in $db.QueryDefs) { if ($q.Name -eq $Name) { $qd = $q; break } }
    $created = $false
    if (-not $qd) {
        $qd = $db.CreateQueryDef($Name)   # 名前だけで作り、SQL は後で設定（パススルー対応のため）
        $created = $true
    }
    try {
        if ($meta -and $meta.connect -and $meta.connect -notmatch '\*\*\*') { $qd.Connect = [string]$meta.connect }
        $qd.SQL = $sql
        if ($meta -and $null -ne $meta.returnsRecords -and $qd.Connect) { $qd.ReturnsRecords = [bool]$meta.returnsRecords }
    }
    catch {
        if ($created) { try { $db.QueryDefs.Delete($Name) } catch {} }
        throw
    }
    # CurrentDb.CreateQueryDef(名前) は自動で追加されるので Append は不要
}

# LoadFromText は、直前の操作の直後などにまれに失敗する（「データベースが開いていない」など）。少し待って 3 回まで試す
function Invoke-LoadFromText($app, [int]$AcType, [string]$Name, [string]$Path) {
    for ($try = 1; ; $try++) {
        try { $app.LoadFromText($AcType, $Name, $Path); return }
        catch {
            if ($try -ge 3) { throw }
            Start-Sleep -Milliseconds (500 * $try)
        }
    }
}

# ソースの Public な手続き名（Sub / Function / Property）
function Get-PublicProcNames([string]$Text) {
    $names = New-Object System.Collections.Generic.List[string]
    foreach ($m in [regex]::Matches($Text, '(?im)^[ \t]*(?:Public[ \t]+)?(?:Static[ \t]+)?(?:Sub|Function|Property[ \t]+(?:Get|Let|Set))[ \t]+([A-Za-z_\u0080-\uFFFF][\w\u0080-\uFFFF]*)[ \t]*\(')) {
        $line = $m.Value.TrimStart()
        if ($line -match '^(Private|Friend)\b') { continue }
        if (-not $names.Contains($m.Groups[1].Value)) { $names.Add($m.Groups[1].Value) }
    }
    return , $names.ToArray()
}

# 取り込んだものが本当に DB に入ったかを、Access 側から確かめる（ファイルの大きさなどでは判断しない）
#   ・オブジェクトがあるか ・標準/クラスモジュールの Public な手続きが消えていないか
function Test-ImportedObjects($app, $Src, [string[]]$Files) {
    $missing = New-Object System.Collections.ArrayList
    $vbTypes = Get-VbComponentTypes $app
    $db = $app.CurrentDb()
    $checked = 0
    foreach ($rel in $Files) {
        $info = Get-ObjectInfoFromRelPath $rel
        if (-not $info -or $info.readOnly) { continue }
        $checked++
        try {
            switch ($info.kind) {
                'modules' {
                    if (-not $vbTypes.ContainsKey($info.name)) { [void]$missing.Add([ordered]@{ file = $rel; message = "モジュール $($info.name) が DB にありません" }); continue }
                    $want = Get-PublicProcNames (Read-SourceText (Join-Path $Src ($rel.Replace('/', '\'))))
                    if (-not $want.Count) { continue }
                    $cm = $app.VBE.ActiveVBProject.VBComponents.Item($info.name).CodeModule
                    $code = if ($cm.CountOfLines -gt 0) { $cm.Lines(1, $cm.CountOfLines) } else { '' }
                    $have = Get-PublicProcNames $code
                    $lost = @($want | Where-Object { $have -notcontains $_ })
                    if ($lost.Count) { [void]$missing.Add([ordered]@{ file = $rel; message = "Public な手続きが DB にありません: $($lost -join ', ')" }) }
                }
                'queries' { [void]$db.QueryDefs.Item($info.name).Name }
                'forms' { [void]$app.CurrentProject.AllForms.Item($info.name).Name }
                'reports' { [void]$app.CurrentProject.AllReports.Item($info.name).Name }
                'macros' { [void]$app.CurrentProject.AllMacros.Item($info.name).Name }
            }
        }
        catch { [void]$missing.Add([ordered]@{ file = $rel; message = "$($info.name) が DB にありません" }) }
    }
    return [ordered]@{ checked = $checked; ok = ($missing.Count -eq 0); missing = $missing.ToArray() }
}

function Invoke-ImportOp($app, $P) {
    $src = $P.src
    $manifest = Read-Manifest $src
    $tmp = New-TempDir
    $imported = New-Object System.Collections.Generic.List[string]
    $deleted = New-Object System.Collections.Generic.List[string]
    $errors = New-Object System.Collections.ArrayList
    $warnings = New-Object System.Collections.Generic.List[string]
    $needCompile = $false
    $db = $app.CurrentDb()
    $vbTypes = Get-VbComponentTypes $app

    try {
        foreach ($rel in @($P.files)) {
            $info = Get-ObjectInfoFromRelPath $rel
            $full = Join-Path $src ($rel.Replace('/', '\'))
            try {
                if (-not $info) { throw 'Not a managed source file (expected modules/*.bas|cls, forms/*.txt, reports/*.txt, macros/*.txt, queries/*.sql)' }
                if ($info.readOnly) { $warnings.Add("$rel is a read-only snapshot and was skipped"); continue }
                $text = Read-SourceText $full
                switch ($info.kind) {
                    'queries' { Import-Query $db $info.name $text }
                    'modules' {
                        $isClass = ($info.ext -eq '.cls')
                        $existing = $vbTypes[$info.name]
                        if ($null -ne $existing -and (($existing -eq 2) -ne $isClass)) {
                            $app.DoCmd.DeleteObject(5, $info.name)   # 標準 <-> クラス の変更は作り直す
                        }
                        $t = Join-Path $tmp ([guid]::NewGuid().ToString('N') + '.txt')
                        Write-ForLoad $t (Convert-ModuleSource $text $isClass) 'ansi' $warnings $rel
                        Invoke-LoadFromText $app 5 $info.name $t
                        $needCompile = $true
                    }
                    default {
                        $enc = 'unicode'
                        if ($manifest.files.ContainsKey($rel) -and $manifest.files[$rel].encoding -in @('unicode', 'ansi', 'utf8bom')) {
                            $enc = $manifest.files[$rel].encoding
                        }
                        $t = Join-Path $tmp ([guid]::NewGuid().ToString('N') + '.txt')
                        Write-ForLoad $t $text $enc $warnings $rel
                        Invoke-LoadFromText $app $info.acType $info.name $t
                        if ($info.kind -ne 'macros') { $needCompile = $true }
                    }
                }
                $imported.Add($rel)
            }
            catch {
                [void]$errors.Add([ordered]@{ file = $rel; message = (Get-InnerMessage $_) })
            }
        }
        foreach ($rel in @($P.deleteFiles)) {
            $info = Get-ObjectInfoFromRelPath $rel
            try {
                if (-not $info -or $info.readOnly) { continue }
                if ($info.kind -eq 'queries') { $db.QueryDefs.Delete($info.name) }
                else { $app.DoCmd.DeleteObject($info.acType, $info.name) }
                $deleted.Add($rel)
                $needCompile = $true
            }
            catch { [void]$errors.Add([ordered]@{ file = $rel; message = 'Delete failed: ' + (Get-InnerMessage $_) }) }
        }
    }
    finally { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }

    $compile = $null
    if ($P.compile -and ($needCompile -or $P.forceCompile)) { $compile = Invoke-Compile $app }
    $verify = if ($imported.Count) { Test-ImportedObjects $app $src $imported.ToArray() } else { $null }

    # マニフェスト更新（成功したものだけ）
    foreach ($rel in $imported) {
        $full = Join-Path $src ($rel.Replace('/', '\'))
        $enc = if ($manifest.files.ContainsKey($rel)) { $manifest.files[$rel].encoding } else { $null }
        if (-not $enc) {
            $info = Get-ObjectInfoFromRelPath $rel
            $enc = switch ($info.kind) { 'modules' { 'ansi' } 'queries' { 'dao' } default { 'unicode' } }
        }
        $manifest.files[$rel] = @{ hash = (Get-FileSha256 $full); encoding = $enc }
    }
    foreach ($rel in $deleted) { [void]$manifest.files.Remove($rel) }
    if (-not $manifest.database) { $manifest.database = $P.db }
    $manifest.importedAt = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')
    Save-Manifest $src $manifest

    $ok = ($errors.Count -eq 0) -and (-not $compile -or $compile.compiled) -and (-not $verify -or $verify.ok)
    $msg = $null
    if ($errors.Count) { $msg = "$($errors.Count) file(s) failed to import" }
    elseif ($compile -and -not $compile.compiled) { $msg = 'Imported, but the VBA project does not compile' }
    elseif ($verify -and -not $verify.ok) { $msg = "Imported, but $($verify.missing.Count) object(s)/procedure(s) are missing in the database (see result.verify)" }
    return [ordered]@{
        ok       = $ok
        result   = [ordered]@{ imported = @($imported); deleted = @($deleted); errors = $errors.ToArray(); compile = $compile; verify = $verify }
        warnings = @($warnings)
        error    = $msg
    }
}

# ============================================================
# 実行・評価・SQL・一覧
# ============================================================
# コンパイルできない VBA を実行すると、エラーのダイアログが出て時間切れまで待つことになる。先に確かめて、エラーの場所を返す
function Test-CompiledBeforeRun($app) {
    if ($app.IsCompiled) { return $null }
    try { $app.RunCommand(125) } catch {}   # コンパイル（保存はしない）
    if ($app.IsCompiled) { return $null }
    $c = Invoke-Compile $app $false
    if ($c.compiled) { return $null }
    return [ordered]@{ ok = $false; result = [ordered]@{ compile = $c }; error = 'The VBA project does not compile, so it cannot run. Fix the compile error first (see result.compile).' }
}

function Invoke-RunOp($app, $P) {
    $notCompiled = Test-CompiledBeforeRun $app
    if ($notCompiled) { return $notCompiled }
    $argv = New-Object System.Collections.ArrayList
    [void]$argv.Add([string]$P.function)
    foreach ($a in @($P.args)) { if ($null -ne $a) { [void]$argv.Add($a) } }
    $v = [System.__ComObject].InvokeMember('Run', [System.Reflection.BindingFlags]::InvokeMethod, $null, $app, $argv.ToArray())
    return [ordered]@{ ok = $true; result = [ordered]@{ function = $P.function; value = (ConvertTo-PlainValue $v); type = $(if ($null -ne $v) { $v.GetType().Name } else { 'Empty' }) } }
}

function Invoke-EvalOp($app, $P) {
    $notCompiled = Test-CompiledBeforeRun $app
    if ($notCompiled) { return $notCompiled }
    $v = $app.Eval([string]$P.expression)
    return [ordered]@{ ok = $true; result = [ordered]@{ expression = $P.expression; value = (ConvertTo-PlainValue $v); type = $(if ($null -ne $v) { $v.GetType().Name } else { 'Null' }) } }
}

function Invoke-SqlOp($app, $P) {
    $db = $app.CurrentDb()
    $sql = [string]$P.sql
    if ($P.write) {
        $db.Execute($sql, 128)   # dbFailOnError
        return [ordered]@{ ok = $true; result = [ordered]@{ recordsAffected = [int]$db.RecordsAffected } }
    }
    if ($sql -match '^\s*(INSERT|UPDATE|DELETE|CREATE|ALTER|DROP)\b' -or $sql -match '^\s*SELECT\b(?:(?!\bFROM\b)[\s\S])*\bINTO\b') {
        throw 'This looks like an action query. sql is read-only by default; pass write=true (CLI: -Write) to execute it.'
    }
    $limit = if ($P.limit) { [int]$P.limit } else { 100 }
    $rs = $db.OpenRecordset($sql, 4)   # dbOpenSnapshot（読み取り専用。更新系 SQL はここでエラーになる）
    try {
        $names = @(); $seen = @{}
        for ($i = 0; $i -lt $rs.Fields.Count; $i++) {
            $n = [string]$rs.Fields.Item($i).Name
            $k = $n; $j = 2
            while ($seen.ContainsKey($k)) { $k = "${n}_$j"; $j++ }
            $seen[$k] = 1; $names += $k
        }
        $rows = New-Object System.Collections.ArrayList
        while (-not $rs.EOF -and $rows.Count -lt $limit) {
            $row = [ordered]@{}
            for ($i = 0; $i -lt $names.Count; $i++) { $row[$names[$i]] = ConvertTo-PlainValue $rs.Fields.Item($i).Value }
            [void]$rows.Add($row)
            $rs.MoveNext()
        }
        $truncated = -not $rs.EOF
        return [ordered]@{ ok = $true; result = [ordered]@{ columns = $names; rowCount = $rows.Count; truncated = $truncated; rows = $rows.ToArray() } }
    }
    finally { try { $rs.Close() } catch {} }
}

function Invoke-ListOp($app, $P) {
    $vbTypes = Get-VbComponentTypes $app
    $modules = @(); foreach ($m in $app.CurrentProject.AllModules) { $modules += [ordered]@{ name = $m.Name; type = $(if ($vbTypes[$m.Name] -eq 2) { 'class' } else { 'standard' }) } }
    $forms = @(); foreach ($o in $app.CurrentProject.AllForms) { $forms += $o.Name }
    $reports = @(); foreach ($o in $app.CurrentProject.AllReports) { $reports += $o.Name }
    $macros = @(); foreach ($o in $app.CurrentProject.AllMacros) { $macros += $o.Name }
    $db = $app.CurrentDb()
    $queries = @(); foreach ($q in $db.QueryDefs) { if (-not $q.Name.StartsWith('~')) { $queries += [ordered]@{ name = $q.Name; type = [int]$q.Type } } }
    $tables = @(); foreach ($t in $db.TableDefs) { if ($t.Name -notmatch '^(MSys|~)') { $tables += [ordered]@{ name = $t.Name; linked = [bool]$t.Connect } } }
    return [ordered]@{
        ok     = $true
        result = [ordered]@{
            accessVersion = [string]$app.Version; isCompiled = [bool]$app.IsCompiled
            tables = $tables; queries = $queries; forms = $forms; reports = $reports; macros = $macros; modules = $modules
        }
    }
}

# ============================================================
# レポート / フォーム：レイアウト・プレビュー・ページ設定
#   位置や大きさの単位は twip（1cm = 567 twip）
# ============================================================
$script:ControlTypes = @{
    100 = 'Label'; 101 = 'Rectangle'; 102 = 'Line'; 103 = 'Image'; 104 = 'CommandButton'; 105 = 'OptionButton'
    106 = 'CheckBox'; 107 = 'OptionGroup'; 108 = 'BoundObjectFrame'; 109 = 'TextBox'; 110 = 'ListBox'; 111 = 'ComboBox'
    112 = 'Subform'; 114 = 'ObjectFrame'; 118 = 'PageBreak'; 119 = 'CustomControl'; 122 = 'ToggleButton'; 123 = 'TabControl'
    124 = 'Page'; 126 = 'Attachment'; 127 = 'EmptyCell'; 128 = 'WebBrowser'; 129 = 'NavigationControl'; 130 = 'NavigationButton'; 132 = 'Chart'
}
# 用紙サイズ（Printer.PaperSize の番号 -> 名前と縦置きの幅・高さ twip）
$script:PaperSizes = @{
    1 = @{ name = 'Letter'; w = 12240; h = 15840 }; 5 = @{ name = 'Legal'; w = 12240; h = 20160 }
    8 = @{ name = 'A3'; w = 16838; h = 23811 }; 9 = @{ name = 'A4'; w = 11906; h = 16838 }; 11 = @{ name = 'A5'; w = 8391; h = 11906 }
    12 = @{ name = 'B4'; w = 14570; h = 20636 }; 13 = @{ name = 'B5'; w = 10319; h = 14570 }
}

function ConvertTo-Cm([double]$Twips) { return [Math]::Round($Twips / 567.0, 2) }

function Open-ObjectDesign($app, [string]$Kind, [string]$Name) {
    if ($Kind -eq 'form') {
        $app.DoCmd.OpenForm($Name, 1, '', '', -1, 1)    # acDesign, acFormPropertySettings, acHidden
        return , $app.Forms.Item($Name)                  # , を付けないと PowerShell がコントロールに展開してしまう
    }
    $app.DoCmd.OpenReport($Name, 1, '', '', 1)          # acViewDesign, acHidden
    return , $app.Reports.Item($Name)
}

function Close-Object($app, [string]$Kind, [string]$Name, [bool]$Save) {
    $t = if ($Kind -eq 'form') { 2 } else { 3 }
    $app.DoCmd.Close($t, $Name, $(if ($Save) { 1 } else { 2 }))   # acSaveYes / acSaveNo
}

function Get-ObjectLayout($app, [string]$Kind, [string]$Name) {
    $obj = Open-ObjectDesign $app $Kind $Name
    try {
        $sections = New-Object System.Collections.ArrayList
        $secHeight = @{}
        $misses = 0
        for ($i = 0; $i -lt 30 -and $misses -lt 6; $i++) {
            try {
                $s = $obj.Section($i)
                $secHeight[$i] = [int]$s.Height
                [void]$sections.Add([ordered]@{ index = $i; name = [string]$s.Name; heightCm = (ConvertTo-Cm $s.Height); height = [int]$s.Height; visible = [bool]$s.Visible })
                $misses = 0
            }
            catch { if ($i -gt 4) { $misses++ } }
        }
        $controls = New-Object System.Collections.ArrayList
        foreach ($c in $obj.Controls) {
            $type = [int]$c.ControlType
            $o = [ordered]@{
                name = [string]$c.Name
                type = $(if ($script:ControlTypes.ContainsKey($type)) { $script:ControlTypes[$type] } else { "Type$type" })
                section = [int]$c.Section
                left = [int]$c.Left; top = [int]$c.Top; width = [int]$c.Width; height = [int]$c.Height
                cm = ('L{0} T{1} W{2} H{3}' -f (ConvertTo-Cm $c.Left), (ConvertTo-Cm $c.Top), (ConvertTo-Cm $c.Width), (ConvertTo-Cm $c.Height))
            }
            foreach ($prop in 'Caption', 'ControlSource', 'FontName', 'FontSize', 'TextAlign', 'Visible', 'CanGrow') {
                try { $v = $c.Properties.Item($prop).Value; if ($null -ne $v -and "$v" -ne '') { $o[$prop.Substring(0, 1).ToLower() + $prop.Substring(1)] = $v } } catch {}
            }
            try { $parent = [string]$c.Parent.Name; if ($parent -and $parent -ne $Name) { $o.attachedTo = $parent } } catch {}
            [void]$controls.Add($o)
        }
        $layout = [ordered]@{
            object = $Name; kind = $Kind
            width = [int]$obj.Width; widthCm = (ConvertTo-Cm $obj.Width)
            sections = $sections.ToArray()
        }
        if ($Kind -eq 'report') {
            $p = $obj.Printer
            $ps = [int]$p.PaperSize
            $pr = [ordered]@{
                paperSize = $(if ($script:PaperSizes.ContainsKey($ps)) { $script:PaperSizes[$ps].name } else { "Code$ps" })
                orientation = $(if ([int]$p.Orientation -eq 2) { 'landscape' } else { 'portrait' })
                marginsMm = [ordered]@{
                    top = [Math]::Round($p.TopMargin / 56.6929, 1); bottom = [Math]::Round($p.BottomMargin / 56.6929, 1)
                    left = [Math]::Round($p.LeftMargin / 56.6929, 1); right = [Math]::Round($p.RightMargin / 56.6929, 1)
                }
            }
            if ($script:PaperSizes.ContainsKey($ps)) {
                $paper = $script:PaperSizes[$ps]
                $pw = if ($pr.orientation -eq 'landscape') { $paper.h } else { $paper.w }
                $pr.printableWidth = [int]($pw - $p.LeftMargin - $p.RightMargin)
                $pr.printableWidthCm = ConvertTo-Cm $pr.printableWidth
            }
            $layout.printer = $pr
        }
        $layout.controls = $controls.ToArray()
        $layout.issues = Test-Layout $layout $secHeight
        return $layout
    }
    finally { Close-Object $app $Kind $Name $false }
}

function Test-Layout($Layout, $SecHeight) {
    # レイアウトの問題を見つける（はみ出し・重なり・用紙に入らない）
    $issues = New-Object System.Collections.Generic.List[string]
    $secName = @{}; foreach ($s in $Layout.sections) { $secName[$s.index] = $s.name }
    if ($Layout.printer -and $Layout.printer.printableWidth -and $Layout.width -gt $Layout.printer.printableWidth) {
        $issues.Add(("レポート幅 {0}cm が印刷可能幅 {1}cm（{2} {3}、左右余白込み）を超えています。空白ページが出ます。レポート幅・コントロール・余白・用紙の向きのいずれかを見直してください。" -f $Layout.widthCm, $Layout.printer.printableWidthCm, $Layout.printer.paperSize, $Layout.printer.orientation))
    }
    $ctrls = @($Layout.controls | Where-Object { $_.type -ne 'PageBreak' })
    foreach ($c in $ctrls) {
        if ($c.left + $c.width -gt $Layout.width) {
            $issues.Add(("{0}（{1}）の右端 {2}cm がレポート/フォームの幅 {3}cm を超えています。" -f $c.name, $secName[$c.section], (ConvertTo-Cm ($c.left + $c.width)), $Layout.widthCm))
        }
        if ($SecHeight.ContainsKey($c.section) -and $c.top + $c.height -gt $SecHeight[$c.section]) {
            $issues.Add(("{0} の下端 {1}cm がセクション「{2}」の高さ {3}cm を超えています。" -f $c.name, (ConvertTo-Cm ($c.top + $c.height)), $secName[$c.section], (ConvertTo-Cm $SecHeight[$c.section])))
        }
        if ($c.left -lt 0 -or $c.top -lt 0) { $issues.Add("$($c.name) の位置がマイナスです。") }
    }
    $boxes = @($ctrls | Where-Object { $_.type -notin @('Line', 'Rectangle', 'Image', 'OptionGroup', 'TabControl', 'Page') -and $_.width -gt 0 -and $_.height -gt 0 -and $_.visible -ne $false })
    for ($i = 0; $i -lt $boxes.Count; $i++) {
        for ($j = $i + 1; $j -lt $boxes.Count; $j++) {
            $a = $boxes[$i]; $b = $boxes[$j]
            if ($a.section -ne $b.section) { continue }
            $ox = [Math]::Min($a.left + $a.width, $b.left + $b.width) - [Math]::Max($a.left, $b.left)
            $oy = [Math]::Min($a.top + $a.height, $b.top + $b.height) - [Math]::Max($a.top, $b.top)
            if ($ox -gt 15 -and $oy -gt 15) {
                $issues.Add(("{0} と {1} が重なっています（{2}、重なり 横{3}cm × 縦{4}cm）。" -f $a.name, $b.name, $secName[$a.section], (ConvertTo-Cm $ox), (ConvertTo-Cm $oy)))
            }
        }
    }
    return , $issues.ToArray()
}

function Get-ObjectKind($P) { if ($P.objectType -eq 'form') { return 'form' } return 'report' }

function Invoke-LayoutOp($app, $P) {
    $kind = Get-ObjectKind $P
    return [ordered]@{ ok = $true; result = (Get-ObjectLayout $app $kind ([string]$P.report)) }
}

function Invoke-PreviewOp($app, $P) {
    $kind = Get-ObjectKind $P
    $name = [string]$P.report
    $warnings = New-Object System.Collections.Generic.List[string]
    $layout = $null
    try { $layout = Get-ObjectLayout $app $kind $name } catch { $warnings.Add('Layout check failed: ' + (Get-InnerMessage $_)) }

    if (-not (Test-Path -LiteralPath $P.outDir)) { [void](New-Item -ItemType Directory -Path $P.outDir -Force) }
    $base = Join-Path $P.outDir ((ConvertTo-SafeFileName $name) + $(if ($kind -eq 'form') { '.form' } else { '' }))
    Get-ChildItem -LiteralPath $P.outDir -File | Where-Object { $_.FullName.StartsWith($base + '.pdf') -or $_.FullName.StartsWith($base + '_p') } |
        Remove-Item -Force -ErrorAction SilentlyContinue
    $pdf = $base + '.pdf'
    $pages = $null
    if ($kind -eq 'form') {
        $app.DoCmd.OutputTo(2, $name, 'PDF Format (*.pdf)', $pdf)
    }
    else {
        $app.DoCmd.OpenReport($name, 2, '', [string]$P.where, 1)     # acViewPreview, acHidden（where で絞り込み）
        try {
            try { $pages = [int]$app.Reports.Item($name).Pages } catch {}
            $app.DoCmd.OutputTo(3, $name, 'PDF Format (*.pdf)', $pdf)
        }
        finally { try { Close-Object $app 'report' $name $false } catch {} }
    }
    $res = [ordered]@{ object = $name; kind = $kind; pdf = $pdf }
    if ($pages) { $res.pages = $pages }
    if ($layout) {
        $res.issues = $layout.issues
        if ($layout.printer) { $res.printer = $layout.printer }
        $res.widthCm = $layout.widthCm
    }
    return [ordered]@{ ok = $true; result = $res; warnings = @($warnings) }
}

function Invoke-PageSetupOp($app, $P) {
    $name = [string]$P.report
    $obj = Open-ObjectDesign $app 'report' $name
    $saved = $false
    try {
        # 注意: $p = $obj.Printer で取り出して $obj.Printer = $p と戻す方法は COM 経由では反映されない。
        #       $obj.Printer.<プロパティ> に直接設定する。
        if ($P.paperSize) {
            $code = ($script:PaperSizes.Keys | Where-Object { $script:PaperSizes[$_].name -eq [string]$P.paperSize } | Select-Object -First 1)
            if (-not $code) { throw "Unknown paperSize: $($P.paperSize) (use A3, A4, A5, B4, B5, Letter, Legal)" }
            $obj.Printer.PaperSize = [int]$code
        }
        if ($P.orientation) {
            if ($P.orientation -notin @('portrait', 'landscape')) { throw 'orientation must be portrait or landscape' }
            $obj.Printer.Orientation = $(if ($P.orientation -eq 'landscape') { 2 } else { 1 })
        }
        foreach ($m in 'Top', 'Bottom', 'Left', 'Right') {
            $v = $P["margin$m"]
            if ($null -ne $v -and "$v" -ne '') { $obj.Printer."${m}Margin" = [int]([double]$v * 56.6929) }
        }
        if ($null -ne $P.widthCm -and "$($P.widthCm)" -ne '') { $obj.Width = [int]([double]$P.widthCm * 567) }
        $app.DoCmd.Save(3, $name)    # Printer の変更は Close(acSaveYes) だけでは保存されないため明示的に保存
        Close-Object $app 'report' $name $true
        $saved = $true
    }
    finally { if (-not $saved) { try { Close-Object $app 'report' $name $false } catch {} } }

    # ソースファイルも更新して、同期状態を保つ
    $rel = 'reports/' + (ConvertTo-SafeFileName $name) + '.txt'
    $tmp = New-TempDir
    try {
        $t = Join-Path $tmp 'r.txt'
        $app.SaveAsText(3, $name, $t)
        $r = Read-TextAuto $t
        $full = Join-Path $P.src ($rel.Replace('/', '\'))
        Write-SourceText $full $r.text
        $manifest = Read-Manifest $P.src
        $manifest.files[$rel] = @{ hash = (Get-FileSha256 $full); encoding = $r.encoding }
        Save-Manifest $P.src $manifest
    }
    finally { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }

    $layout = Get-ObjectLayout $app 'report' $name
    return [ordered]@{ ok = $true; result = [ordered]@{ report = $name; sourceFile = $rel; printer = $layout.printer; widthCm = $layout.widthCm; issues = $layout.issues } }
}

# ============================================================
# ユニットテスト
#   DB のコピーに「呼び出し役」モジュール（AVSC_TestHarness）を入れ、
#   各テストを Module.Proc で直接呼んでエラーを受け止める（ダイアログが出ないので速い）。
#   実行時エラー（Assert 以外）は、もう一度素のまま実行して見張り役にエラー行を取らせる。
# ============================================================
$script:AssertErr = -2147212503   # vbObjectError + 9001

function New-TestHarness($Tests) {
    $run = New-Object System.Collections.Generic.List[string]
    $setup = New-Object System.Collections.Generic.List[string]
    $tear = New-Object System.Collections.Generic.List[string]
    for ($i = 0; $i -lt $Tests.Count; $i++) {
        $t = $Tests[$i]; $n = $i + 1
        $run.Add("        Case ${n}: $($t.module).$($t.name)")
        if ($t.setup) { $setup.Add("        Case ${n}: $($t.module).TestSetup") }
        if ($t.teardown) { $tear.Add("        Case ${n}: $($t.module).TestTeardown") }
    }
    $fn = {
        param($name, $cases)
        $body = @(
            "Public Function $name(ByVal i As Long) As String"
            '    On Error GoTo EH'
        )
        if ($cases.Count) { $body += '    Select Case i'; $body += $cases; $body += '    End Select' }
        $body += @(
            "    $name = ""PASS"""
            '    Exit Function'
            'EH:'
            "    $name = ""FAIL"" & Chr`$(31) & CStr(Err.Number) & Chr`$(31) & Err.Source & Chr`$(31) & Err.Description"
            'End Function'
            ''
        )
        return $body
    }
    $lines = @('Option Compare Database', 'Option Explicit', "' Access VSC test harness (generated; exists only in the test copy)", '')
    $lines += & $fn 'AVSC_Run' $run
    $lines += & $fn 'AVSC_Setup' $setup
    $lines += & $fn 'AVSC_Teardown' $tear
    return ($lines -join "`r`n")
}

function Invoke-AppRun($app, [object[]]$ArgList) {
    return [System.__ComObject].InvokeMember('Run', [System.Reflection.BindingFlags]::InvokeMethod, $null, $app, $ArgList)
}

function ConvertFrom-HarnessResult([string]$Text) {
    if ($Text -eq 'PASS') { return @{ pass = $true } }
    $p = $Text -split [char]31
    return @{ pass = $false; number = [long]$p[1]; source = $p[2]; message = $p[3] }
}

function Invoke-TestOp($app, $P, $State) {
    $tests = @($P.tests)
    $tmp = New-TempDir
    try {
        $t = Join-Path $tmp 'h.txt'
        [System.IO.File]::WriteAllText($t, (New-TestHarness $tests), [System.Text.Encoding]::Default)
        $app.LoadFromText(5, 'AVSC_TestHarness', $t)
    }
    finally { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }

    $compile = Invoke-Compile $app
    if (-not $compile.compiled) {
        return [ordered]@{ ok = $false; result = [ordered]@{ compile = $compile; tests = @() }; error = 'The VBA project does not compile, so tests cannot run' }
    }

    $nameCount = @{}
    foreach ($x in $tests) { $nameCount[$x.name] = 1 + [int]$nameCount[$x.name] }
    $results = New-Object System.Collections.ArrayList
    $total = [System.Diagnostics.Stopwatch]::StartNew()
    for ($i = 0; $i -lt $tests.Count; $i++) {
        $tc = $tests[$i]; $n = $i + 1
        $State.phase = "test $($tc.id)"
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $r = [ordered]@{ id = $tc.id; module = $tc.module; name = $tc.name; status = 'passed'; ms = 0; message = $null }
        try {
            $s = ConvertFrom-HarnessResult (Invoke-AppRun $app @('AVSC_Setup', $n))
            if (-not $s.pass) {
                $r.status = 'error'; $r.message = "TestSetup でエラー: $($s.message) ($($s.number))"
            }
            else {
                $res = ConvertFrom-HarnessResult (Invoke-AppRun $app @('AVSC_Run', $n))
                if (-not $res.pass) {
                    if ($res.number -eq $script:AssertErr -or $res.source -eq 'Assert') { $r.status = 'failed'; $r.message = $res.message }
                    else { $r.status = 'error'; $r.message = "実行時エラー $($res.number): $($res.message)"; $r.errorNumber = $res.number }
                }
            }
            $td = ConvertFrom-HarnessResult (Invoke-AppRun $app @('AVSC_Teardown', $n))
            if (-not $td.pass -and $r.status -eq 'passed') { $r.status = 'error'; $r.message = "TestTeardown でエラー: $($td.message) ($($td.number))" }
        }
        catch {
            $r.status = 'error'; $r.message = 'テストを呼び出せませんでした: ' + (Get-InnerMessage $_)
        }
        $r.ms = $sw.ElapsedMilliseconds

        # 実行時エラーの行を特定する：素のまま実行して、見張り役に「デバッグ」→ 行取得 →「リセット」させる
        if ($r.status -eq 'error' -and $r.errorNumber -and $P.locate -ne $false -and $nameCount[$tc.name] -eq 1) {
            $State.runtimeError = $null
            try { [void](Invoke-AppRun $app @('AVSC_Setup', $n)) } catch {}
            try { [void](Invoke-AppRun $app @($tc.name)) } catch {}
            if ($State.runtimeError) { $r.location = $State.runtimeError; $State.runtimeError = $null }
            try { [void](Invoke-AppRun $app @('AVSC_Teardown', $n)) } catch {}
        }
        [void]$results.Add($r)
    }
    $passed = @($results | Where-Object { $_.status -eq 'passed' }).Count
    $failed = @($results | Where-Object { $_.status -eq 'failed' }).Count
    $errors = @($results | Where-Object { $_.status -eq 'error' }).Count
    return [ordered]@{
        ok     = ($failed + $errors -eq 0)
        result = [ordered]@{
            summary = [ordered]@{ total = $results.Count; passed = $passed; failed = $failed; errors = $errors; ms = $total.ElapsedMilliseconds }
            tests   = $results.ToArray()
        }
        error  = $(if ($failed + $errors) { "$failed failed, $errors error(s) out of $($results.Count) tests" } else { $null })
    }
}

# ============================================================
# 帳票スナップショット：指定したレポートをすべて PDF にする（1 回の起動でまとめて）
# ============================================================
function Invoke-SnapshotOp($app, $P, $State) {
    if (-not (Test-Path -LiteralPath $P.outDir)) { [void](New-Item -ItemType Directory -Path $P.outDir -Force) }
    $items = New-Object System.Collections.ArrayList
    foreach ($name in @($P.reports)) {
        $State.phase = "snapshot $name"
        $pdf = Join-Path $P.outDir ((ConvertTo-SafeFileName $name) + '.pdf')
        $r = [ordered]@{ name = $name; pdf = $pdf; error = $null }
        try {
            $app.DoCmd.OpenReport($name, 2, '', '', 1)   # acViewPreview, acHidden
            try { $app.DoCmd.OutputTo(3, $name, 'PDF Format (*.pdf)', $pdf) }
            finally { try { Close-Object $app 'report' $name $false } catch {} }
        }
        catch { $r.error = Get-InnerMessage $_ }
        [void]$items.Add($r)
    }
    return [ordered]@{ ok = $true; result = [ordered]@{ reports = $items.ToArray() } }
}

# ============================================================
# リリース：版番号の書き込み・テーブルの接続先の張り替え・コンパイル（作業用コピーに対して）
# ============================================================
function Set-DbProperty($db, [string]$Name, [string]$Value) {
    try { $db.Properties.Item($Name).Value = $Value }
    catch { $db.Properties.Append($db.CreateProperty($Name, 10, $Value)) }   # dbText
}

function Invoke-StampOp($app, $P) {
    $db = $app.CurrentDb()
    Set-DbProperty $db 'AppVersion' ([string]$P.version)
    Set-DbProperty $db 'AppBuiltAt' ((Get-Date).ToString('yyyy-MM-ddTHH:mm:ss'))
    $relinked = New-Object System.Collections.ArrayList
    $errors = New-Object System.Collections.ArrayList
    if ($P.backendPath) {
        foreach ($td in $db.TableDefs) {
            $c = [string]$td.Connect
            if ($c -notmatch '(?i)(^|;)DATABASE=') { continue }        # Access のリンクテーブルだけ（ODBC は対象外）
            $old = [regex]::Match($c, '(?i)DATABASE=([^;]*)').Groups[1].Value
            if ($P.backendFrom -and ([System.IO.Path]::GetFileName($old) -ne [System.IO.Path]::GetFileName([string]$P.backendFrom))) { continue }
            $new = [regex]::Replace($c, '(?i)DATABASE=[^;]*', ('DATABASE=' + [string]$P.backendPath).Replace('$', '$$'))
            try {
                $td.Connect = $new
                $td.RefreshLink()
                [void]$relinked.Add([ordered]@{ table = $td.Name; from = $old; to = [string]$P.backendPath })
            }
            catch { [void]$errors.Add([ordered]@{ table = $td.Name; message = (Get-InnerMessage $_) }) }
        }
    }
    $compile = Invoke-Compile $app
    $ok = $compile.compiled -and $errors.Count -eq 0
    $msg = $null
    if (-not $compile.compiled) { $msg = 'The VBA project does not compile' }
    elseif ($errors.Count) { $msg = "Could not relink $($errors.Count) table(s) to $($P.backendPath)" }
    return [ordered]@{ ok = $ok; result = [ordered]@{ version = $P.version; relinked = $relinked.ToArray(); relinkErrors = $errors.ToArray(); compile = $compile }; error = $msg }
}

# .accde の作成。SysCmd 603 は「DB を開いていない」かつ「Access が表示されている」状態でないと失敗する
function Invoke-MakeAccdeOp($app, $P, $State) {
    if (Test-Path -LiteralPath $P.accde) { Remove-Item -LiteralPath $P.accde -Force }
    $app.Visible = $true
    try { [AccessVsc.Win32]::MinimizeNoActivate([long]$State.hwnd) } catch {}
    $r = [System.__ComObject].InvokeMember('SysCmd', [System.Reflection.BindingFlags]::InvokeMethod, $null, $app, [object[]]@([int]603, [string]$P.stage, [string]$P.accde))
    if (-not (Test-Path -LiteralPath $P.accde)) {
        return [ordered]@{ ok = $false; result = $null; error = "Access could not create the ACCDE (SysCmd 603 returned $r). Make sure the project compiles and the file is not open." }
    }
    return [ordered]@{ ok = $true; result = [ordered]@{ accde = $P.accde; size = (Get-Item -LiteralPath $P.accde).Length } }
}

# ============================================================
# 入口：Access を起動 -> 操作 -> 後片付け
# ============================================================
function Invoke-AccessOp([string]$Op, $P, $State) {
    $app = $null
    $out = $null
    $keep = [bool]$State.keep
    try {
        if ($keep -and $global:AvscApp) {
            # セッション：前の呼び出しで開いた Access を使う（生きているか確かめてから）
            try { [void]$global:AvscApp.CurrentProject.FullName; $app = $global:AvscApp } catch { $global:AvscApp = $null }
        }
        if ($app) {
            $State.app = $app
            $State.hwnd = $global:AvscHwnd
        }
        else {
            $app = New-Object -ComObject Access.Application
            $State.app = $app
            $State.hwnd = [long]$app.hWndAccessApp()
            $app.Visible = [bool]$P.visible
            $app.AutomationSecurity = 1     # msoAutomationSecurityLow（マクロ有効）
            $State.phase = 'open'
            if (-not $P.noOpen) { $app.OpenCurrentDatabase([string]$P.db, [bool]$P.exclusive) }
            if ($keep) { $global:AvscApp = $app; $global:AvscHwnd = $State.hwnd }
        }
        $State.phase = $Op
        switch ($Op) {
            'export' { $out = Invoke-ExportOp $app $P }
            'import' { $out = Invoke-ImportOp $app $P }
            'compile' { $c = Invoke-Compile $app; $out = [ordered]@{ ok = $c.compiled; result = $c; error = $(if ($c.compiled) { $null } else { 'The VBA project does not compile' }) } }
            'run' { $out = Invoke-RunOp $app $P }
            'eval' { $out = Invoke-EvalOp $app $P }
            'sql' { $out = Invoke-SqlOp $app $P }
            'list' { $out = Invoke-ListOp $app $P }
            'layout' { $out = Invoke-LayoutOp $app $P }
            'preview' { $out = Invoke-PreviewOp $app $P }
            'pagesetup' { $out = Invoke-PageSetupOp $app $P }
            'test' { $out = Invoke-TestOp $app $P $State }
            'snapshot' { $out = Invoke-SnapshotOp $app $P $State }
            'stamp' { $out = Invoke-StampOp $app $P }
            'makeaccde' { $out = Invoke-MakeAccdeOp $app $P $State }
            default { throw "Unknown operation: $Op" }
        }
        $State.result = $out   # 終了処理で固まっても、結果は見張り側で拾えるようにする
    }
    catch {
        $out = [ordered]@{ ok = $false; result = $null; error = (Get-InnerMessage $_) }
    }
    finally {
        $State.phase = 'close'
        if ($app -and $keep) {
            # セッション：開いたままにする。次の操作に影響しないよう、開いているフォーム・レポートは閉じる
            $alive = $true
            try {
                for ($i = 0; $i -lt 50 -and $app.Forms.Count -gt 0; $i++) { $app.DoCmd.Close(2, $app.Forms.Item(0).Name, 2) }
                for ($i = 0; $i -lt 50 -and $app.Reports.Count -gt 0; $i++) { $app.DoCmd.Close(3, $app.Reports.Item(0).Name, 2) }
                [void]$app.CurrentProject.FullName
            }
            catch { $alive = $false }
            if ($alive) {
                $State.phase = 'done'
                $app = $null
            }
            else {
                $State.keep = $false
                $global:AvscApp = $null
            }
        }
        if ($app) {
            # 読み取りだけの操作では、終了時に何も保存しない（Quit の既定は「すべて保存」）。
            # 書き込む操作は、それぞれの処理の中で保存済み（LoadFromText・コンパイルして保存・DoCmd.Save・DAO）
            $readOnlyOp = $Op -in @('export', 'run', 'eval', 'list', 'layout', 'preview', 'snapshot') -or ($Op -eq 'sql' -and -not $P.write)
            if ($readOnlyOp) { try { $app.Quit(2) } catch {} }   # acQuitSaveNone（データベースも一緒に閉じる）
            else {
                try { $app.CloseCurrentDatabase() } catch {}
                try { $app.Quit() } catch {}
            }
            $State.app = $null
            try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) } catch {}
            $app = $null
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        $State.phase = 'done'
    }
    return $out
}
