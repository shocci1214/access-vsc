<#
  AccessVsc.Common.ps1
  CLI / MCP サーバー / Access 実行用ランスペースの全部から読み込む共通関数。
  （ファイル名の変換、文字コード、同期マニフェスト、ソース行の対応付け）
#>

$script:ManagedKinds = [ordered]@{
    # フォルダ名 = 種類情報。acType は SaveAsText / LoadFromText の種類番号
    modules = @{ acType = 5; exts = @('.bas', '.cls'); readOnly = $false }
    forms   = @{ acType = 2; exts = @('.txt'); readOnly = $false }
    reports = @{ acType = 3; exts = @('.txt'); readOnly = $false }
    macros  = @{ acType = 4; exts = @('.txt'); readOnly = $false }
    queries = @{ acType = 1; exts = @('.sql'); readOnly = $false }
    tables  = @{ acType = 0; exts = @('.json'); readOnly = $true }
}
$script:ManifestName = '.accessvsc.json'
$script:RelationshipsFile = 'relationships.json'
$script:ReferencesFile = 'references.json'   # VBA の参照設定（見るだけ）

# ------------------------------------------------------------
# オブジェクト名 <-> ファイル名
#   ファイル名に使えない文字は %XX に置き換える（% 自身も）
# ------------------------------------------------------------
function ConvertTo-SafeFileName([string]$Name) {
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $Name.ToCharArray()) {
        if ('\/:*?"<>|%'.IndexOf($ch) -ge 0 -or [int]$ch -lt 32) {
            [void]$sb.AppendFormat('%{0:X2}', [int]$ch)
        }
        else { [void]$sb.Append($ch) }
    }
    $s = $sb.ToString()
    if ($s.Length -gt 0 -and ($s.EndsWith('.') -or $s.EndsWith(' '))) {
        $s = $s.Substring(0, $s.Length - 1) + ('%{0:X2}' -f [int]$s[$s.Length - 1])
    }
    if ($s -match '^(?i)(CON|PRN|AUX|NUL|COM\d|LPT\d)$') {
        $s = ('%{0:X2}' -f [int]$s[0]) + $s.Substring(1)
    }
    return $s
}

function ConvertFrom-SafeFileName([string]$FileName) {
    return [regex]::Replace($FileName, '%([0-9A-Fa-f]{2})', {
            param($m) [string][char][Convert]::ToInt32($m.Groups[1].Value, 16)
        })
}

# ------------------------------------------------------------
# 文字コード
#   Access の SaveAsText は モジュール=ANSI(Shift-JIS) / フォーム等=UTF-16
#   ソースフォルダ上はすべて UTF-8（BOM なし）で持つ
# ------------------------------------------------------------
function Get-TextEncoding([string]$Name) {
    switch ($Name) {
        'unicode' { return New-Object System.Text.UnicodeEncoding($false, $true) }
        'ansi' { return [System.Text.Encoding]::Default }
        'utf8bom' { return New-Object System.Text.UTF8Encoding($true) }
        default { return New-Object System.Text.UTF8Encoding($false) }
    }
}

function Read-TextAuto([string]$Path) {
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
        return @{ text = [System.Text.Encoding]::Unicode.GetString($bytes, 2, $bytes.Length - 2); encoding = 'unicode' }
    }
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        return @{ text = [System.Text.Encoding]::UTF8.GetString($bytes, 3, $bytes.Length - 3); encoding = 'utf8bom' }
    }
    return @{ text = [System.Text.Encoding]::Default.GetString($bytes); encoding = 'ansi' }
}

function Read-SourceText([string]$Path) {
    # ソースフォルダのファイル（UTF-8 前提。BOM 付きも可）
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    $start = 0
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) { $start = 3 }
    return [System.Text.Encoding]::UTF8.GetString($bytes, $start, $bytes.Length - $start)
}

function Write-SourceText([string]$Path, [string]$Text) {
    $dir = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
    [System.IO.File]::WriteAllText($Path, $Text, (New-Object System.Text.UTF8Encoding($false)))
}

function ConvertTo-PrettyJson($InputObject, [int]$Depth = 20, [switch]$Compress) {
    $j = ConvertTo-Json -InputObject $InputObject -Depth $Depth -Compress
    if ($Compress) { return $j }
    if ('AccessVsc.Json' -as [type]) { return [AccessVsc.Json]::Format($j) }
    return ConvertTo-Json -InputObject $InputObject -Depth $Depth
}

function Get-FileSha256([string]$Path) {
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $hash = $sha.ComputeHash([System.IO.File]::ReadAllBytes($Path))
        return ([BitConverter]::ToString($hash) -replace '-', '').ToLowerInvariant()
    }
    finally { $sha.Dispose() }
}

function Get-StringSha1([string]$Text) {
    $sha = [System.Security.Cryptography.SHA1]::Create()
    try {
        $hash = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($Text))
        return ([BitConverter]::ToString($hash) -replace '-', '').Substring(0, 16).ToLowerInvariant()
    }
    finally { $sha.Dispose() }
}

# ------------------------------------------------------------
# ソースファイル <-> Access オブジェクト
# ------------------------------------------------------------
function Get-RelPath([string]$Src, [string]$FullPath) {
    $s = [System.IO.Path]::GetFullPath($Src).TrimEnd('\') + '\'
    $f = [System.IO.Path]::GetFullPath($FullPath)
    if (-not $f.StartsWith($s, [StringComparison]::OrdinalIgnoreCase)) { return $null }
    return $f.Substring($s.Length).Replace('\', '/')
}

function Get-ObjectInfoFromRelPath([string]$RelPath) {
    # 'modules/Foo.bas' -> @{ kind='modules'; name='Foo'; ext='.bas'; acType=5 }
    $parts = $RelPath.Split('/')
    if ($parts.Count -eq 1 -and $parts[0] -eq $script:RelationshipsFile) {
        return @{ kind = 'relationships'; name = ''; ext = '.json'; acType = 0; readOnly = $true; relPath = $RelPath }
    }
    if ($parts.Count -eq 1 -and $parts[0] -eq $script:ReferencesFile) {
        return @{ kind = 'references'; name = ''; ext = '.json'; acType = 0; readOnly = $true; relPath = $RelPath }
    }
    if ($parts.Count -ne 2) { return $null }
    $kind = $parts[0]
    if (-not $script:ManagedKinds.Contains($kind)) { return $null }
    $info = $script:ManagedKinds[$kind]
    $ext = [System.IO.Path]::GetExtension($parts[1]).ToLowerInvariant()
    if ($info.exts -notcontains $ext) { return $null }
    $base = [System.IO.Path]::GetFileNameWithoutExtension($parts[1])
    return @{
        kind = $kind; name = (ConvertFrom-SafeFileName $base); ext = $ext
        acType = $info.acType; readOnly = $info.readOnly; relPath = $RelPath
    }
}

function Get-SourceFiles([string]$Src) {
    # 管理対象のファイルを相対パス（/区切り）で列挙。呼び出し側は必ず @(Get-SourceFiles ...) で受ける
    $list = New-Object System.Collections.Generic.List[string]
    if (-not (Test-Path -LiteralPath $Src)) { return }
    foreach ($kind in $script:ManagedKinds.Keys) {
        $dir = Join-Path $Src $kind
        if (-not (Test-Path -LiteralPath $dir)) { continue }
        foreach ($f in Get-ChildItem -LiteralPath $dir -File) {
            if ($script:ManagedKinds[$kind].exts -contains $f.Extension.ToLowerInvariant()) {
                $list.Add("$kind/$($f.Name)")
            }
        }
    }
    if (Test-Path -LiteralPath (Join-Path $Src $script:RelationshipsFile)) { $list.Add($script:RelationshipsFile) }
    if (Test-Path -LiteralPath (Join-Path $Src $script:ReferencesFile)) { $list.Add($script:ReferencesFile) }
    return $list.ToArray()
}

# ------------------------------------------------------------
# 同期マニフェスト（最後にエクスポート/インポートした時点のファイルのハッシュ）
# ------------------------------------------------------------
function Read-Manifest([string]$Src) {
    $path = Join-Path $Src $script:ManifestName
    $m = @{ version = 1; database = $null; exportedAt = $null; importedAt = $null; files = @{} }
    if (-not (Test-Path -LiteralPath $path)) { return $m }
    $j = (Read-SourceText $path) | ConvertFrom-Json
    foreach ($p in $j.PSObject.Properties) {
        if ($p.Name -ne 'files') { $m[$p.Name] = $p.Value }
    }
    if ($j.files) {
        foreach ($p in $j.files.PSObject.Properties) {
            $m.files[$p.Name] = @{ hash = $p.Value.hash; encoding = $p.Value.encoding }
        }
    }
    return $m
}

function Save-Manifest([string]$Src, $Manifest) {
    $files = [ordered]@{}
    foreach ($k in ($Manifest.files.Keys | Sort-Object)) {
        $files[$k] = [ordered]@{ hash = $Manifest.files[$k].hash; encoding = $Manifest.files[$k].encoding }
    }
    $out = [ordered]@{
        version    = 1
        note       = 'Managed by Access VSC. Do not edit.'
        database   = $Manifest.database
        exportedAt = $Manifest.exportedAt
        importedAt = $Manifest.importedAt
        files      = $files
    }
    Write-SourceText (Join-Path $Src $script:ManifestName) (ConvertTo-PrettyJson $out)
}

function Get-SyncStatus([string]$Src) {
    # マニフェストと比べて 変更 / 追加 / 削除 されたファイルを返す
    $m = Read-Manifest $Src
    $current = @(Get-SourceFiles $Src)
    $changed = New-Object System.Collections.Generic.List[string]
    $added = New-Object System.Collections.Generic.List[string]
    $missing = New-Object System.Collections.Generic.List[string]
    $readOnlyChanged = New-Object System.Collections.Generic.List[string]
    foreach ($rel in $current) {
        $info = Get-ObjectInfoFromRelPath $rel
        $full = Join-Path $Src ($rel.Replace('/', '\'))
        if (-not $m.files.ContainsKey($rel)) {
            if ($info.readOnly) { $readOnlyChanged.Add($rel) } else { $added.Add($rel) }
        }
        elseif ((Get-FileSha256 $full) -ne $m.files[$rel].hash) {
            if ($info.readOnly) { $readOnlyChanged.Add($rel) } else { $changed.Add($rel) }
        }
    }
    foreach ($rel in $m.files.Keys) {
        if ($current -notcontains $rel) {
            $info = Get-ObjectInfoFromRelPath $rel
            if ($info -and -not $info.readOnly) { $missing.Add($rel) }
        }
    }
    return [ordered]@{
        hasManifest     = (Test-Path -LiteralPath (Join-Path $Src $script:ManifestName))
        exportedAt      = $m.exportedAt
        importedAt      = $m.importedAt
        changed         = @($changed | Sort-Object)
        added           = @($added | Sort-Object)
        missing         = @($missing | Sort-Object)
        readOnlyChanged = @($readOnlyChanged | Sort-Object)
        fileCount       = $current.Count
    }
}

# ------------------------------------------------------------
# VBE のコード行番号 -> ソースファイルの行番号
#   VBE は Attribute 行を表示しないので、その分ずらす
# ------------------------------------------------------------
function Find-SourceForModule([string]$Src, [string]$ModuleName) {
    if ($ModuleName -match '^Form_(.+)$') {
        $p = Join-Path $Src ('forms\' + (ConvertTo-SafeFileName $Matches[1]) + '.txt')
        if (Test-Path -LiteralPath $p) { return $p }
    }
    if ($ModuleName -match '^Report_(.+)$') {
        $p = Join-Path $Src ('reports\' + (ConvertTo-SafeFileName $Matches[1]) + '.txt')
        if (Test-Path -LiteralPath $p) { return $p }
    }
    foreach ($ext in '.bas', '.cls') {
        $p = Join-Path $Src ('modules\' + (ConvertTo-SafeFileName $ModuleName) + $ext)
        if (Test-Path -LiteralPath $p) { return $p }
    }
    return $null
}

function Convert-CodeLineToFileLine([string]$FilePath, [int]$CodeLine) {
    # 戻り値は 1 始まりのファイル行番号。対応できなければ $null
    if (-not $FilePath -or $CodeLine -lt 1) { return $null }
    $lines = (Read-SourceText $FilePath) -split "`r?`n"
    $i = 0
    $isForm = $FilePath -match '\\(forms|reports)\\'
    if ($isForm) {
        while ($i -lt $lines.Count -and $lines[$i].Trim() -notmatch '^CodeBehind(Form|Report)$') { $i++ }
        if ($i -ge $lines.Count) { return $null }
        $i++
    }
    else {
        # VERSION 1.0 CLASS / BEGIN ... END ブロック（VBE 形式の .cls）
        if ($i -lt $lines.Count -and $lines[$i] -match '^VERSION\s') {
            $i++
            if ($i -lt $lines.Count -and $lines[$i].Trim() -eq 'BEGIN') {
                while ($i -lt $lines.Count -and $lines[$i].Trim() -ne 'END') { $i++ }
                $i++
            }
        }
    }
    while ($i -lt $lines.Count -and $lines[$i] -match '^Attribute VB_') { $i++ }
    $n = 0
    for (; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match '^\s*Attribute\s+[\w\[\]]+\.VB_') { continue }  # メンバー属性も VBE では非表示
        $n++
        if ($n -eq $CodeLine) { return $i + 1 }
    }
    return $null
}

# ------------------------------------------------------------
# テストの検出
#   標準モジュール（modules/*.bas）の、名前が Test_ で始まる引数なしの Public Sub / Function。
#   同じモジュールに TestSetup / TestTeardown があれば各テストの前後に呼ぶ。
# ------------------------------------------------------------
function Get-TestCases([string]$Src) {
    $cases = New-Object System.Collections.ArrayList
    $dir = Join-Path $Src 'modules'
    if (-not (Test-Path -LiteralPath $dir)) { return }   # 呼び出し側は @(Get-TestCases ...) で受ける
    foreach ($f in Get-ChildItem -LiteralPath $dir -File -Filter '*.bas' | Sort-Object Name) {
        $module = ConvertFrom-SafeFileName $f.BaseName
        if ($module -eq 'AVSC_TestHarness') { continue }
        $lines = @((Read-SourceText $f.FullName) -split "`r?`n")
        $hasSetup = $false; $hasTeardown = $false
        $found = New-Object System.Collections.ArrayList
        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($lines[$i] -match '^\s*(Public\s+)?(Sub|Function)\s+(TestSetup|TestTeardown)\s*\(\s*\)') {
                if ($Matches[3] -eq 'TestSetup') { $hasSetup = $true } else { $hasTeardown = $true }
                continue
            }
            if ($lines[$i] -match '^\s*(Public\s+)?(Sub|Function)\s+(Test_\w+)\s*\(\s*\)') {
                [void]$found.Add([ordered]@{ module = $module; name = $Matches[3]; id = "$module.$($Matches[3])"; file = $f.FullName; line = $i + 1 })
            }
        }
        foreach ($c in $found) { $c.setup = $hasSetup; $c.teardown = $hasTeardown; [void]$cases.Add($c) }
    }
    return $cases.ToArray()
}
