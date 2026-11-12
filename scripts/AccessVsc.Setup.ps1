<#
  AccessVsc.Setup.ps1
  テスト・CI/CD の準備（setup-ci コマンド）。Access は使わない。
    - tools/access-vsc/          … CI でも使えるよう、スクリプト一式をリポジトリにコピー
    - <src>/modules/Assert.bas   … テスト用アサーション（なければ）
    - <src>/modules/Test_Example.bas … テストの例（テストが 1 つもなければ）
    - <src>/modules/UI.bas       … フォームを操作するテスト用（フォームがあって、なければ）
    - .github/workflows/access-ci.yml … GitHub Actions（lint はクラウド、test はセルフホストランナー）
    - .githooks/pre-commit       … コミット前に lint（git config core.hooksPath .githooks を設定）
    - .gitattributes / .gitignore
#>

function Get-RelForward([string]$Root, [string]$Path) {
    $r = [System.IO.Path]::GetFullPath($Root).TrimEnd('\') + '\'
    $p = [System.IO.Path]::GetFullPath($Path)
    if ($p.StartsWith($r, [StringComparison]::OrdinalIgnoreCase)) { return $p.Substring($r.Length).Replace('\', '/') }
    return $p.Replace('\', '/')
}

function Find-GitRoot([string]$Start) {
    $d = [System.IO.Path]::GetFullPath($Start)
    while ($d) {
        if (Test-Path -LiteralPath (Join-Path $d '.git')) { return $d }
        $parent = Split-Path -Parent $d
        if (-not $parent -or $parent -eq $d) { break }
        $d = $parent
    }
    return $null
}

function Invoke-SetupCi([string]$Db, [string]$Src, [string]$Root, [bool]$Force, [bool]$GitInit) {
    $tpl = Join-Path $script:ScriptDir 'templates'
    $u8 = New-Object System.Text.UTF8Encoding($false)
    $created = New-Object System.Collections.Generic.List[string]
    $updated = New-Object System.Collections.Generic.List[string]
    $skipped = New-Object System.Collections.Generic.List[string]
    $notes = New-Object System.Collections.Generic.List[string]

    if (-not $Root) { $Root = Find-GitRoot (Split-Path -Parent $Db) }
    if (-not $Root) { $Root = Split-Path -Parent $Db }
    $Root = [System.IO.Path]::GetFullPath($Root)
    $isGit = Test-Path -LiteralPath (Join-Path $Root '.git')
    if (-not $isGit -and $GitInit) {
        $o = & git -C $Root init 2>&1
        if ($LASTEXITCODE -eq 0) { $isGit = $true; $created.Add('.git (git init)') } else { $notes.Add("git init に失敗しました: $o") }
    }

    $put = {
        param([string]$Rel, [string]$Text, [bool]$Overwrite)
        $full = Join-Path $Root ($Rel.Replace('/', '\'))
        $exists = Test-Path -LiteralPath $full
        if ($exists -and -not $Overwrite) { $skipped.Add($Rel); return }
        $dir = Split-Path -Parent $full
        if (-not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
        [System.IO.File]::WriteAllText($full, $Text, $u8)
        if ($exists) { $updated.Add($Rel) } else { $created.Add($Rel) }
    }

    # 1) スクリプト一式（CI 用）。常に最新に更新する
    $toolsRel = 'tools/access-vsc'
    $toolsDir = Join-Path $Root ($toolsRel.Replace('/', '\'))
    foreach ($f in Get-ChildItem -LiteralPath $script:ScriptDir -File) {
        $dest = Join-Path $toolsDir $f.Name
        if (-not (Test-Path -LiteralPath $toolsDir)) { [void](New-Item -ItemType Directory -Path $toolsDir -Force) }
        $same = (Test-Path -LiteralPath $dest) -and ((Get-FileSha256 $dest) -eq (Get-FileSha256 $f.FullName))
        if (-not $same) { Copy-Item -LiteralPath $f.FullName -Destination $dest -Force }
    }
    $tplDest = Join-Path $toolsDir 'templates'
    if (-not (Test-Path -LiteralPath $tplDest)) { [void](New-Item -ItemType Directory -Path $tplDest -Force) }
    Copy-Item -Path (Join-Path $tpl '*') -Destination $tplDest -Force
    $created.Add("$toolsRel/ (スクリプト一式)")

    # 2) テスト用モジュール
    $modDir = Join-Path $Src 'modules'
    if (-not (Test-Path -LiteralPath (Join-Path $modDir 'Assert.bas')) -and -not (Test-Path -LiteralPath (Join-Path $modDir 'Assert.cls'))) {
        & $put (Get-RelForward $Root (Join-Path $modDir 'Assert.bas')) ([System.IO.File]::ReadAllText((Join-Path $tpl 'Assert.bas'), $u8)) $false
    }
    $formDir = Join-Path $Src 'forms'
    $hasForms = (Test-Path -LiteralPath $formDir) -and @(Get-ChildItem -LiteralPath $formDir -Filter '*.txt' -File).Count -gt 0
    if ($hasForms -and -not (Test-Path -LiteralPath (Join-Path $modDir 'UI.bas')) -and -not (Test-Path -LiteralPath (Join-Path $modDir 'UI.cls'))) {
        & $put (Get-RelForward $Root (Join-Path $modDir 'UI.bas')) ([System.IO.File]::ReadAllText((Join-Path $tpl 'UI.bas'), $u8)) $false
    }
    if (@(Get-TestCases $Src).Count -eq 0) {
        & $put (Get-RelForward $Root (Join-Path $modDir 'Test_Example.bas')) ([System.IO.File]::ReadAllText((Join-Path $tpl 'Test_Example.bas'), $u8)) $false
    }

    # 3) GitHub Actions / commit 前フック
    $dbRel = Get-RelForward $Root $Db
    $srcRel = Get-RelForward $Root $Src
    $dbDirRel = Get-RelForward $Root (Split-Path -Parent $Db)
    $dbDirRel = if ($dbDirRel -eq ([System.IO.Path]::GetFullPath($Root).TrimEnd('\')).Replace('\', '/') -or -not $dbDirRel) { '' } else { "$dbDirRel/" }
    $fill = { param([string]$t) $t.Replace('{{TOOLS}}', $toolsRel).Replace('{{DB}}', $dbRel).Replace('{{SRC}}', $srcRel).Replace('{{DBDIR}}', $dbDirRel) }
    & $put '.github/workflows/access-ci.yml' (& $fill ([System.IO.File]::ReadAllText((Join-Path $tpl 'access-ci.yml'), $u8))) $Force
    & $put '.githooks/pre-commit' ((& $fill ([System.IO.File]::ReadAllText((Join-Path $tpl 'pre-commit.sh'), $u8))) -replace "`r`n", "`n") $Force
    if ($isGit) {
        $o = & git -C $Root config core.hooksPath .githooks 2>&1
        if ($LASTEXITCODE -eq 0) { $notes.Add('git config core.hooksPath .githooks を設定しました（コミット前に lint が動きます）') }
        else { $notes.Add("core.hooksPath の設定に失敗しました: $o") }
    }
    else { $notes.Add('git リポジトリではありません。git init 後に「git config core.hooksPath .githooks」を実行すると、コミット前の lint が有効になります。') }

    # 4) .gitattributes / .gitignore（足りない行だけ追記）
    $append = {
        param([string]$Rel, [string[]]$Want, [string]$Header)
        $full = Join-Path $Root $Rel
        $text = if (Test-Path -LiteralPath $full) { [System.IO.File]::ReadAllText($full) } else { '' }
        $have = @($text -split "`r?`n")
        $need = @($Want | Where-Object { $have -notcontains $_ })
        if (-not $need.Count) { return }
        $prefix = if ($text -and -not $text.EndsWith("`n")) { "`n" } else { '' }
        [System.IO.File]::AppendAllText($full, $prefix + "$Header`n" + ($need -join "`n") + "`n", $u8)
        if ($text) { $updated.Add($Rel) } else { $created.Add($Rel) }
    }
    & $append '.gitattributes' @(
        "$srcRel/** text eol=crlf",
        '*.accdb binary',
        '*.mdb binary',
        '*.accde binary'
    ) '# Access VSC: ソースは CRLF のまま、DB はバイナリとして扱う'
    & $append '.gitignore' @('.accessvsc/', '*.laccdb', '*.ldb', 'test-results/') '# Access VSC'

    $notes.Add('追加した Assert.bas / Test_Example.bas（フォームがあれば UI.bas も）は「Access へ反映（インポート）」すると DB に入り、テストを実行できます。')
    $notes.Add('UI.bas を使うと、フォームに入力してボタンを押すテストが書けます（例: UI.OpenForm "frmOrder" → UI.TypeText "txtQty", "3" → UI.Click "cmdCalc"）。')
    $notes.Add('.accdb もコミットしてください（CI はコミットされた .accdb にソースを取り込んでテストします）。大きい場合は Git LFS の利用を検討してください。')
    return [ordered]@{
        root    = $Root
        created = @($created)
        updated = @($updated)
        skipped = @($skipped)
        notes   = @($notes)
    }
}
