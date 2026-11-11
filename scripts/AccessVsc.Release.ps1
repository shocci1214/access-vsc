<#
  AccessVsc.Release.ps1
  Runner から読み込まれる「帳票スナップショット」と「配布（ビルド・デプロイ）」の処理。

  帳票スナップショット（snapshot）
    - 全レポート（または指定したもの）を PNG にし、<src>/snapshots/ の基準画像と比べる
    - 違うページは差分画像（違う所が赤）を .accessvsc/snapshots/diff/ に出す
    - -Update で今の見た目を基準画像にする（基準画像は git にコミットする）

  配布（build / deploy / deploy-init）
    - access-deploy.json（DB と同じフォルダ）に設定を書く
    - build : DB を作業用にコピー → 版番号を書き込み → テーブルの接続先を本番に張り替え → .accde を作成
    - deploy: build した .accde を共有フォルダへ公開（releases/<版> に保管、app/ を差し替え、version.json を更新）
              利用者は共有フォルダの Start-<アプリ名>.cmd から起動（最新版を自動でローカルにコピーして開く）
#>

# ------------------------------------------------------------
# PDF -> PNG（まとめて 1 プロセスで）
# ------------------------------------------------------------
function Convert-PdfsToPng([string[]]$Pdfs, [int]$MaxPages = 5) {
    $map = @{}
    if (-not $Pdfs -or -not $Pdfs.Count) { return $map }
    $list = [System.IO.Path]::GetTempFileName()
    try {
        [System.IO.File]::WriteAllLines($list, $Pdfs, (New-Object System.Text.UTF8Encoding($false)))
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = 'powershell.exe'
        $psi.Arguments = '-NoProfile -ExecutionPolicy Bypass -File "{0}" -ListFile "{1}" -MaxPages {2}' -f (Join-Path $script:ScriptDir 'pdf2png.ps1'), $list, $MaxPages
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        $psi.CreateNoWindow = $true
        $psi.StandardOutputEncoding = New-Object System.Text.UTF8Encoding($false)
        $proc = [System.Diagnostics.Process]::Start($psi)
        $task = $proc.StandardOutput.ReadToEndAsync()
        if (-not $proc.WaitForExit(60000 + 15000 * $Pdfs.Count)) { try { $proc.Kill() } catch {}; throw 'PDF to PNG conversion timed out' }
        $line = ($task.Result -split "`r?`n" | Where-Object { $_.Trim().StartsWith('{') } | Select-Object -Last 1)
        if (-not $line) { throw 'PDF to PNG conversion produced no output' }
        $j = $line | ConvertFrom-Json
        if ($j.error) { throw "PDF to PNG conversion failed: $($j.error)" }
        foreach ($r in @($j.results)) { $map[[string]$r.pdf] = $r }
    }
    finally { Remove-Item -LiteralPath $list -Force -ErrorAction SilentlyContinue }
    return $map
}

# ------------------------------------------------------------
# 帳票スナップショット
# ------------------------------------------------------------
function Get-ReportNames([string]$Src) {
    $dir = Join-Path $Src 'reports'
    if (-not (Test-Path -LiteralPath $dir)) { return }
    foreach ($f in Get-ChildItem -LiteralPath $dir -File -Filter '*.txt' | Sort-Object Name) { ConvertFrom-SafeFileName $f.BaseName }
}

function Invoke-SnapshotCommand([string]$Db, [string]$Src, $P, $Res) {
    $names = @(Get-ReportNames $Src)
    if ($P.reports) {
        $pats = @($P.reports | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        $names = @($names | Where-Object { $n = $_; @($pats | Where-Object { $n -like $_ }).Count -gt 0 })
    }
    if (-not $names.Count) {
        $Res.ok = $true
        $Res.message = 'No reports to snapshot (reports/*.txt). Export first.'
        $Res.result = [ordered]@{ summary = [ordered]@{ total = 0 }; reports = @() }
        return
    }
    $st = Get-SyncStatus $Src
    $pending = @(@($st.changed) + @($st.added) | Where-Object { $_ -like 'reports/*' -or $_ -like 'modules/*' -or $_ -like 'queries/*' })
    if ($pending.Count) { $Res.warnings = @("Unimported changes are not reflected in the snapshot: $($pending -join ', ')") }

    $work = Join-Path (Split-Path -Parent $Db) '.accessvsc\snapshots'
    $curDir = Join-Path $work 'current'
    $diffDir = Join-Path $work 'diff'
    $baseDir = Join-Path $Src 'snapshots'
    foreach ($d in $curDir, $diffDir) { if (-not (Test-Path -LiteralPath $d)) { [void](New-Item -ItemType Directory -Path $d -Force) } }
    foreach ($n in $names) {
        $safe = ConvertTo-SafeFileName $n
        Get-ChildItem -LiteralPath $curDir, $diffDir -File -ErrorAction SilentlyContinue | Where-Object { $_.BaseName -eq $safe -or $_.BaseName -like "${safe}_p*" } | Remove-Item -Force
    }

    $P.reports = $names
    $P.outDir = $curDir
    $timeout = if ($P.timeoutSec) { [int]$P.timeoutSec } else { 120 + 20 * $names.Count }
    $out = Invoke-InAccess 'snapshot' $P $timeout
    if (-not $out.ok) { $Res.ok = $false; $Res.error = $out.error; if ($out.dialogs) { $Res.dialogs = $out.dialogs }; return }

    $pdfs = @($out.result.reports | Where-Object { -not $_.error -and (Test-Path -LiteralPath $_.pdf) } | ForEach-Object { [string]$_.pdf })
    $maxPages = if ($P.pages) { [int]$P.pages } else { 5 }
    $png = Convert-PdfsToPng $pdfs $maxPages
    $tolerance = if ($null -ne $P.tolerance) { [int]$P.tolerance } else { 32 }
    $threshold = if ($null -ne $P.threshold) { [double]$P.threshold } else { 0.0 }   # 違うピクセルの割合（%）がこれを超えたら「変化あり」

    $reports = New-Object System.Collections.ArrayList
    foreach ($r in @($out.result.reports)) {
        $safe = ConvertTo-SafeFileName $r.name
        $item = [ordered]@{ name = $r.name; status = 'same'; pageCount = 0; pages = @() }
        if ($r.error) { $item.status = 'error'; $item.message = $r.error; [void]$reports.Add($item); continue }
        $conv = $png[[string]$r.pdf]
        if (-not $conv -or $conv.error) { $item.status = 'error'; $item.message = "PNG conversion failed: $($conv.error)"; [void]$reports.Add($item); continue }
        $item.pageCount = [int]$conv.pageCount
        $current = @($conv.images)
        $baselines = @(Get-ChildItem -LiteralPath $baseDir -File -Filter "${safe}_p*.png" -ErrorAction SilentlyContinue | Where-Object { $_.BaseName -match "^$([regex]::Escape($safe))_p\d+$" } | Sort-Object { [int]($_.BaseName -replace '.*_p', '') } | ForEach-Object { $_.FullName })

        if ($P.update) {
            if (-not (Test-Path -LiteralPath $baseDir)) { [void](New-Item -ItemType Directory -Path $baseDir -Force) }
            foreach ($b in $baselines) { Remove-Item -LiteralPath $b -Force }
            $pages = @()
            for ($i = 0; $i -lt $current.Count; $i++) {
                $dest = Join-Path $baseDir ("${safe}_p$($i + 1).png")
                Copy-Item -LiteralPath $current[$i] -Destination $dest -Force
                $pages += [ordered]@{ page = $i + 1; status = 'updated'; baseline = $dest; current = $current[$i] }
            }
            $item.status = 'updated'; $item.pages = $pages
            [void]$reports.Add($item); continue
        }
        if (-not $baselines.Count) {
            $item.status = 'new'
            $item.message = 'No baseline yet. Run snapshot with update to accept the current look.'
            $item.pages = @(for ($i = 0; $i -lt $current.Count; $i++) { [ordered]@{ page = $i + 1; status = 'new'; current = $current[$i] } })
            [void]$reports.Add($item); continue
        }
        $pages = @()
        $changed = $false
        $n = [Math]::Max($baselines.Count, $current.Count)
        for ($i = 0; $i -lt $n; $i++) {
            $pg = [ordered]@{ page = $i + 1 }
            if ($i -ge $current.Count) { $pg.status = 'removed'; $pg.baseline = $baselines[$i]; $changed = $true }
            elseif ($i -ge $baselines.Count) { $pg.status = 'added'; $pg.current = $current[$i]; $changed = $true }
            else {
                $diff = Join-Path $diffDir ("${safe}_p$($i + 1).png")
                $cmp = [AccessVsc.ImageDiff]::Compare($baselines[$i], $current[$i], $diff, $tolerance)
                $pg.baseline = $baselines[$i]; $pg.current = $current[$i]
                $pg.diffPercent = [Math]::Round($cmp.Ratio * 100, 3)
                if ($cmp.DiffPixels -gt 0 -and ($cmp.Ratio * 100) -gt $threshold) { $pg.status = 'changed'; $pg.diff = $diff; $changed = $true }
                else { $pg.status = 'same'; Remove-Item -LiteralPath $diff -Force -ErrorAction SilentlyContinue }
                if ($cmp.SizeChanged) { $pg.sizeChanged = $true }
            }
            $pages += $pg
        }
        $item.pages = $pages
        if ($changed) { $item.status = 'changed' }
        [void]$reports.Add($item)
    }
    $count = { param($s) @($reports | Where-Object { $_.status -eq $s }).Count }
    $summary = [ordered]@{
        total = $reports.Count; same = (& $count 'same'); changed = (& $count 'changed'); new = (& $count 'new')
        updated = (& $count 'updated'); errors = (& $count 'error')
    }
    $Res.result = [ordered]@{ summary = $summary; baselineDir = $baseDir; reports = $reports.ToArray() }
    $Res.ok = ($summary.changed -eq 0 -and $summary.errors -eq 0)
    if (-not $Res.ok) { $Res.error = "$($summary.changed) report(s) changed, $($summary.errors) error(s)" }
    if ($out.dialogs) { $Res.dialogs = $out.dialogs }
    if ($P.summary) { Add-SnapshotSummaryMarkdown $Res.result $P.summary }
}

function Add-SnapshotSummaryMarkdown($Result, [string]$Path) {
    $s = $Result.summary
    $icon = if ($s.changed + $s.errors) { '❌' } else { '✅' }
    $lines = @("## $icon 帳票の見た目: $($s.same) 件同じ / $($s.changed) 件変化 / $($s.new) 件基準なし", '')
    $bad = @($Result.reports | Where-Object { $_.status -in @('changed', 'error', 'new') })
    if ($bad.Count) {
        $lines += '| レポート | 状態 | 内容 |'
        $lines += '|---|---|---|'
        foreach ($r in $bad) {
            $detail = if ($r.status -eq 'changed') { (@($r.pages | Where-Object { $_.status -ne 'same' } | ForEach-Object { "p$($_.page): $($_.status)$(if ($_.diffPercent) { " $($_.diffPercent)%" })" }) -join ', ') } else { [string]$r.message }
            $lines += "| $($r.name) | $($r.status) | $detail |"
        }
        $lines += ''
        $lines += '差分画像は成果物（artifact）の `.accessvsc/snapshots/diff/` にあります（違う所が赤）。'
    }
    [System.IO.File]::AppendAllText($Path, (($lines -join "`n") + "`n"), (New-Object System.Text.UTF8Encoding($false)))
}

# ------------------------------------------------------------
# 配布設定
# ------------------------------------------------------------
function Get-DeployConfigPath([string]$Db, [string]$Config) {
    if ($Config) { return [System.IO.Path]::GetFullPath($(if ([System.IO.Path]::IsPathRooted($Config)) { $Config } else { Join-Path (Get-Location).Path $Config })) }
    return Join-Path (Split-Path -Parent $Db) 'access-deploy.json'
}

function Read-DeployConfig([string]$Db, [string]$Config) {
    $path = Get-DeployConfigPath $Db $Config
    $c = @{ appName = [System.IO.Path]::GetFileNameWithoutExtension($Db); output = 'dist'; deployTo = $null; backend = $null; keepReleases = 10; path = $path; exists = $false }
    if (Test-Path -LiteralPath $path) {
        $j = ConvertTo-HashtableDeep ((Read-SourceText $path) | ConvertFrom-Json)
        foreach ($k in @($j.Keys)) { $c[$k] = $j[$k] }
        $c.exists = $true
    }
    if ([string]$c.appName -notmatch '^[\w\-. \p{L}]+$' -or [string]$c.appName -match '\.\.') { throw "appName に使えない文字があります（英数字・日本語・- _ . と空白のみ）: $($c.appName)" }
    $base = Split-Path -Parent $path
    $c.outputDir = [System.IO.Path]::GetFullPath($(if ([System.IO.Path]::IsPathRooted([string]$c.output)) { $c.output } else { Join-Path $base $c.output }))
    return $c
}

function Get-SafeVersion([string]$Version) {
    if (-not $Version) { $Version = (Get-Date).ToString('yyyy.MM.dd.HHmm') }
    return ($Version -replace '[\\/:*?"<>|\s]', '_')
}

function Get-GitCommit([string]$Dir) {
    try {
        $c = & git -C $Dir rev-parse --short HEAD 2>$null
        if ($LASTEXITCODE -eq 0) { return [string]$c }
    }
    catch {}
    return $null
}

function Invoke-DeployInit([string]$Db, [string]$Src, $P) {
    $cfgPath = Get-DeployConfigPath $Db $P.config
    $created = New-Object System.Collections.Generic.List[string]
    $skipped = New-Object System.Collections.Generic.List[string]
    if ((Test-Path -LiteralPath $cfgPath) -and -not $P.force) { $skipped.Add($cfgPath) }
    else {
        $cfg = [ordered]@{
            '$note'      = 'Access VSC deploy settings. deployTo: shared folder users start the app from. backend.path: back-end database the release should link to (leave null to keep current links).'
            appName      = $(if ($P.appName) { [string]$P.appName } else { [System.IO.Path]::GetFileNameWithoutExtension($Db) })
            output       = 'dist'
            deployTo     = $(if ($P.deployTo) { [string]$P.deployTo } else { $null })
            backend      = $(if ($P.backendPath) { [ordered]@{ path = [string]$P.backendPath } } else { $null })
            keepReleases = 10
        }
        Write-SourceText $cfgPath (ConvertTo-PrettyJson $cfg)
        $created.Add($cfgPath)
    }
    # リリース用ワークフロー
    $root = if ($P.root) { [System.IO.Path]::GetFullPath($P.root) } else { Find-GitRoot (Split-Path -Parent $Db) }
    if (-not $root) { $root = Split-Path -Parent $Db }
    $wf = Join-Path $root '.github\workflows\access-release.yml'
    if ((Test-Path -LiteralPath $wf) -and -not $P.force) { $skipped.Add($wf) }
    else {
        $c = Read-DeployConfig $Db $P.config
        $u8 = New-Object System.Text.UTF8Encoding($false)
        $t = [System.IO.File]::ReadAllText((Join-Path $script:ScriptDir 'templates\access-release.yml'), $u8)
        $t = $t.Replace('{{TOOLS}}', 'tools/access-vsc').Replace('{{DB}}', (Get-RelForward $root $Db)).Replace('{{OUTPUT}}', (Get-RelForward $root $c.outputDir))
        $dir = Split-Path -Parent $wf
        if (-not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
        [System.IO.File]::WriteAllText($wf, $t, $u8)
        $created.Add($wf)
    }
    $gi = Join-Path $root '.gitignore'
    $giText = if (Test-Path -LiteralPath $gi) { [System.IO.File]::ReadAllText($gi) } else { '' }
    $outRel = (Get-RelForward $root (Read-DeployConfig $Db $P.config).outputDir) + '/'
    if (@($giText -split "`r?`n") -notcontains $outRel) {
        [System.IO.File]::AppendAllText($gi, $(if ($giText -and -not $giText.EndsWith("`n")) { "`n" } else { '' }) + "# Access VSC build output`n$outRel`n", (New-Object System.Text.UTF8Encoding($false)))
    }
    return [ordered]@{
        config  = $cfgPath
        created = @($created)
        skipped = @($skipped)
        notes   = @(
            'access-deploy.json の deployTo（配布先の共有フォルダ）と backend.path（本番のバックエンド DB）を確認してください。'
            'リリース用ワークフロー（access-release.yml）はタグ v* の push で動きます。CI のツール一式（tools/access-vsc）は setup-ci で用意してください。'
        )
    }
}

# ------------------------------------------------------------
# build: .accde を作る
# ------------------------------------------------------------
function Invoke-BuildCommand([string]$Db, [string]$Src, $P, $Res) {
    $cfg = Read-DeployConfig $Db $P.config
    $version = Get-SafeVersion ([string]$P.version)
    $app = [string]$cfg.appName
    $stageDir = Join-Path (Split-Path -Parent $Db) '.accessvsc\build'
    if (Test-Path -LiteralPath $stageDir) { Remove-Item -LiteralPath $stageDir -Recurse -Force }
    [void](New-Item -ItemType Directory -Path $stageDir -Force)
    $stage = Join-Path $stageDir ($app + [System.IO.Path]::GetExtension($Db))
    Copy-Item -LiteralPath $Db -Destination $stage -Force

    $st = Get-SyncStatus $Src
    $pending = @(@($st.changed) + @($st.added) | Where-Object { $_ -notmatch '^(tables/|relationships)' })
    if ($pending.Count) { $Res.warnings = @("Unimported source changes are NOT in this build: $($pending -join ', ')") }

    # 1) 作業用コピーに版番号を書き込み、接続先を張り替え、コンパイル
    $sp = @{ db = $stage; version = $version; timeoutSec = $P.timeoutSec }
    if ($cfg.backend -and $cfg.backend.path) { $sp.backendPath = [string]$cfg.backend.path; $sp.backendFrom = $cfg.backend.from }
    $stamp = Invoke-InAccess 'stamp' $sp $(if ($P.timeoutSec) { [int]$P.timeoutSec } else { 300 })
    if (-not $stamp.ok) {
        $Res.ok = $false; $Res.error = "Build failed while preparing the release copy: $($stamp.error)"; $Res.result = $stamp.result
        if ($stamp.result -and $stamp.result.compile -and $stamp.result.compile.error) { Add-SourceLocation $stamp.result.compile.error $Src }
        return
    }
    # 2) .accde を作成（DB を開いていない Access で）
    $accdeStage = Join-Path $stageDir ($app + '.accde')
    $mk = Invoke-InAccess 'makeaccde' @{ noOpen = $true; stage = $stage; accde = $accdeStage } $(if ($P.timeoutSec) { [int]$P.timeoutSec } else { 300 })
    if (-not $mk.ok) { $Res.ok = $false; $Res.error = "Build failed while creating the ACCDE: $($mk.error)"; if ($mk.dialogs) { $Res.dialogs = $mk.dialogs }; return }

    # 3) 出力フォルダへ
    if (-not (Test-Path -LiteralPath $cfg.outputDir)) { [void](New-Item -ItemType Directory -Path $cfg.outputDir -Force) }
    $outFile = Join-Path $cfg.outputDir ("{0}_{1}.accde" -f $app, $version)
    Copy-Item -LiteralPath $accdeStage -Destination $outFile -Force
    $manifest = [ordered]@{
        appName  = $app
        version  = $version
        file     = [System.IO.Path]::GetFileName($outFile)
        sha256   = (Get-FileSha256 $outFile)
        size     = (Get-Item -LiteralPath $outFile).Length
        builtAt  = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')
        builtBy  = $env:USERNAME
        commit   = (Get-GitCommit (Split-Path -Parent $Db))
        relinked = @($stamp.result.relinked)
    }
    $manifestPath = [System.IO.Path]::ChangeExtension($outFile, '.json')
    Write-SourceText $manifestPath (ConvertTo-PrettyJson $manifest)
    $Res.ok = $true
    $Res.result = [ordered]@{ accde = $outFile; manifest = $manifestPath; version = $version; sha256 = $manifest.sha256; size = $manifest.size; relinked = $manifest.relinked; config = $(if ($cfg.exists) { $cfg.path } else { $null }) }
}

# ------------------------------------------------------------
# deploy: 共有フォルダへ公開（-Yes がないときは計画だけ表示）
# ------------------------------------------------------------
function Invoke-DeployCommand([string]$Db, [string]$Src, $P, $Res) {
    $cfg = Read-DeployConfig $Db $P.config
    if (-not $cfg.deployTo) { throw "deployTo is not set. Run deploy-init or edit $($cfg.path)" }
    $to = [string]$cfg.deployTo
    $app = [string]$cfg.appName
    $plan = [ordered]@{
        deployTo = $to
        appName  = $app
        source   = $(if ($P.from) { [string]$P.from } else { 'build now from ' + $Db })
        writes   = @("$to\releases\<version>\$app.accde", "$to\app\$app.accde", "$to\version.json", "$to\launcher.ps1", "$to\Start-$app.cmd")
    }
    if (-not $P.yes) {
        $Res.ok = $true
        $Res.result = [ordered]@{ dryRun = $true; plan = $plan }
        $Res.message = 'Dry run: nothing was published. Add -Yes (CLI) to publish.'
        return
    }
    if (-not (Test-Path -LiteralPath $to)) { throw "Deploy destination not found or not reachable: $to" }

    # 公開する .accde
    if ($P.from) {
        $accde = [System.IO.Path]::GetFullPath([string]$P.from)
        if (-not (Test-Path -LiteralPath $accde)) { throw "File not found: $accde" }
        $metaPath = [System.IO.Path]::ChangeExtension($accde, '.json')
        $version = if (Test-Path -LiteralPath $metaPath) { [string]((Read-SourceText $metaPath) | ConvertFrom-Json).version } else { Get-SafeVersion ([string]$P.version) }
    }
    else {
        Invoke-BuildCommand $Db $Src $P $Res
        if (-not $Res.ok) { return }
        $accde = $Res.result.accde
        $version = $Res.result.version
    }
    $build = $Res.result
    $sha = Get-FileSha256 $accde

    # releases/<版>/ に保管 → app/ を差し替え → version.json を最後に更新
    $relDir = Join-Path $to "releases\$version"
    if (-not (Test-Path -LiteralPath $relDir)) { [void](New-Item -ItemType Directory -Path $relDir -Force) }
    Copy-Item -LiteralPath $accde -Destination (Join-Path $relDir "$app.accde") -Force
    $appDir = Join-Path $to 'app'
    if (-not (Test-Path -LiteralPath $appDir)) { [void](New-Item -ItemType Directory -Path $appDir -Force) }
    $tmp = Join-Path $appDir "$app.accde.uploading"
    Copy-Item -LiteralPath $accde -Destination $tmp -Force
    Move-Item -LiteralPath $tmp -Destination (Join-Path $appDir "$app.accde") -Force
    $info = [ordered]@{
        appName     = $app
        version     = $version
        file        = "app/$app.accde"
        sha256      = $sha
        publishedAt = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')
        publishedBy = $env:USERNAME
    }
    $u8 = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText((Join-Path $to 'version.json'), (ConvertTo-PrettyJson $info), $u8)
    [System.IO.File]::WriteAllText((Join-Path $relDir 'version.json'), (ConvertTo-PrettyJson $info), $u8)

    # ランナー（利用者の起動用）
    Copy-Item -LiteralPath (Join-Path $script:ScriptDir 'templates\launcher.ps1') -Destination (Join-Path $to 'launcher.ps1') -Force
    $cmd = "@echo off`r`nrem $app を起動します（最新版を自動で取得）`r`npowershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"%~dp0launcher.ps1`"`r`n"
    [System.IO.File]::WriteAllText((Join-Path $to "Start-$app.cmd"), $cmd, [System.Text.Encoding]::Default)

    # 古いリリースを整理
    $keep = if ($cfg.keepReleases) { [int]$cfg.keepReleases } else { 10 }
    $removed = @()
    $rels = @(Get-ChildItem -LiteralPath (Join-Path $to 'releases') -Directory | Sort-Object LastWriteTime -Descending)
    foreach ($d in ($rels | Select-Object -Skip $keep)) { Remove-Item -LiteralPath $d.FullName -Recurse -Force; $removed += $d.Name }

    $Res.ok = $true
    $Res.result = [ordered]@{
        version   = $version
        deployTo  = $to
        published = (Join-Path $appDir "$app.accde")
        release   = $relDir
        launcher  = (Join-Path $to "Start-$app.cmd")
        sha256    = $sha
        build     = $build
        pruned    = $removed
    }
}
