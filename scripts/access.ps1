<#
.SYNOPSIS
  Access VSC コマンドライン。Access データベースを操作し、結果を JSON で返す。
  AI エージェント（Claude Code など）が直接呼び出すことを想定している。

.EXAMPLE
  powershell -NoProfile -ExecutionPolicy Bypass -File access.ps1 status  -DbPath .\app.accdb
  powershell -NoProfile -ExecutionPolicy Bypass -File access.ps1 export  -DbPath .\app.accdb
  powershell -NoProfile -ExecutionPolicy Bypass -File access.ps1 import  -DbPath .\app.accdb
  powershell -NoProfile -ExecutionPolicy Bypass -File access.ps1 import  -DbPath .\app.accdb -Files modules/Module1.bas
  powershell -NoProfile -ExecutionPolicy Bypass -File access.ps1 compile -DbPath .\app.accdb
  powershell -NoProfile -ExecutionPolicy Bypass -File access.ps1 run     -DbPath .\app.accdb -Function MyFunc -Arg 1,abc
  powershell -NoProfile -ExecutionPolicy Bypass -File access.ps1 eval    -DbPath .\app.accdb -Expr "DCount('*','T1')"
  powershell -NoProfile -ExecutionPolicy Bypass -File access.ps1 sql     -DbPath .\app.accdb -Sql "SELECT * FROM T1" -Limit 20
  powershell -NoProfile -ExecutionPolicy Bypass -File access.ps1 list    -DbPath .\app.accdb
  powershell -NoProfile -ExecutionPolicy Bypass -File access.ps1 layout  -DbPath .\app.accdb -Report rptSales
  powershell -NoProfile -ExecutionPolicy Bypass -File access.ps1 preview -DbPath .\app.accdb -Report rptSales -Where "Year=2026"
  powershell -NoProfile -ExecutionPolicy Bypass -File access.ps1 pagesetup -DbPath .\app.accdb -Report rptSales -PaperSize A4 -Orientation landscape -Margins 10,10,15,15

  Commands:
    status   Show source files changed since last export/import (does not start Access)
    export   Access -> source folder (refuses if there are unimported changes, unless -Force)
    import   Source folder -> Access. Default: only changed/added files. Then compiles.
    compile  Compile the VBA project and report the first error location
    run      Run a public VBA Function/Sub in a standard module (Application.Run)
    eval     Evaluate an Access expression (Application.Eval)
    sql      Run a SELECT and return rows (read-only). With -Write, run an action query.
    list     List database objects
    layout   Report/form layout: sections, controls (twips and cm), printer settings, detected issues (overflow, overlap)
    preview  Output a report (or form) to PDF and PNG images so you can look at the result
    pagesetup  Change report paper size / orientation / margins (mm) / width (cm). Also rewrites the source file.
    test     Run VBA unit tests (Public Sub Test_* in modules/*.bas) on a copy of the database. Writes JUnit XML.
               -Filter "Test_Tax*"  -JUnit results.xml  -Summary $env:GITHUB_STEP_SUMMARY  -Text
    lint     Static checks of the source folder. Does not need Access or the .accdb (works on any Windows CI).
    serve    Stay running and read one JSON request per line from stdin ({"id":1,"command":"run","params":{"db":"...","function":"F"}}).
             Keeps Access open between requests (closed after ACCESSVSC_SESSION_IDLE seconds idle, default 300). Used by the VS Code extension.
    document Write Markdown documentation of the database structure (tables, ER diagram, queries, forms, reports, VBA,
             who-uses-what) from the source folder. No Access needed. -Out <folder> (default: <db>.docs). Notes are kept.
    format   Fix VBA indentation and trailing spaces in the source folder (no Access). -Files ...  -Check (only list)
    setup-ci Add Assert module, example test, GitHub Actions workflow, pre-commit hook, .gitattributes (no Access)
               -Root <repo root>  -GitInit  -Force
    snapshot Render reports to PNG and compare with baselines in <src>/snapshots (visual regression).
               -Reports "rpt*"  -Update (accept current look)  -Threshold <% of pixels>  -Pages 5  -Summary <md>
    build    Make a release .accde: copy DB, stamp AppVersion, relink tables to backend (access-deploy.json), compile.
               -Version 1.2.0
    deploy   Publish to the shared folder in access-deploy.json (dry run unless -Yes). -From <accde> to republish.
    deploy-init  Create access-deploy.json and .github/workflows/access-release.yml
               -AppName Shop -DeployTo \\server\share\Shop -BackendPath \\server\data\Shop_be.accdb
#>
param(
    [Parameter(Position = 0, Mandatory = $true)]
    [ValidateSet('status', 'export', 'import', 'compile', 'run', 'eval', 'sql', 'list', 'layout', 'preview', 'pagesetup', 'test', 'lint', 'format', 'document', 'serve', 'setup-ci', 'snapshot', 'build', 'deploy', 'deploy-init')]
    [string]$Command,

    [string]$DbPath,
    [string]$Src,

    # import
    [string[]]$Files,
    [switch]$All,
    [switch]$DeleteMissing,
    [switch]$NoCompile,
    # export
    [switch]$Force,
    [string[]]$Kinds,
    # run
    [string]$Function,
    [string[]]$Arg,
    [string]$ArgsJson,
    # eval
    [string]$Expr,
    # sql
    [string]$Sql,
    [int]$Limit = 100,
    [switch]$Write,
    # layout / preview / pagesetup
    [string]$Report,
    [ValidateSet('report', 'form')][string]$ObjectType = 'report',
    [string]$Where,
    [int]$Pages = 3,
    [string]$PaperSize,
    [string]$Orientation,
    [string]$Margins,
    [string]$WidthCm,
    # test / lint / setup-ci
    [string[]]$Filter,
    [string]$JUnit,
    [string]$Summary,
    [switch]$NoLocate,
    [switch]$Text,
    [string]$Root,
    [switch]$GitInit,
    # format
    [switch]$Check,
    # document
    [string]$Out,
    # snapshot
    [string[]]$Reports,
    [switch]$Update,
    [double]$Threshold = -1,
    [int]$Tolerance = -1,
    # build / deploy / deploy-init
    [string]$Version,
    [switch]$Yes,
    [string]$From,
    [string]$Config,
    [string]$AppName,
    [string]$DeployTo,
    [string]$BackendPath,

    # パラメータを JSON ファイルで渡す（引用符の問題を避けるため。VS Code 拡張機能が使う）
    #   例: {"function":"Greet","args":["abc"]} / {"sql":"SELECT ...","limit":50}
    [string]$ParamsFile,

    [int]$TimeoutSec = 0,
    [switch]$Visible,
    [switch]$NoBackup,
    [switch]$Compress
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
. (Join-Path $PSScriptRoot 'AccessVsc.Runner.ps1')

if ($Command -eq 'serve') {
    # 常駐モード：1 行 1 リクエストの JSON を読み、1 行の JSON で返す。Access は開いたままにして使い回す
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    $reader = New-Object System.IO.StreamReader([Console]::OpenStandardInput(), $utf8)
    $writer = New-Object System.IO.StreamWriter([Console]::OpenStandardOutput(), $utf8)
    $writer.AutoFlush = $true
    $idle = 300
    if ($env:ACCESSVSC_SESSION_IDLE -match '^\d+$') { $idle = [int]$env:ACCESSVSC_SESSION_IDLE }
    $useSession = $env:ACCESSVSC_SESSION -ne '0'
    $task = $null
    try {
        while ($true) {
            if (-not $task) { $task = $reader.ReadLineAsync() }
            if (-not $task.Wait(500)) { Invoke-SessionMaintenance $idle; continue }
            $line = $task.Result
            $task = $null
            if ($null -eq $line) { break }
            if (-not $line.Trim()) { continue }
            $id = $null
            try {
                $req = $line | ConvertFrom-Json
                $id = $req.id
                $P = ConvertTo-HashtableDeep $req.params
                if (-not $P) { $P = @{} }
                switch ($req.command) {
                    'ping' { $res = [ordered]@{ ok = $true; command = 'ping'; session = [bool]$script:Session } }
                    'release' { Close-AccessSession; $res = [ordered]@{ ok = $true; command = 'release' } }
                    default {
                        $P.session = $useSession
                        if ($P.db -and [System.IO.Path]::IsPathRooted([string]$P.db)) { Set-Location -LiteralPath (Split-Path -Parent ([string]$P.db)) }
                        $res = Invoke-AccessVsc ([string]$req.command) $P
                    }
                }
            }
            catch { $res = [ordered]@{ ok = $false; command = 'serve'; error = $_.Exception.Message } }
            $writer.WriteLine((ConvertTo-PrettyJson ([ordered]@{ id = $id; response = $res }) -Compress))
        }
    }
    finally { Close-AccessSession }
    exit 0
}

function ConvertFrom-LenientJson([string]$Text) {
    # Windows PowerShell 5.1 から別プロセスに渡すと引数内の " が消えるため、'文字列' も受け付ける
    #   -ArgsJson '[1, "abc"]'  /  -ArgsJson "[1, 'abc']"
    try { return ($Text | ConvertFrom-Json) } catch {}
    $fixed = [regex]::Replace($Text, "'((?:[^'\\]|\\.)*)'", {
            param($m) '"' + ($m.Groups[1].Value -replace '\\''', "'" -replace '"', '\"') + '"'
        })
    try { return ($fixed | ConvertFrom-Json) }
    catch { throw "ArgsJson is not valid JSON: $Text  (from PowerShell, write strings with single quotes: -ArgsJson `"[1, 'abc']`")" }
}

function Split-List([string[]]$v) {
    if (-not $v) { return @() }
    return @($v | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

$P = @{
    db = $DbPath; src = $Src; visible = [bool]$Visible; noBackup = [bool]$NoBackup
    timeoutSec = $TimeoutSec
}
switch ($Command) {
    'import' {
        $P.files = @(Split-List $Files); $P.all = [bool]$All; $P.deleteMissing = [bool]$DeleteMissing
        $P.compile = -not $NoCompile
        if ($P.files.Count -eq 0) { $P.files = $null }
    }
    'export' { $P.force = [bool]$Force; if ($Kinds) { $P.kinds = @(Split-List $Kinds) } }
    'run' {
        $P.function = $Function
        if ($ArgsJson) { $P.args = ConvertTo-HashtableDeep (ConvertFrom-LenientJson $ArgsJson) }
        else { $P.args = @(Split-List $Arg) }
    }
    'eval' { $P.expression = $Expr }
    'sql' { $P.sql = $Sql; $P.limit = $Limit; $P.write = [bool]$Write }
    { $_ -in @('layout', 'preview', 'pagesetup') } {
        $P.report = $Report; $P.objectType = $ObjectType; $P.where = $Where; $P.pages = $Pages
        $P.paperSize = $PaperSize; $P.orientation = $Orientation
        if ($WidthCm) { $P.widthCm = $WidthCm }
        if ($Margins) {
            # 上,下,左,右（mm）。1 つだけなら全部同じ
            $m = @(Split-List $Margins)
            if ($m.Count -eq 1) { $m = @($m[0], $m[0], $m[0], $m[0]) }
            if ($m.Count -ne 4) { throw '-Margins は「上,下,左,右」(mm) で指定してください' }
            $P.marginTop = $m[0]; $P.marginBottom = $m[1]; $P.marginLeft = $m[2]; $P.marginRight = $m[3]
        }
    }
    'test' {
        if ($Filter) { $P.filter = @(Split-List $Filter) }
        if ($JUnit) { $P.junit = [System.IO.Path]::GetFullPath($(if ([System.IO.Path]::IsPathRooted($JUnit)) { $JUnit } else { Join-Path (Get-Location).Path $JUnit })) }
        if ($Summary) { $P.summary = $Summary }
        $P.locate = -not $NoLocate
    }
    'lint' { if ($Files) { $P.files = @(Split-List $Files) } }
    'format' { if ($Files) { $P.files = @(Split-List $Files) }; $P.check = [bool]$Check }
    'document' { if ($Out) { $P.out = $Out } }
    'setup-ci' { $P.root = $Root; $P.force = [bool]$Force; $P.gitInit = [bool]$GitInit }
    'snapshot' {
        if ($Reports) { $P.reports = @(Split-List $Reports) }
        $P.update = [bool]$Update
        if ($Threshold -ge 0) { $P.threshold = $Threshold }
        if ($Tolerance -ge 0) { $P.tolerance = $Tolerance }
        if ($PSBoundParameters.ContainsKey('Pages')) { $P.pages = $Pages }
        if ($Summary) { $P.summary = $Summary }
    }
    { $_ -in @('build', 'deploy') } {
        $P.version = $Version; $P.config = $Config
        if ($Command -eq 'deploy') { $P.yes = [bool]$Yes; if ($From) { $P.from = $From } }
    }
    'deploy-init' { $P.config = $Config; $P.appName = $AppName; $P.deployTo = $DeployTo; $P.backendPath = $BackendPath; $P.root = $Root; $P.force = [bool]$Force }
}

function Write-TextResult($res) {
    # 人が読む / CI のログ向けの出力（GitHub Actions 上では注釈も出す）
    $gha = $env:GITHUB_ACTIONS -eq 'true'
    $relFromCwd = {
        param([string]$p)
        $cwd = (Get-Location).Path.TrimEnd('\') + '\'
        if ($p -and $p.StartsWith($cwd, [StringComparison]::OrdinalIgnoreCase)) { return $p.Substring($cwd.Length).Replace('\', '/') }
        return $p
    }
    if ($res.command -eq 'lint' -and $res.result) {
        foreach ($p in @($res.result.problems)) {
            $file = & $relFromCwd (Join-Path $res.source ($p.file.Replace('/', '\')))
            Write-Output ("{0}:{1}: {2} [{3}] {4}" -f $file, $p.line, $p.severity, $p.rule, $p.message)
            if ($gha) { Write-Output ("::{0} file={1},line={2},title={3}::{4}" -f $(switch ($p.severity) { 'error' { 'error' } 'info' { 'notice' } default { 'warning' } }), $file, $p.line, $p.rule, $p.message) }
        }
        Write-Output ("lint: {0} error(s), {1} warning(s), {2} info / {3} files" -f $res.result.errorCount, $res.result.warningCount, [int]$res.result.infoCount, $res.result.files)
        return
    }
    if ($res.command -eq 'document' -and $res.result) {
        $c = $res.result.counts
        Write-Output ("document: {0}  (tables {1}, queries {2}, forms {3}, reports {4}, modules {5}; notes kept {6})" -f $res.result.index, $c.tables, $c.queries, $c.forms, $c.reports, $c.modules, $res.result.notesKept)
        return
    }
    if ($res.command -eq 'format' -and $res.result) {
        foreach ($c in @($res.result.changed)) { Write-Output ("{0}: {1}" -f $(if ($res.result.check) { 'not formatted' } else { 'formatted' }), $c) }
        foreach ($s in @($res.result.skipped)) {
            Write-Output ("{0}: skipped - {1}" -f $s.file, $s.reason)
            if ($gha) { Write-Output ("::warning file={0},title=format::{1}" -f $s.file, $s.reason) }
        }
        Write-Output ("format: {0} changed, {1} skipped / {2} files" -f @($res.result.changed).Count, @($res.result.skipped).Count, $res.result.files)
        return
    }
    if ($res.command -eq 'test' -and $res.result -and $res.result.summary) {
        foreach ($t in @($res.result.tests)) {
            $mark = switch ($t.status) { 'passed' { 'PASS ' } 'failed' { 'FAIL ' } default { 'ERROR' } }
            Write-Output ("{0} {1} ({2} ms)" -f $mark, $t.id, $t.ms)
            if ($t.status -ne 'passed') {
                Write-Output "      $($t.message)"
                $file = $t.file; $line = $t.line
                if ($t.location -and $t.location.code) {
                    Write-Output "      at $($t.location.module) line $($t.location.line): $($t.location.code.Trim())"
                    if ($t.location.file) { $file = $t.location.file; $line = $t.location.fileLine }
                }
                if ($gha -and $file) { Write-Output ("::error file={0},line={1},title={2}::{3}" -f (& $relFromCwd $file), $line, $t.id, $t.message) }
            }
        }
        $s = $res.result.summary
        Write-Output ("tests: {0} passed, {1} failed, {2} error(s) / {3} ({4} ms)" -f $s.passed, $s.failed, $s.errors, $s.total, $s.ms)
        if ($res.result.junit) { Write-Output "JUnit: $($res.result.junit)" }
        foreach ($w in @($res.warnings)) { if ($w) { Write-Output "warning: $w" } }
        if ($res.error -and $res.error -notmatch ' failed, ') { Write-Output "error: $($res.error)" }
        return
    }
    if ($res.command -eq 'snapshot' -and $res.result -and $res.result.reports) {
        foreach ($r in @($res.result.reports)) {
            $mark = switch ($r.status) { 'same' { 'SAME   ' } 'changed' { 'CHANGED' } 'new' { 'NEW    ' } 'updated' { 'UPDATED' } default { 'ERROR  ' } }
            $pages = @($r.pages | Where-Object { $_.status -ne 'same' } | ForEach-Object { "p$($_.page) $($_.status)$(if ($_.diffPercent) { " $($_.diffPercent)%" })" }) -join ', '
            Write-Output ("{0} {1}{2}{3}" -f $mark, $r.name, $(if ($pages) { "  ($pages)" } else { '' }), $(if ($r.message) { "  - $($r.message)" } else { '' }))
            if ($gha -and $r.status -eq 'changed') { Write-Output "::error title=帳票の見た目が変わりました::$($r.name) ($pages)" }
        }
        $s = $res.result.summary
        Write-Output ("snapshot: {0} same, {1} changed, {2} new, {3} updated, {4} error(s)" -f $s.same, $s.changed, $s.new, $s.updated, $s.errors)
        foreach ($w in @($res.warnings)) { if ($w) { Write-Output "warning: $w" } }
        return
    }
    if ($res.command -in @('build', 'deploy') -and $res.ok -and $res.result) {
        $x = $res.result
        if ($x.dryRun) {
            Write-Output $res.message
            Write-Output "deploy to: $($x.plan.deployTo)"
            foreach ($w in $x.plan.writes) { Write-Output "  will write: $w" }
            return
        }
        if ($res.command -eq 'build') { Write-Output "built: $($x.accde)  (version $($x.version), $([math]::Round($x.size / 1KB)) KB)" }
        else { Write-Output "published version $($x.version) to $($x.deployTo)"; Write-Output "users start: $($x.launcher)" }
        foreach ($r in @($x.relinked) + @($x.build.relinked)) { if ($r) { Write-Output "  relinked $($r.table): $($r.from) -> $($r.to)" } }
        foreach ($w in @($res.warnings)) { if ($w) { Write-Output "warning: $w" } }
        return
    }
    if ($res.message) { Write-Output $res.message }
    if ($res.error) { Write-Output "error: $($res.error)" }
    if ($res.result -and $res.result.compile -and $res.result.compile.error) {
        $ce = $res.result.compile.error
        Write-Output ("{0}:{1}: {2}  <- {3}" -f (& $relFromCwd $ce.file), $ce.fileLine, $ce.message, ([string]$ce.code).Trim())
        if ($gha -and $ce.file) { Write-Output ("::error file={0},line={1},title=compile::{2}" -f (& $relFromCwd $ce.file), $ce.fileLine, $ce.message) }
    }
    if (-not $res.message -and -not $res.error) { ConvertTo-PrettyJson $res }
}

if ($ParamsFile) {
    $extra = ConvertTo-HashtableDeep ((Read-SourceText $ParamsFile) | ConvertFrom-Json)
    if ($extra) { foreach ($k in @($extra.Keys)) { $P[$k] = $extra[$k] } }
}

$res = Invoke-AccessVsc $Command $P
try { Initialize-Win32 } catch {}
if ($Text) { Write-TextResult $res }
else { ConvertTo-PrettyJson $res -Compress:$Compress }
if (-not $res.ok) { exit 1 }
exit 0
