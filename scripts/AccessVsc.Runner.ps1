<#
  AccessVsc.Runner.ps1
  CLI と MCP サーバーの共通エンジン。
    - Access の操作を別スレッド（STA ランスペース）で実行する
    - 見張り役：Access のダイアログを検出して内容を記録し、自動でボタンを押す
        * VBA 実行時エラー → [デバッグ] でエラー行を取得してから [リセット]
        * MsgBox など      → 安全側（キャンセル → いいえ → 中止 → OK の順に探す）
    - タイムアウトしたら、自分が起動した Access だけを強制終了する
    - 同じ DB への同時実行を防ぐ（ミューテックス）
    - 書き換える前に DB をバックアップする
#>

$script:ScriptDir = $PSScriptRoot
. (Join-Path $script:ScriptDir 'AccessVsc.Common.ps1')
. (Join-Path $script:ScriptDir 'AccessVsc.Lint.ps1')
. (Join-Path $script:ScriptDir 'AccessVsc.Format.ps1')
. (Join-Path $script:ScriptDir 'AccessVsc.Document.ps1')
. (Join-Path $script:ScriptDir 'AccessVsc.Setup.ps1')
. (Join-Path $script:ScriptDir 'AccessVsc.Release.ps1')

$script:DefaultTimeouts = @{ export = 600; import = 600; compile = 180; run = 60; eval = 60; sql = 60; list = 60; layout = 90; preview = 180; pagesetup = 120; test = 900; snapshot = 0; build = 0; deploy = 0 }
$script:WriteCommands = @('import', 'compile', 'pagesetup')
$script:MaxDialogs = 10

# ------------------------------------------------------------
# セッション：常駐するプロセス（access.ps1 serve・MCP サーバー）では、Access を開いたままにして使い回す。
#   Access の起動・DB を開く・終了（合わせて 3〜4 秒）を省く。
#   ・使うのはこの一覧の操作だけ（テスト・スナップショット・ビルドなどは、これまでどおり別の Access）
#   ・しばらく使わなければ閉じる（Invoke-SessionMaintenance）
#   ・別のプロセスが同じ DB を使いたいときは、名前付きイベントで閉じてもらう（Request-SessionRelease）
# ------------------------------------------------------------
$script:SessionOps = @('export', 'import', 'compile', 'run', 'eval', 'sql', 'list', 'layout', 'preview', 'pagesetup')
$script:Session = $null

function Get-SessionEventName([string]$Db) { return 'Local\AccessVsc_Release_' + (Get-StringSha1 $Db.ToLowerInvariant()) }
function Get-SessionInfoPath([string]$Db) { return Join-Path (Split-Path -Parent $Db) '.accessvsc\session.json' }

# このプロセスのセッションを閉じる（Access を保存せずに終了）
function Close-AccessSession {
    $s = $script:Session
    if (-not $s) { return }
    $script:Session = $null
    try {
        $ps = [powershell]::Create()
        $ps.Runspace = $s.rs
        [void]$ps.AddScript('if ($global:AvscApp) { try { $global:AvscApp.Quit(2) } catch {}; try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($global:AvscApp) } catch {}; $global:AvscApp = $null; [GC]::Collect(); [GC]::WaitForPendingFinalizers() }')
        $h = $ps.BeginInvoke()
        if ($h.AsyncWaitHandle.WaitOne(20000)) {
            try { [void]$ps.EndInvoke($h) } catch {}
            $ps.Dispose()
            $s.rs.Dispose()
        }
        else { try { [void]$ps.BeginStop($null, $null) } catch {} }
    }
    catch {}
    if ($s.accessPid) {
        $proc = Get-Process -Id $s.accessPid -ErrorAction SilentlyContinue
        if ($proc -and -not $proc.WaitForExit(8000)) { Stop-Process -Id $s.accessPid -Force -ErrorAction SilentlyContinue }
    }
    if ($s.event) { try { $s.event.Dispose() } catch {} }
    if ($s.infoFile) { Remove-Item -LiteralPath $s.infoFile -Force -ErrorAction SilentlyContinue }
}

# Access が強制終了されたときなど：閉じる処理をせずに捨てる
function Remove-AccessSession {
    $s = $script:Session
    if (-not $s) { return }
    $script:Session = $null
    if ($s.accessPid) { Stop-Process -Id $s.accessPid -Force -ErrorAction SilentlyContinue }
    try { $s.rs.Dispose() } catch {}
    if ($s.event) { try { $s.event.Dispose() } catch {} }
    if ($s.infoFile) { Remove-Item -LiteralPath $s.infoFile -Force -ErrorAction SilentlyContinue }
}

# 常駐プロセスが定期的に呼ぶ：使っていない・閉じてと頼まれた・Access がいなくなった
function Invoke-SessionMaintenance([int]$IdleSec = 300) {
    $s = $script:Session
    if (-not $s) { return }
    if ($s.accessPid -and -not (Get-Process -Id $s.accessPid -ErrorAction SilentlyContinue)) { Remove-AccessSession; return }
    $asked = $false
    if ($s.event) { try { $asked = $s.event.WaitOne(0) } catch {} }
    if ($asked -or ((Get-Date) - $s.lastUsed).TotalSeconds -gt $IdleSec) { Close-AccessSession }
}

# ほかの常駐プロセスが DB を開いたままにしていたら、閉じてもらって待つ
function Request-SessionRelease([string]$Db) {
    $info = Get-SessionInfoPath $Db
    if (-not (Test-Path -LiteralPath $info)) { return $false }
    try { $j = [System.IO.File]::ReadAllText($info) | ConvertFrom-Json } catch { return $false }
    if ([int]$j.pid -eq $PID -or -not (Get-Process -Id ([int]$j.pid) -ErrorAction SilentlyContinue)) { return $false }
    $ev = $null
    if (-not [System.Threading.EventWaitHandle]::TryOpenExisting((Get-SessionEventName $Db), [ref]$ev)) { return $false }
    try { [void]$ev.Set() } finally { $ev.Dispose() }
    $deadline = (Get-Date).AddSeconds(30)
    while ((Get-Date) -lt $deadline) {
        if (-not (Test-DatabaseLocked $Db)) { return $true }
        Start-Sleep -Milliseconds 300
    }
    return $false
}

# ユーザーが Access の画面でこの DB を開いているか（ウィンドウのタイトルにパスが出る）
function Test-AccessGuiHasDb([string]$Db, [int]$ExceptPid = 0) {
    foreach ($p in @(Get-Process MSACCESS -ErrorAction SilentlyContinue)) {
        if ($p.Id -eq $ExceptPid -or -not $p.MainWindowTitle) { continue }
        if ($p.MainWindowTitle.IndexOf($Db, [StringComparison]::OrdinalIgnoreCase) -ge 0) { return $true }
    }
    return $false
}

# DB を使っているのはだれか：none / session（Access VSC のセッションだけ）/ other（ユーザーの Access など）
function Get-DatabaseLockState([string]$Db) {
    $own = $script:Session -and $script:Session.db -eq $Db
    if ($own) {
        if (Test-AccessGuiHasDb $Db $script:Session.accessPid) { return 'other' }
        return 'session'
    }
    if (-not (Test-DatabaseLocked $Db)) { return 'none' }
    $info = Get-SessionInfoPath $Db
    if (Test-Path -LiteralPath $info) {
        try {
            $j = [System.IO.File]::ReadAllText($info) | ConvertFrom-Json
            if ((Get-Process -Id ([int]$j.pid) -ErrorAction SilentlyContinue) -and -not (Test-AccessGuiHasDb $Db ([int]$j.accessPid))) { return 'session' }
        }
        catch {}
    }
    return 'other'
}

# ------------------------------------------------------------
# Win32 ヘルパーの読み込み（コンパイル結果を %LOCALAPPDATA% にキャッシュ）
# ------------------------------------------------------------
function Initialize-Win32 {
    if ('AccessVsc.Win32' -as [type]) { return }
    $csPath = Join-Path $script:ScriptDir 'AccessVsc.Win32.cs'
    $code = [System.IO.File]::ReadAllText($csPath)
    $cacheDir = Join-Path $env:LOCALAPPDATA 'access-vscode\cache'
    $dll = Join-Path $cacheDir ('AccessVsc.Win32.' + (Get-StringSha1 $code) + '.dll')
    if (Test-Path -LiteralPath $dll) {
        try { Add-Type -LiteralPath $dll; return } catch {}
    }
    try {
        if (-not (Test-Path -LiteralPath $cacheDir)) { [void](New-Item -ItemType Directory -Path $cacheDir -Force) }
        $tmp = "$dll.$PID.tmp"
        Add-Type -TypeDefinition $code -ReferencedAssemblies System.Drawing -OutputAssembly $tmp -OutputType Library
        try { Move-Item -LiteralPath $tmp -Destination $dll -Force -ErrorAction Stop; $tmp = $dll } catch {}
        Add-Type -LiteralPath $tmp
    }
    catch {
        if (-not ('AccessVsc.Win32' -as [type])) { Add-Type -TypeDefinition $code -ReferencedAssemblies System.Drawing }
    }
    Add-Type -AssemblyName System.Drawing
}

# ------------------------------------------------------------
# DB とソースフォルダの決定
# ------------------------------------------------------------
function Resolve-DatabasePath([string]$Db) {
    if (-not $Db) { $Db = $env:ACCESSVSC_DB }
    if (-not $Db) {
        $cands = @(Get-ChildItem -LiteralPath (Get-Location).Path -File | Where-Object { $_.Extension -in '.accdb', '.mdb' })
        if ($cands.Count -eq 1) { $Db = $cands[0].FullName }
        elseif ($cands.Count -gt 1) { throw ('Multiple databases found; specify one with -DbPath: ' + (($cands | ForEach-Object Name) -join ', ')) }
        else { throw 'No database specified. Use -DbPath <path.accdb> or set ACCESSVSC_DB.' }
    }
    if (-not [System.IO.Path]::IsPathRooted($Db)) { $Db = Join-Path (Get-Location).Path $Db }
    $Db = [System.IO.Path]::GetFullPath($Db)
    if (-not (Test-Path -LiteralPath $Db)) { throw "Database not found: $Db" }
    return $Db
}

function Resolve-SourcePath([string]$DbPath, [string]$Src) {
    if (-not $Src) { $Src = $env:ACCESSVSC_SRC }
    if (-not $Src) { return "$DbPath.src" }   # 例: C:\work\app.accdb -> C:\work\app.accdb.src
    if (-not [System.IO.Path]::IsPathRooted($Src)) { $Src = Join-Path (Split-Path -Parent $DbPath) $Src }
    return [System.IO.Path]::GetFullPath($Src)
}

function Test-DatabaseLocked([string]$DbPath) {
    # 他のプロセス（ユーザーが開いている Access など）が使っていれば排他オープンに失敗する
    try {
        $fs = [System.IO.File]::Open($DbPath, 'Open', 'ReadWrite', 'None')
        $fs.Close()
        return $false
    }
    catch { return $true }
}

function New-DatabaseBackup([string]$DbPath, [int]$Keep = 10) {
    $dir = Join-Path (Split-Path -Parent $DbPath) '.accessvsc\backups'
    if (-not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
    $base = [System.IO.Path]::GetFileNameWithoutExtension($DbPath)
    $ext = [System.IO.Path]::GetExtension($DbPath)
    $dest = Join-Path $dir ("{0}_{1}{2}" -f $base, (Get-Date).ToString('yyyyMMdd_HHmmss_fff'), $ext)
    Copy-Item -LiteralPath $DbPath -Destination $dest
    $old = @(Get-ChildItem -LiteralPath $dir -File -Filter "$($base)_*$ext" | Sort-Object Name -Descending | Select-Object -Skip $Keep)
    foreach ($f in $old) { Remove-Item -LiteralPath $f.FullName -Force -ErrorAction SilentlyContinue }
    return $dest
}

# ------------------------------------------------------------
# 見張り役：ダイアログ 1 つを処理する
# ------------------------------------------------------------
function Resolve-Dialog($dlg, $state, [bool]$AllowDebug) {
    $texts = @($dlg.Texts | ForEach-Object { ($_ -replace '\s+', ' ').Trim() } | Where-Object { $_ })
    $buttons = @($dlg.Buttons | ForEach-Object { $_.Text })
    $rec = [ordered]@{ title = $dlg.Title; message = ($texts -join ' '); buttons = $buttons; pressed = $null }

    $debugBtn = $dlg.Buttons | Where-Object { $_.Text -match 'デバッグ|Debug' } | Select-Object -First 1
    $endBtn = $dlg.Buttons | Where-Object { $_.Text -match '終了|^&?End' } | Select-Object -First 1

    if ($debugBtn -and $endBtn) {
        # VBA の実行時エラー
        $rt = [ordered]@{ message = $rec.message; number = $null; module = $null; line = $null; code = $null }
        if ($rec.message -match "'(-?\d+)[^']*'") { $rt.number = [long]$Matches[1] }
        $app = $state.app
        if ($AllowDebug -and $app) {
            [AccessVsc.Win32]::Click($dlg.Hwnd, $debugBtn)
            $rec.pressed = $debugBtn.Text
            Start-Sleep -Milliseconds 700
            try {
                $pane = $app.VBE.ActiveCodePane
                if ($pane) {
                    $sl = 0; $sc = 0; $el = 0; $ec = 0
                    $pane.GetSelection([ref]$sl, [ref]$sc, [ref]$el, [ref]$ec)
                    $rt.module = $pane.CodeModule.Parent.Name
                    $rt.line = $sl
                    if ($sl -ge 1) { $rt.code = $pane.CodeModule.Lines($sl, 1) }
                }
            }
            catch {}
            try { $app.VBE.CommandBars.FindControl(1, 228).Execute() } catch {}   # [実行] - [リセット]
            $rec.pressed += ' -> Reset'
        }
        else {
            [AccessVsc.Win32]::Click($dlg.Hwnd, $endBtn)
            $rec.pressed = $endBtn.Text
        }
        $state.runtimeError = $rt
        return $rec
    }

    # 安全側に答える：確認の質問（はい/いいえ、OK/キャンセル）は「キャンセル」→「いいえ」→「中止」を優先。
    # 「OK」しかない通知は OK。削除などの確認に、自動で「はい」と答えてしまわないため。
    $pick = $null
    foreach ($pat in '^(キャンセル|&?Cancel)', '^(いいえ|&?No)\b', '^(中止|&?Abort)', '^(&?OK|閉じる|&?Close)') {
        $pick = $dlg.Buttons | Where-Object { $_.Text -match $pat } | Select-Object -First 1
        if ($pick) { break }
    }
    if (-not $pick) { $pick = $dlg.Buttons | Select-Object -First 1 }
    if ($pick) { [AccessVsc.Win32]::Click($dlg.Hwnd, $pick); $rec.pressed = $pick.Text }
    else { [AccessVsc.Win32]::Close($dlg.Hwnd); $rec.pressed = '(WM_CLOSE)' }
    return $rec
}

# ------------------------------------------------------------
# Access で 1 操作を実行（別スレッド + 見張り + タイムアウト）
# ------------------------------------------------------------
function Invoke-InAccess([string]$Op, [hashtable]$P, [int]$TimeoutSec, [bool]$UseSession = $false) {
    Initialize-Win32
    if ($UseSession -and $script:Session -and $script:Session.db -ne $P.db) {
        # 別の DB に切り替える：開いたままの Access を先に閉じる（テストなど、DB のコピーを別の Access で開く操作では閉じない）
        Close-AccessSession
    }
    $state = [hashtable]::Synchronized(@{ app = $null; hwnd = $null; phase = 'start'; runtimeError = $null; keep = $UseSession })
    $dialogs = New-Object System.Collections.ArrayList
    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    if ($UseSession -and $script:Session) {
        $rs = $script:Session.rs
        $state.hwnd = $script:Session.hwnd
    }
    else {
        $rs = [runspacefactory]::CreateRunspace()
        $rs.ApartmentState = 'STA'
        $rs.ThreadOptions = 'ReuseThread'
        $rs.Open()
        if ($UseSession) { $script:Session = @{ db = $P.db; rs = $rs; hwnd = $null; accessPid = 0; lastUsed = (Get-Date); event = $null; infoFile = $null } }
    }
    $rs.SessionStateProxy.SetVariable('state', $state)
    $ps = [powershell]::Create()
    $ps.Runspace = $rs
    # 文字列で渡すとグローバル スコープで動くので、読み込んだ関数と開いた Access が次の呼び出しまで残る
    [void]$ps.AddScript(@"
param(`$scriptDir, `$op, `$p)
if (-not (Get-Command Invoke-AccessOp -ErrorAction SilentlyContinue)) {
    . (Join-Path `$scriptDir 'AccessVsc.Common.ps1')
    . (Join-Path `$scriptDir 'AccessVsc.Ops.ps1')
}
Invoke-AccessOp `$op `$p `$state
"@).AddArgument($script:ScriptDir).AddArgument($Op).AddArgument($P)

    $handle = $ps.BeginInvoke()
    $accessPid = 0
    $handled = @{}
    $killed = $null
    $closeSince = $null
    $closeHang = $false
    $state.result = $null
    $allowDebug = ($Op -in @('run', 'eval', 'test'))

    while (-not $handle.IsCompleted) {
        [void]$handle.AsyncWaitHandle.WaitOne(200)   # 終わったらすぐ抜ける
        if (-not $accessPid -and $state.hwnd) { $accessPid = [int][AccessVsc.Win32]::ProcessIdOf([long]$state.hwnd) }
        if ($accessPid) {
            foreach ($d in [AccessVsc.Win32]::FindDialogs([uint32]$accessPid)) {
                $key = [string]$d.Hwnd
                if ($handled.ContainsKey($key)) {
                    # 押しても消えないダイアログは閉じ直す
                    if (($sw.ElapsedMilliseconds - $handled[$key]) -gt 3000) { [AccessVsc.Win32]::Close($d.Hwnd); $handled[$key] = $sw.ElapsedMilliseconds }
                    continue
                }
                $handled[$key] = $sw.ElapsedMilliseconds
                [void]$dialogs.Add((Resolve-Dialog $d $state $allowDebug))
            }
            if ($dialogs.Count -gt $script:MaxDialogs) { $killed = "more than $($script:MaxDialogs) dialogs appeared" }
        }
        # 終了処理（CloseCurrentDatabase / Quit）で固まることがある。結果は出ているので 20 秒で打ち切る
        if ($state.phase -eq 'close') {
            if (-not $closeSince) { $closeSince = $sw.Elapsed.TotalSeconds }
            elseif (-not $killed -and ($sw.Elapsed.TotalSeconds - $closeSince) -gt 20) { $killed = 'Access did not close within 20 sec' ; $closeHang = $true }
        }
        if (-not $killed -and $sw.Elapsed.TotalSeconds -gt $TimeoutSec) { $killed = "did not finish within $TimeoutSec sec (phase: $($state.phase)); a modal form or input wait may be blocking" }
        if ($killed) {
            if ($accessPid) { Stop-Process -Id $accessPid -Force -ErrorAction SilentlyContinue }
            $deadline = $sw.Elapsed.TotalSeconds + 10
            while (-not $handle.IsCompleted -and $sw.Elapsed.TotalSeconds -lt $deadline) { Start-Sleep -Milliseconds 200 }
            break
        }
    }

    $out = $null
    $threadError = $null
    if ($handle.IsCompleted) {
        try {
            $res = $ps.EndInvoke($handle)
            if ($res.Count -gt 0) { $out = $res[$res.Count - 1].psobject.BaseObject }
        }
        catch { $threadError = $_.Exception.Message }
        if (-not $out -and $ps.Streams.Error.Count -gt 0) { $threadError = $ps.Streams.Error[0].ToString() }
        $ps.Dispose()
        if (-not $UseSession) { $rs.Dispose() }
    }
    else {
        # スレッドが戻ってこない：止めるよう依頼だけして先へ進む
        try { [void]$ps.BeginStop($null, $null) } catch {}
    }

    if ($UseSession -and $script:Session) {
        if ($killed -or -not $handle.IsCompleted -or -not $state.keep) {
            # Access を止めた・使えなくなった：セッションを捨てる（次の呼び出しで作り直す）
            Remove-AccessSession
        }
        else {
            $s = $script:Session
            $s.lastUsed = Get-Date
            if ($state.hwnd) { $s.hwnd = $state.hwnd }
            if (-not $s.accessPid -and $accessPid) { $s.accessPid = $accessPid }
            if (-not $s.event) {
                # 別のプロセスから「閉じて」と頼まれるための目印
                try {
                    $s.event = New-Object System.Threading.EventWaitHandle -ArgumentList $false, ([System.Threading.EventResetMode]::AutoReset), (Get-SessionEventName $s.db)
                    $s.infoFile = Get-SessionInfoPath $s.db
                    $dir = Split-Path -Parent $s.infoFile
                    if (-not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
                    [System.IO.File]::WriteAllText($s.infoFile, (ConvertTo-PrettyJson ([ordered]@{ pid = $PID; accessPid = $s.accessPid; db = $s.db; started = (Get-Date).ToString('s') })), (New-Object System.Text.UTF8Encoding($false)))
                }
                catch {}
            }
        }
        $accessPid = 0   # 開いたままにするので、終了は待たない
    }

    # Quit 後も Access が残ることがあるので、少し待ってから終了させる（DB のロックを残さないため）
    if ($accessPid) {
        $proc = Get-Process -Id $accessPid -ErrorAction SilentlyContinue
        if ($proc) {
            if (-not $proc.WaitForExit(8000)) { Stop-Process -Id $accessPid -Force -ErrorAction SilentlyContinue }
        }
    }

    if (-not $out -and $state.result -and ($closeHang -or $state.phase -in @('close', 'done'))) {
        # 処理は終わっていて、終了処理だけ固まった
        $out = $state.result
        $out.warnings = @(@($out.warnings) + 'Access did not close cleanly and was terminated after finishing the operation.' | Where-Object { $_ })
    }
    elseif (-not $out) {
        $out = [ordered]@{ ok = $false; result = $null; error = $(if ($killed) { "Access was terminated: $killed" } elseif ($threadError) { $threadError } else { 'No result from Access' }) }
    }
    elseif ($killed -and -not $closeHang) {
        $out.ok = $false
        $out.error = "Access was terminated: $killed" + $(if ($out.error) { " / $($out.error)" } else { '' })
    }
    $out.dialogs = $dialogs.ToArray()
    if ($state.runtimeError -and $Op -ne 'test') {
        $out.ok = $false
        $out.runtimeError = $state.runtimeError
        $out.error = 'VBA runtime error: ' + $state.runtimeError.message
    }
    $out.accessMs = $sw.ElapsedMilliseconds
    return $out
}

function Convert-PdfToPng([string]$Pdf, [int]$MaxPages = 3) {
    # 別プロセスで変換（WinRT のスレッドが残ってプロセスが終わらなくなるのを避ける）
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = 'powershell.exe'
    $psi.Arguments = '-NoProfile -ExecutionPolicy Bypass -File "{0}" -Pdf "{1}" -MaxPages {2}' -f (Join-Path $script:ScriptDir 'pdf2png.ps1'), $Pdf, $MaxPages
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.CreateNoWindow = $true
    $psi.StandardOutputEncoding = New-Object System.Text.UTF8Encoding($false)
    try {
        $proc = [System.Diagnostics.Process]::Start($psi)
        $task = $proc.StandardOutput.ReadToEndAsync()
        if (-not $proc.WaitForExit(60000)) { try { $proc.Kill() } catch {}; return @{ error = 'timeout' } }
        $line = ($task.Result -split "`r?`n" | Where-Object { $_.Trim().StartsWith('{') } | Select-Object -Last 1)
        if (-not $line) { return @{ error = 'no output' } }
        $j = $line | ConvertFrom-Json
        if ($j.error) { return @{ error = $j.error } }
        return @{ pageCount = [int]$j.pageCount; images = @($j.images) }
    }
    catch { return @{ error = $_.Exception.Message } }
}

function Write-JUnitReport($Result, [string]$Path) {
    # CI（GitHub Actions など）で読める JUnit XML を書く
    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
    $esc = { param($s) [System.Security.SecurityElement]::Escape([string]$s) }
    $s = $Result.summary
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<?xml version="1.0" encoding="UTF-8"?>')
    [void]$sb.AppendLine(('<testsuites name="Access VBA tests" tests="{0}" failures="{1}" errors="{2}" time="{3:0.000}">' -f $s.total, $s.failed, $s.errors, ($s.ms / 1000.0)))
    foreach ($g in @($Result.tests | Group-Object { $_.module })) {
        $tests = @($g.Group)
        $f = @($tests | Where-Object { $_.status -eq 'failed' }).Count
        $e = @($tests | Where-Object { $_.status -eq 'error' }).Count
        $ms = 0; foreach ($t in $tests) { $ms += [long]$t.ms }
        [void]$sb.AppendLine(('  <testsuite name="{0}" tests="{1}" failures="{2}" errors="{3}" time="{4:0.000}">' -f (& $esc $g.Name), $tests.Count, $f, $e, ($ms / 1000.0)))
        foreach ($t in $tests) {
            $attrs = 'classname="{0}" name="{1}" time="{2:0.000}"' -f (& $esc $t.module), (& $esc $t.name), ($t.ms / 1000.0)
            if ($t.file) { $attrs += ' file="{0}" line="{1}"' -f (& $esc $t.file), $t.line }
            if ($t.status -eq 'passed') { [void]$sb.AppendLine("    <testcase $attrs />"); continue }
            $tag = if ($t.status -eq 'failed') { 'failure' } else { 'error' }
            $detail = $t.message
            if ($t.location -and $t.location.code) { $detail += "`n  at $($t.location.module) line $($t.location.line): $($t.location.code.Trim())" }
            [void]$sb.AppendLine("    <testcase $attrs>")
            [void]$sb.AppendLine(('      <{0} message="{1}">{2}</{0}>' -f $tag, (& $esc $t.message), (& $esc $detail)))
            [void]$sb.AppendLine('    </testcase>')
        }
        [void]$sb.AppendLine('  </testsuite>')
    }
    [void]$sb.AppendLine('</testsuites>')
    [System.IO.File]::WriteAllText($Path, $sb.ToString(), (New-Object System.Text.UTF8Encoding($false)))
    return [System.IO.Path]::GetFullPath($Path)
}

function Add-TestSummaryMarkdown($Result, [string]$Path) {
    # GitHub Actions のジョブ概要（$GITHUB_STEP_SUMMARY）向けの Markdown
    $s = $Result.summary
    $icon = if ($s.failed + $s.errors) { '❌' } else { '✅' }
    $lines = @("## $icon Access VBA テスト: $($s.passed) / $($s.total) 成功", '')
    $bad = @($Result.tests | Where-Object { $_.status -ne 'passed' })
    if ($bad.Count) {
        $lines += '| 結果 | テスト | 内容 |'
        $lines += '|---|---|---|'
        foreach ($t in $bad) {
            $where = if ($t.location -and $t.location.line) { " ($($t.location.module) $($t.location.line) 行目)" } else { '' }
            $lines += "| $(if ($t.status -eq 'failed') { '失敗' } else { 'エラー' }) | ``$($t.id)`` | $(([string]$t.message) -replace '\|', '\|')$where |"
        }
    }
    [System.IO.File]::AppendAllText($Path, (($lines -join "`n") + "`n"), (New-Object System.Text.UTF8Encoding($false)))
}

function Add-SourceLocation($Loc, [string]$Src) {
    # { module, line } に ソースファイルと行番号 を付け足す
    if (-not $Loc -or -not $Loc.module) { return }
    $file = Find-SourceForModule $Src $Loc.module
    if ($file) {
        $Loc.file = $file
        $Loc.fileLine = Convert-CodeLineToFileLine $file ([int]$Loc.line)
    }
}

# ------------------------------------------------------------
# 公開関数：コマンドを 1 つ実行して結果（ハッシュテーブル）を返す
# ------------------------------------------------------------
function Invoke-AccessVsc([string]$Command, [hashtable]$P) {
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $res = [ordered]@{ ok = $false; command = $Command }
    $mutex = $null
    try {
        if ($Command -in @('lint', 'format', 'document')) {
            # Access も DB ファイルも不要（CI のクラウド実行環境でも動く）
            if ($P.src -and -not $P.db) { $src = [System.IO.Path]::GetFullPath($(if ([System.IO.Path]::IsPathRooted($P.src)) { $P.src } else { Join-Path (Get-Location).Path $P.src })) }
            else {
                $dbGuess = $P.db
                if ($dbGuess -and -not [System.IO.Path]::IsPathRooted($dbGuess)) { $dbGuess = Join-Path (Get-Location).Path $dbGuess }
                if (-not $dbGuess) { $dbGuess = Resolve-DatabasePath $null }
                $src = Resolve-SourcePath ([System.IO.Path]::GetFullPath($dbGuess)) $P.src
            }
            if (-not (Test-Path -LiteralPath $src)) { throw "Source folder not found: $src" }
            $res.source = $src
            if ($Command -eq 'document') {
                $out = $null
                if ($P.out) { $out = [System.IO.Path]::GetFullPath($(if ([System.IO.Path]::IsPathRooted([string]$P.out)) { [string]$P.out } else { Join-Path (Get-Location).Path ([string]$P.out) })) }
                $res.result = Invoke-SourceDocument $src $out
                $res.ok = $true
                return $res
            }
            if ($Command -eq 'format') {
                $res.result = Invoke-SourceFormat $src $(if ($P.files) { @($P.files) } else { $null }) ([bool]$P.check)
                $res.ok = -not ($P.check -and $res.result.changed.Count)
                if (-not $res.ok) { $res.error = "$($res.result.changed.Count) file(s) are not formatted" }
                return $res
            }
            $res.result = Invoke-SourceLint $src $(if ($P.files) { @($P.files) } else { $null })
            $res.ok = ($res.result.errorCount -eq 0)
            if (-not $res.ok) { $res.error = "$($res.result.errorCount) error(s), $($res.result.warningCount) warning(s)" }
            return $res
        }
        $db = Resolve-DatabasePath $P.db
        $src = Resolve-SourcePath $db $P.src
        $res.database = $db
        $res.source = $src
        $P.db = $db
        $P.src = $src
        $timeout = if ($P.timeoutSec) { [int]$P.timeoutSec } else { $script:DefaultTimeouts[$Command] }

        if ($Command -eq 'status') {
            $res.result = Get-SyncStatus $src
            $lockState = Get-DatabaseLockState $db
            $res.result.databaseLocked = ($lockState -eq 'other')
            $res.result.accessSession = ($lockState -eq 'session')
            $res.ok = $true
            return $res
        }
        if ($Command -eq 'setup-ci') {
            if (-not (Test-Path -LiteralPath $src)) { throw "Source folder not found: $src (run export first)" }
            $res.result = Invoke-SetupCi $db $src $P.root ([bool]$P.force) ([bool]$P.gitInit)
            $res.ok = $true
            return $res
        }
        if ($Command -eq 'deploy-init') {
            $res.result = Invoke-DeployInit $db $src $P
            $res.ok = $true
            return $res
        }
        if (-not $script:DefaultTimeouts.ContainsKey($Command)) { throw "Unknown command: $Command" }

        # 同じ DB に対する同時実行を防ぐ
        $mutexName = 'Local\AccessVsc_' + (Get-StringSha1 $db.ToLowerInvariant())
        $mutex = New-Object System.Threading.Mutex -ArgumentList $false, $mutexName
        try { $got = $mutex.WaitOne(120000) } catch [System.Threading.AbandonedMutexException] { $got = $true }
        if (-not $got) { $mutex = $null; throw 'Another Access VSC operation on this database is still running.' }

        $isWrite = ($Command -in $script:WriteCommands) -or ($Command -eq 'sql' -and $P.write)
        $lock = Get-DatabaseLockState $db
        if ($lock -eq 'session' -and -not ($script:Session -and $script:Session.db -eq $db)) {
            # ほかの常駐プロセス（VS Code / AI）が開いたままにしている：閉じてもらう
            [void](Request-SessionRelease $db)
            $lock = Get-DatabaseLockState $db
        }
        if ($lock -ne 'none' -and -not ($lock -eq 'session' -and $script:Session -and $script:Session.db -eq $db)) {
            if ($isWrite) { throw 'The database is open in another process (probably Access). Close it and retry.' }
            $res.warnings = @('The database is open in another process; results may not reflect unsaved changes.')
        }
        # セッション（開いたままの Access）を使うのは、ほかに使っている人がいないときだけ
        $useSession = [bool]$P.session -and $Command -in $script:SessionOps -and ($lock -eq 'none' -or ($lock -eq 'session' -and $script:Session -and $script:Session.db -eq $db))
        # 同じ DB を別の Access で開く操作（帳票スナップショットなど）の前に、開いたままの Access を閉じる。
        # テスト・ビルド・配布は DB のコピーを使うので閉じなくてよい
        if (-not $useSession -and $script:Session -and $script:Session.db -eq $db -and $Command -notin @('test', 'build', 'deploy')) { Close-AccessSession }

        if ($Command -in @('snapshot', 'build', 'deploy')) {
            # 複数回 Access を使う処理（AccessVsc.Release.ps1）
            Initialize-Win32
            switch ($Command) {
                'snapshot' { Invoke-SnapshotCommand $db $src $P $res }
                'build' { Invoke-BuildCommand $db $src $P $res }
                'deploy' { Invoke-DeployCommand $db $src $P $res }
            }
            return $res
        }

        switch ($Command) {
            'export' {
                if (-not $P.force) {
                    $st = Get-SyncStatus $src
                    $dirty = @($st.changed) + @($st.added) + @($st.missing)
                    if ($dirty.Count -gt 0) {
                        $res.result = [ordered]@{ unimported = $dirty }
                        throw ("Source folder has changes that are not imported yet; export would overwrite them. Run import first, or export with force. Files: " + ($dirty -join ', '))
                    }
                }
                if (-not (Test-Path -LiteralPath $src)) { [void](New-Item -ItemType Directory -Path $src -Force) }
            }
            'import' {
                $st = Get-SyncStatus $src
                $files = @()
                if ($P.files) {
                    foreach ($f in @($P.files)) {
                        $full = if ([System.IO.Path]::IsPathRooted($f)) { $f } else { Join-Path $src $f }
                        $rel = Get-RelPath $src $full
                        if (-not $rel) { throw "File is outside the source folder ($src): $f" }
                        if (-not (Test-Path -LiteralPath $full)) { throw "File not found: $full" }
                        $files += $rel
                    }
                }
                elseif ($P.all) { $files = @(Get-SourceFiles $src | Where-Object { -not (Get-ObjectInfoFromRelPath $_).readOnly }) }
                else { $files = @($st.changed) + @($st.added) }
                $P.files = $files
                $P.deleteFiles = $(if ($P.deleteMissing) { @($st.missing) } else { @() })
                if ($null -eq $P.compile) { $P.compile = $true }
                if ($files.Count -eq 0 -and $P.deleteFiles.Count -eq 0) {
                    $res.ok = $true
                    $res.result = [ordered]@{ imported = @(); deleted = @(); errors = @(); compile = $null; missingNotDeleted = @($st.missing) }
                    $res.message = 'Nothing to import (no changed files).'
                    return $res
                }
                # 並び順：モジュール -> クエリ -> フォーム/レポート -> マクロ（フォームのコードがモジュールを参照するため）
                $order = @{ modules = 0; queries = 1; forms = 2; reports = 3; macros = 4 }
                $P.files = @($files | Sort-Object { $i = Get-ObjectInfoFromRelPath $_; if ($i -and $order.ContainsKey($i.kind)) { $order[$i.kind] } else { 9 } }, { $_ })
            }
            'compile' { $P.forceCompile = $true }
            'run' { if (-not $P.function) { throw 'function is required' } }
            'eval' { if (-not $P.expression) { throw 'expression is required' } }
            'sql' { if (-not $P.sql) { throw 'sql is required' } }
            'test' {
                $cases = @(Get-TestCases $src)
                if ($P.filter) {
                    $pats = @($P.filter | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
                    $cases = @($cases | Where-Object { $c = $_; @($pats | Where-Object { $c.id -like $_ -or $c.name -like $_ -or $c.module -like $_ }).Count -gt 0 })
                }
                if ($cases.Count -eq 0) {
                    $res.ok = $true
                    $res.result = [ordered]@{ summary = [ordered]@{ total = 0; passed = 0; failed = 0; errors = 0; ms = 0 }; tests = @() }
                    $res.message = 'No tests found. Tests are Public Sub/Function named Test_* (no arguments) in modules/*.bas.'
                    return $res
                }
                $st = Get-SyncStatus $src
                $pending = @(@($st.changed) + @($st.added) | Where-Object { $_ -notmatch '^(tables/|relationships)' })
                if ($pending.Count) { $res.warnings = @("Unimported source changes are NOT included in this test run (tests run against the database): $($pending -join ', ')") }
                # テストは DB のコピーで実行（本番の DB を汚さない）
                $testDir = Join-Path (Split-Path -Parent $db) '.accessvsc\test'
                if (-not (Test-Path -LiteralPath $testDir)) { [void](New-Item -ItemType Directory -Path $testDir -Force) }
                $copy = Join-Path $testDir ([System.IO.Path]::GetFileNameWithoutExtension($db) + '_test' + [System.IO.Path]::GetExtension($db))
                Copy-Item -LiteralPath $db -Destination $copy -Force
                $P.db = $copy
                $P.tests = $cases
                $res.testDatabase = $copy
            }
            { $_ -in @('layout', 'preview', 'pagesetup') } {
                if (-not $P.report) { throw 'report (object name) is required' }
                $kindDir = if ($P.objectType -eq 'form') { 'forms' } else { 'reports' }
                $rel = "$kindDir/" + (ConvertTo-SafeFileName ([string]$P.report)) + '.txt'
                $st = Get-SyncStatus $src
                if (@($st.changed) + @($st.added) -contains $rel) {
                    if ($Command -eq 'pagesetup') { throw "$rel has unimported changes. Import it first (page setup rewrites this file)." }
                    $res.warnings = @("$rel has unimported changes; the result shows the version currently in the database. Import first to see your edits.")
                }
                $P.outDir = Join-Path (Split-Path -Parent $db) '.accessvsc\preview'
            }
        }

        if ($isWrite -and -not $P.noBackup) { $res.backup = New-DatabaseBackup $db }

        $out = Invoke-InAccess $Command $P $timeout $useSession
        if ($useSession) { $res.session = $true }
        $res.ok = [bool]$out.ok
        $res.result = $out.result
        if ($out.error) { $res.error = $out.error }
        if ($out.warnings -and @($out.warnings).Count) { $res.warnings = @(@($res.warnings) + @($out.warnings) | Where-Object { $_ }) }
        if ($out.dialogs -and @($out.dialogs).Count) { $res.dialogs = $out.dialogs }
        if ($out.runtimeError) {
            Add-SourceLocation $out.runtimeError $src
            $res.runtimeError = $out.runtimeError
        }
        $compileErr = $null
        if ($Command -eq 'import' -and $res.result -and $res.result.compile) { $compileErr = $res.result.compile.error }
        if ($Command -eq 'compile' -and $res.result) { $compileErr = $res.result.error }
        if ($compileErr) {
            Add-SourceLocation $compileErr $src
            # 見張り役が拾ったコンパイルエラーのダイアログ文言を、エラー内容として使う
            $cd = @($out.dialogs) | Where-Object { $_ -and $_.message -match 'コンパイル エラー|Compile error' } | Select-Object -First 1
            if ($cd) { $compileErr.message = $cd.message }
        }
        if ($Command -eq 'import' -and $res.result -and -not $P.deleteMissing) {
            $miss = @((Get-SyncStatus $src).missing)
            if ($miss.Count) { $res.result.missingNotDeleted = $miss }
        }
        if ($Command -eq 'test' -and $res.result -and $res.result.tests) {
            $byId = @{}; foreach ($c in @($P.tests)) { $byId[$c.id] = $c }
            foreach ($t in @($res.result.tests)) {
                $c = $byId[$t.id]
                if ($c) { $t.file = $c.file; $t.line = $c.line }
                if ($t.location) { Add-SourceLocation $t.location $src }
            }
            $res.result.junit = Write-JUnitReport $res.result $(if ($P.junit) { $P.junit } else { Join-Path (Split-Path -Parent $db) '.accessvsc\test-results\junit.xml' })
            if ($P.summary) { Add-TestSummaryMarkdown $res.result $P.summary }
        }
        if ($Command -eq 'test' -and $res.result -and $res.result.compile -and $res.result.compile.error) { Add-SourceLocation $res.result.compile.error $src }
        if ($Command -eq 'preview' -and $res.ok -and $res.result.pdf -and (Test-Path -LiteralPath $res.result.pdf) -and $P.png -ne $false) {
            $maxPages = if ($P.pages) { [int]$P.pages } else { 3 }
            $img = Convert-PdfToPng $res.result.pdf $maxPages
            if ($img.error) { $res.warnings = @(@($res.warnings) + "PNG conversion failed: $($img.error)" | Where-Object { $_ }) }
            else {
                $res.result.pageCount = $img.pageCount
                $res.result.images = @($img.images)
            }
        }
    }
    catch {
        $res.ok = $false
        $res.error = $_.Exception.Message
    }
    finally {
        if ($mutex) { try { $mutex.ReleaseMutex() } catch {}; $mutex.Dispose() }
        $res.elapsedMs = $sw.ElapsedMilliseconds
    }
    return $res
}

function ConvertTo-HashtableDeep($obj) {
    # ConvertFrom-Json の結果（PSCustomObject）をハッシュテーブルに
    if ($null -eq $obj) { return $null }
    if ($obj -is [System.Management.Automation.PSCustomObject]) {
        $h = @{}
        foreach ($p in $obj.PSObject.Properties) { $h[$p.Name] = ConvertTo-HashtableDeep $p.Value }
        return $h
    }
    if ($obj -is [array]) { return , @($obj | ForEach-Object { ConvertTo-HashtableDeep $_ }) }
    return $obj
}
