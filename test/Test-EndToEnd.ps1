<#
  Test-EndToEnd.ps1
  一時フォルダにテスト DB を作り、CLI の主要な流れを通しで確認する。
  使い方: powershell -NoProfile -ExecutionPolicy Bypass -File .\test\Test-EndToEnd.ps1
#>
$ErrorActionPreference = 'Stop'
$here = $PSScriptRoot
$cli = Join-Path $here '..\scripts\access.ps1'
$work = Join-Path ([IO.Path]::GetTempPath()) ('accessvsc_e2e_' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $work | Out-Null
$db = Join-Path $work 'shop.accdb'
$src = "$db.src"
$u8 = New-Object System.Text.UTF8Encoding($false)
$script:fail = 0

function Acc([string]$Command, [string[]]$More) {
    $json = & powershell -NoProfile -ExecutionPolicy Bypass -File $cli $Command -DbPath $db -Compress @More
    return ($json -join "`n") | ConvertFrom-Json
}
function Check([string]$Name, [bool]$Cond, $Detail) {
    if ($Cond) { Write-Output "PASS $Name" }
    else { $script:fail++; Write-Output "FAIL $Name"; if ($Detail) { Write-Output ("     " + ($Detail | ConvertTo-Json -Depth 8 -Compress)) } }
}
function Put([string]$Rel, [string]$Text) { $p = Join-Path $src $Rel; [IO.File]::WriteAllText($p, $Text, $u8) }

try {
    & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $here 'New-TestDatabase.ps1') -Path $db | Out-Null

    $r = Acc export; Check 'export' ($r.ok -and $r.result.objects -ge 6) $r
    $r = Acc export; Check 'export is stable' ($r.ok -and @($r.result.written).Count -eq 0) $r
    $r = Acc status; Check 'status clean' ($r.ok -and @($r.result.changed).Count -eq 0) $r

    Put 'modules/modE2E.bas' "Option Compare Database`r`nOption Explicit`r`n`r`nPublic Function E2E() As Long`r`n    E2E = missingVar`r`nEnd Function`r`n"
    $r = Acc export; Check 'export refuses with unimported changes' (-not $r.ok -and @($r.result.unimported) -contains 'modules/modE2E.bas') $r
    $r = Acc import; Check 'import reports compile error location' (-not $r.ok -and $r.result.compile.error.fileLine -eq 5 -and $r.result.compile.error.module -eq 'modE2E') $r
    Check 'compile error has message' ($r.result.compile.error.message -match 'コンパイル|Compile') $r.result.compile.error

    Put 'modules/modE2E.bas' "Option Compare Database`r`nOption Explicit`r`n`r`nPublic Function E2E(ByVal a As Long, ByVal s As String) As String`r`n    E2E = s & (a * 2)`r`nEnd Function`r`n`r`nPublic Function E2EDiv() As Long`r`n    Dim z As Long`r`n    E2EDiv = 1 / z`r`nEnd Function`r`n`r`nPublic Function E2EMsg() As Long`r`n    MsgBox ""テスト""`r`n    E2EMsg = 1`r`nEnd Function`r`n"
    $r = Acc import; Check 'import + compile ok' ($r.ok -and $r.result.compile.compiled) $r
    $r = Acc run @('-Function', 'E2E', '-ArgsJson', "[21, '答え=']"); Check 'run with args' ($r.ok -and $r.result.value -eq '答え=42') $r
    $r = Acc run @('-Function', 'E2EDiv'); Check 'runtime error location' (-not $r.ok -and $r.runtimeError.number -eq 11 -and $r.runtimeError.fileLine -eq 10) $r
    $r = Acc run @('-Function', 'E2EMsg'); Check 'msgbox auto-answered' ($r.ok -and $r.result.value -eq 1 -and @($r.dialogs).Count -eq 1) $r
    $r = Acc sql @('-Sql', 'SELECT CustomerName FROM Customers ORDER BY ID'); Check 'sql select' ($r.ok -and $r.result.rows[0].CustomerName -eq '山田商店') $r
    $r = Acc sql @('-Sql', 'DELETE FROM Orders'); Check 'sql read-only guard' (-not $r.ok) $r
    $r = Acc eval @('-Expr', "DCount('*','Customers')"); Check 'eval' ($r.ok -and $r.result.value -eq 3) $r

    Put 'queries/qryE2E.sql' 'SELECT Count(*) AS N FROM Customers'
    $f = Join-Path $src 'forms\frmCustomers.txt'
    [IO.File]::WriteAllText($f, ([IO.File]::ReadAllText($f).Replace('Caption ="顧客"', 'Caption ="顧客E2E"')), $u8)
    $r = Acc import; Check 'import query + form' ($r.ok -and @($r.result.imported).Count -eq 2) $r
    $r = Acc sql @('-Sql', 'qryE2E'); Check 'new query runs' ($r.ok -and $r.result.rows[0].N -eq 3) $r

    Remove-Item (Join-Path $src 'queries\qryE2E.sql')
    $r = Acc import @('-DeleteMissing'); Check 'delete missing' ($r.ok -and @($r.result.deleted) -contains 'queries/qryE2E.sql') $r

    $r = Acc export; Check 'export after import' ($r.ok) $r
    Check 'form edit survived round trip' ([IO.File]::ReadAllText($f) -match '顧客E2E') $null
    $r = Acc run @('-Function', 'E2E', '-ArgsJson', "[1, 'x']", '-TimeoutSec', '60'); Check 'still runs after round trip' ($r.ok -and $r.result.value -eq 'x2') $r

    # --- 帳票（レポート） ---
    $r = Acc layout @('-Report', 'rptCustomers')
    Check 'report layout' ($r.ok -and @($r.result.controls).Count -eq 3 -and $r.result.printer.paperSize -eq 'A4' -and @($r.result.issues).Count -eq 0) $r
    $rf = Join-Path $src 'reports\rptCustomers.txt'
    $rt = [IO.File]::ReadAllText($rf)
    Check 'report source has twip positions' ($rt -match 'Left =4700' -and $rt -match 'Name ="txtCity"') $null
    [IO.File]::WriteAllText($rf, $rt.Replace('Left =4700', 'Left =4000').Replace('Caption ="顧客一覧"', 'Caption ="顧客一覧表"'), $u8)
    $r = Acc preview @('-Report', 'rptCustomers'); Check 'preview warns about unimported report' ($r.ok -and (@($r.warnings) -join ' ') -match 'unimported') $r
    $r = Acc import; Check 'import report' ($r.ok -and @($r.result.imported) -contains 'reports/rptCustomers.txt') $r
    $r = Acc layout @('-Report', 'rptCustomers')
    $city = @($r.result.controls) | Where-Object { $_.name -eq 'txtCity' }
    Check 'report edit applied' ($city.left -eq 4000 -and (@($r.result.controls | Where-Object { $_.name -eq 'lblTitle' })[0].caption -eq '顧客一覧表')) $r.result.controls
    Check 'overlap detected' ((@($r.result.issues) -join ' ') -match 'txtName と txtCity が重なっています') $r.result.issues
    $r = Acc preview @('-Report', 'rptCustomers', '-Where', "City='東京'")
    Check 'preview pdf + png' ($r.ok -and (Test-Path $r.result.pdf) -and @($r.result.images).Count -ge 1 -and (Test-Path $r.result.images[0]) -and $r.result.pageCount -ge 1) $r
    $r = Acc pagesetup @('-Report', 'rptCustomers', '-PaperSize', 'B5', '-Orientation', 'portrait', '-Margins', '20,20,40,40')
    Check 'pagesetup applied' ($r.ok -and $r.result.printer.paperSize -eq 'B5' -and $r.result.printer.marginsMm.left -eq 40) $r
    Check 'too wide for paper detected' ((@($r.result.issues) -join ' ') -match '印刷可能幅') $r.result.issues
    $r = Acc layout @('-Report', 'rptCustomers'); Check 'pagesetup persisted' ($r.result.printer.paperSize -eq 'B5' -and $r.result.printer.orientation -eq 'portrait') $r.result.printer
    $r = Acc status; Check 'source in sync after pagesetup' ($r.ok -and @($r.result.changed).Count -eq 0) $r

    # --- エクスポートで、DB にないオブジェクトのファイルが消えること ---
    Put 'queries/qryGhost.sql' 'SELECT 1 AS X'
    $r = Acc export @('-Force'); Check 'export removes stale files' ($r.ok -and @($r.result.deleted) -contains 'queries/qryGhost.sql' -and -not (Test-Path (Join-Path $src 'queries\qryGhost.sql'))) $r

    # --- テスト・CI の準備 ---
    $r = Acc setup-ci @('-Root', $work)
    Check 'setup-ci' ($r.ok -and (Test-Path (Join-Path $src 'modules\Assert.bas')) -and (Test-Path (Join-Path $src 'modules\Test_Example.bas')) -and (Test-Path (Join-Path $work '.github\workflows\access-ci.yml')) -and (Test-Path (Join-Path $work 'tools\access-vsc\access.ps1'))) $r
    $wf = [IO.File]::ReadAllText((Join-Path $work '.github\workflows\access-ci.yml'))
    Check 'workflow paths filled' ($wf -match 'tools/access-vsc/access.ps1' -and $wf -match '-DbPath "shop.accdb"' -and $wf -notmatch '\{\{') $null

    # --- lint ---
    $r = Acc lint; Check 'lint clean' ($r.ok -and $r.result.errorCount -eq 0 -and $r.result.files -ge 8) $r
    Put 'modules/modEmoji.bas' "Option Compare Database`r`nOption Explicit`r`n`r`nPublic Function Smile() As String`r`n    Smile = `"$([char]::ConvertFromUtf32(0x1F600))`"`r`nEnd Function`r`n"
    $r = Acc lint; Check 'lint detects non Shift-JIS char' (-not $r.ok -and @($r.result.problems | Where-Object { $_.rule -eq 'sjis' -and $_.file -eq 'modules/modEmoji.bas' -and $_.line -eq 5 }).Count -eq 1) $r.result
    Remove-Item (Join-Path $src 'modules\modEmoji.bas')

    # --- ユニットテスト ---
    Put 'modules/Test_Shop.bas' "Option Compare Database`r`nOption Explicit`r`n`r`nPublic Sub Test_Greet()`r`n    Assert.AreEqual `"こんにちは、A`", Greet(`"A`")`r`nEnd Sub`r`n`r`nPublic Sub Test_Wrong()`r`n    Assert.AreEqual 5, AddNumbers(2, 2)`r`nEnd Sub`r`n`r`nPublic Sub Test_Crash()`r`n    Dim z As Long`r`n    Debug.Print 1 / z`r`nEnd Sub`r`n`r`nPublic Sub Test_DeleteAll()`r`n    CurrentDb.Execute `"DELETE FROM Orders`", dbFailOnError`r`n    Assert.AreEqual 0, DCount(`"*`", `"Orders`")`r`nEnd Sub`r`n"
    $r = Acc import; Check 'import test modules' ($r.ok) $r
    $junit = Join-Path $work 'junit.xml'
    $r = Acc test @('-JUnit', $junit)
    $s = $r.result.summary
    Check 'test summary' (-not $r.ok -and $s.total -eq 7 -and $s.passed -eq 5 -and $s.failed -eq 1 -and $s.errors -eq 1) $s
    $crash = @($r.result.tests | Where-Object { $_.name -eq 'Test_Crash' })[0]
    Check 'test runtime error location' ($crash.status -eq 'error' -and $crash.location.fileLine -eq 14) $crash
    $wrong = @($r.result.tests | Where-Object { $_.name -eq 'Test_Wrong' })[0]
    Check 'test assertion message' ($wrong.status -eq 'failed' -and $wrong.message -match '期待値 5') $wrong
    Check 'junit written' ((Test-Path $junit) -and ([xml](Get-Content -Raw -Encoding UTF8 $junit)).testsuites.failures -eq '1') $null
    $r = Acc sql @('-Sql', 'SELECT Count(*) AS N FROM Orders'); Check 'tests ran on a copy (real data untouched)' ($r.result.rows[0].N -eq 1) $r
    $r = Acc test @('-Filter', 'Test_Example.*'); Check 'test filter' ($r.ok -and $r.result.summary.total -eq 3) $r.result.summary

    # --- 帳票スナップショット ---
    $r = Acc snapshot; Check 'snapshot without baseline = new' ($r.ok -and $r.result.summary.new -eq 1) $r.result.summary
    $r = Acc snapshot @('-Update'); Check 'snapshot update' ($r.ok -and $r.result.summary.updated -eq 1 -and (Test-Path (Join-Path $src 'snapshots\rptCustomers_p1.png'))) $r
    $r = Acc snapshot; Check 'snapshot same' ($r.ok -and $r.result.summary.same -eq 1) $r.result
    $rf = Join-Path $src 'reports\rptCustomers.txt'
    [IO.File]::WriteAllText($rf, ([IO.File]::ReadAllText($rf) -replace 'Left =4000', 'Left =5000'), $u8)
    [void](Acc import)
    $r = Acc snapshot
    $pg = @($r.result.reports[0].pages | Where-Object { $_.status -eq 'changed' })
    Check 'snapshot detects layout change' (-not $r.ok -and $r.result.summary.changed -eq 1 -and $pg.Count -ge 1 -and (Test-Path $pg[0].diff)) $r.result
    $r = Acc snapshot @('-Threshold', '50'); Check 'snapshot threshold' ($r.ok) $r.result.summary

    # --- リリース（.accde 作成・バックエンド張り替え・配布） ---
    & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $here 'New-TestBackend.ps1') -FrontEnd $db -Dir $work | Out-Null
    $share = Join-Path $work 'share'
    New-Item -ItemType Directory -Path $share | Out-Null
    $r = Acc deploy-init @('-AppName', 'ShopApp', '-DeployTo', $share, '-BackendPath', (Join-Path $work 'shop_be_prod.accdb'), '-Root', $work)
    Check 'deploy-init' ($r.ok -and (Test-Path (Join-Path $work 'access-deploy.json')) -and (Test-Path (Join-Path $work '.github\workflows\access-release.yml'))) $r
    $r = Acc build @('-Version', '2.0.0')
    Check 'build accde' ($r.ok -and (Test-Path $r.result.accde) -and $r.result.accde -like '*ShopApp_2.0.0.accde') $r
    Check 'build relinked backend' (@($r.result.relinked).Count -eq 1 -and $r.result.relinked[0].to -like '*shop_be_prod.accdb') $r.result
    $accde = $r.result.accde
    $r = Acc deploy; Check 'deploy dry run writes nothing' ($r.ok -and $r.result.dryRun -and @(Get-ChildItem $share).Count -eq 0) $r
    $r = Acc deploy @('-Yes', '-From', $accde)
    Check 'deploy publish' ($r.ok -and (Test-Path (Join-Path $share 'app\ShopApp.accde')) -and (Test-Path (Join-Path $share 'Start-ShopApp.cmd')) -and (Test-Path (Join-Path $share 'releases\2.0.0\ShopApp.accde'))) $r
    $ver = [IO.File]::ReadAllText((Join-Path $share 'version.json')) | ConvertFrom-Json
    Check 'version.json' ($ver.version -eq '2.0.0' -and $ver.sha256 -eq (Get-FileHash (Join-Path $share 'app\ShopApp.accde') -Algorithm SHA256).Hash.ToLower()) $ver
    $fakeLocal = Join-Path $work 'localappdata'
    $savedLocal = $env:LOCALAPPDATA
    try { $env:LOCALAPPDATA = $fakeLocal; $lo = & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $share 'launcher.ps1') -NoStart } finally { $env:LOCALAPPDATA = $savedLocal }
    $localAccde = Join-Path $fakeLocal 'AccessApps\ShopApp\ShopApp.accde'
    Check 'launcher copies latest version' ((Test-Path $localAccde) -and ($lo -join ' ') -match '2\.0\.0') $lo
    $dbSaved = $db; $db = $localAccde
    try {
        $r = Acc sql @('-Sql', 'SELECT Count(*) AS N FROM Products'); Check 'accde uses production backend' ($r.ok -and $r.result.rows[0].N -eq 3) $r
        $r = Acc run @('-Function', 'Greet', '-ArgsJson', "['accde']"); Check 'accde runs VBA' ($r.ok -and $r.result.value -eq 'こんにちは、accde') $r
    }
    finally { $db = $dbSaved }
}
finally {
    $left = @(Get-CimInstance Win32_Process -Filter "Name='MSACCESS.EXE'" | Where-Object { $_.CommandLine -match '-Embedding' -and $_.CreationDate -gt (Get-Item $work).CreationTime })
    Write-Output "leftover automation Access processes: $($left.Count)"
    Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
}
if ($script:fail) { Write-Output "$($script:fail) test(s) failed"; exit 1 }
Write-Output 'all tests passed'
