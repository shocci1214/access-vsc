<#
  mcp-server.ps1
  Access VSC の MCP サーバー（stdio / JSON-RPC 2.0、1 行 1 メッセージ）。
  Claude Code・GitHub Copilot などの AI エージェントから Access を操作するための窓口。

  登録例（Claude Code の .mcp.json）:
    {
      "mcpServers": {
        "access": {
          "command": "powershell.exe",
          "args": ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "<このファイル>", "-DbPath", "C:\\work\\app.accdb"]
        }
      }
    }
#>
param(
    [string]$DbPath,
    [string]$Src
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'AccessVsc.Runner.ps1')

$script:Version = '0.7.0'
$utf8 = New-Object System.Text.UTF8Encoding($false)
$stdin = New-Object System.IO.StreamReader([Console]::OpenStandardInput(), $utf8)
$stdout = New-Object System.IO.StreamWriter([Console]::OpenStandardOutput(), $utf8)
$stdout.AutoFlush = $true
$stdout.NewLine = "`n"

function Write-Log([string]$Message) { [Console]::Error.WriteLine("[access-mcp] $Message") }
function Send-Message($Obj) { $stdout.WriteLine((ConvertTo-Json -InputObject $Obj -Depth 30 -Compress)) }
function Send-Result($Id, $Result) { Send-Message ([ordered]@{ jsonrpc = '2.0'; id = $Id; result = $Result }) }
function Send-Error($Id, [int]$Code, [string]$Message) { Send-Message ([ordered]@{ jsonrpc = '2.0'; id = $Id; error = [ordered]@{ code = $Code; message = $Message } }) }

try { Initialize-Win32 } catch { Write-Log "Win32 helper load failed: $($_.Exception.Message)" }

# ------------------------------------------------------------
# ツール定義
# ------------------------------------------------------------
$dbProp = [ordered]@{ type = 'string'; description = 'Path to the .accdb/.mdb file. Optional when the server was started with a default database.' }
function New-Tool([string]$Name, [string]$Title, [string]$Description, $Props, [string[]]$Required, [bool]$ReadOnly = $false, [bool]$Destructive = $true) {
    $schemaProps = [ordered]@{}
    foreach ($k in $Props.Keys) { $schemaProps[$k] = $Props[$k] }
    $schemaProps['database'] = $dbProp
    $schema = [ordered]@{ type = 'object'; properties = $schemaProps }
    if ($Required) { $schema.required = @($Required) }
    $t = [ordered]@{ name = $Name; title = $Title; description = $Description; inputSchema = $schema }
    # MCP annotations: read-only tools never change anything; the others may overwrite the database or source files.
    if ($ReadOnly) { $t.annotations = [ordered]@{ readOnlyHint = $true } }
    else { $t.annotations = [ordered]@{ readOnlyHint = $false; destructiveHint = $Destructive } }
    return $t
}

$tools = @(
    (New-Tool 'access_status' 'Source status' 'Show which source files were changed, added or deleted since the last export/import, and whether the database is locked. Fast: does not start Access. Call this first.' ([ordered]@{}) @() $true)
    (New-Tool 'access_list' 'List objects' 'List tables, queries, forms, reports, macros and VBA modules in the database, and whether the VBA project is compiled.' ([ordered]@{}) @() $true)
    (New-Tool 'access_export' 'Export Access to source' 'Export all objects from the database to the source folder (modules/*.bas|cls, forms/*.txt, reports/*.txt, macros/*.txt, queries/*.sql, tables/*.json read-only schema). Refuses when there are unimported source changes unless force=true (which overwrites them).' ([ordered]@{
                force = [ordered]@{ type = 'boolean'; description = 'Overwrite unimported changes in the source folder.' }
                kinds = [ordered]@{ type = 'array'; items = [ordered]@{ type = 'string'; enum = @('modules', 'forms', 'reports', 'macros', 'queries', 'tables') }; description = 'Only export these kinds.' }
            }) @())
    (New-Tool 'access_import' 'Import source to Access' 'Import source files into the database, then compile the VBA project. Default: only files changed/added since the last sync. Returns compile errors with file and fileLine so you can fix them. A backup of the database is made first.' ([ordered]@{
                files         = [ordered]@{ type = 'array'; items = [ordered]@{ type = 'string' }; description = 'Files to import, relative to the source folder (e.g. "modules/modMain.bas") or absolute.' }
                all           = [ordered]@{ type = 'boolean'; description = 'Import every source file.' }
                deleteMissing = [ordered]@{ type = 'boolean'; description = 'Delete objects whose source files were deleted.' }
                compile       = [ordered]@{ type = 'boolean'; description = 'Compile after import (default true).' }
            }) @())
    (New-Tool 'access_compile' 'Compile VBA' 'Compile and save all VBA modules. On failure returns the module, line, code and the source file/line of the first error.' ([ordered]@{}) @() $false $false)
    (New-Tool 'access_run' 'Run VBA procedure' 'Run a Public Function or Sub in a standard module via Application.Run and return its value. MsgBox and other dialogs are answered automatically and reported: OK-only messages get OK; confirmations get the safe answer (Cancel, then No), so code that asks "Delete?" will NOT proceed. VBA runtime errors are returned with module, line and source file/line.' ([ordered]@{
                function = [ordered]@{ type = 'string'; description = 'Procedure name, e.g. "modMain.Greet" or "Greet".' }
                args     = [ordered]@{ type = 'array'; items = @{}; description = 'Arguments (strings, numbers, booleans).' }
            }) @('function'))
    (New-Tool 'access_eval' 'Evaluate expression' 'Evaluate an Access expression with Application.Eval, e.g. "DCount(''*'',''Customers'')" or "MyFunc(1)".' ([ordered]@{
                expression = [ordered]@{ type = 'string' }
            }) @('expression'))
    (New-Tool 'access_sql' 'Query data' 'Run Access SQL (or a saved query name) and return rows as JSON. Read-only by default; set write=true to execute INSERT/UPDATE/DELETE/DDL (a backup is made first).' ([ordered]@{
                sql   = [ordered]@{ type = 'string' }
                limit = [ordered]@{ type = 'integer'; description = 'Max rows (default 100).' }
                write = [ordered]@{ type = 'boolean'; description = 'Execute an action query.' }
            }) @('sql'))
    (New-Tool 'access_report_layout' 'Report layout' 'Get the layout of a report (or form): report width, sections, every control with position/size (twips and cm), caption, control source, font, the printer settings (paper, orientation, margins, printable width) and detected issues (controls overflowing the width or section, overlapping controls, report wider than the printable width). Use this before and after editing reports/*.txt. 1 cm = 567 twips.' ([ordered]@{
                report     = [ordered]@{ type = 'string'; description = 'Report (or form) name.' }
                objectType = [ordered]@{ type = 'string'; enum = @('report', 'form'); description = 'Default: report.' }
            }) @('report') $true)
    (New-Tool 'access_report_preview' 'Preview report' 'Render a report (or form) to PDF and PNG and return the page images so you can SEE the printed layout. Also returns layout issues. Use after access_import to check layout changes visually.' ([ordered]@{
                report     = [ordered]@{ type = 'string'; description = 'Report (or form) name.' }
                where      = [ordered]@{ type = 'string'; description = 'Optional filter (SQL WHERE without the word WHERE), e.g. "City=''Tokyo''".' }
                pages      = [ordered]@{ type = 'integer'; description = 'Max pages to render as images (default 2).' }
                objectType = [ordered]@{ type = 'string'; enum = @('report', 'form') }
            }) @('report') $true)
    (New-Tool 'access_report_pagesetup' 'Report page setup' 'Change a report''s paper size, orientation, margins (mm) and/or width (cm). Saves the report and rewrites reports/<name>.txt so the source stays in sync. Paper sizes and margins cannot be edited in the text file (they are binary), so use this tool.' ([ordered]@{
                report       = [ordered]@{ type = 'string' }
                paperSize    = [ordered]@{ type = 'string'; enum = @('A3', 'A4', 'A5', 'B4', 'B5', 'Letter', 'Legal') }
                orientation  = [ordered]@{ type = 'string'; enum = @('portrait', 'landscape') }
                marginTop    = [ordered]@{ type = 'number'; description = 'mm' }
                marginBottom = [ordered]@{ type = 'number'; description = 'mm' }
                marginLeft   = [ordered]@{ type = 'number'; description = 'mm' }
                marginRight  = [ordered]@{ type = 'number'; description = 'mm' }
                widthCm      = [ordered]@{ type = 'number'; description = 'Report width in cm' }
            }) @('report'))
    (New-Tool 'access_test' 'Run VBA unit tests' 'Run VBA unit tests on a COPY of the database (the real database is not modified). Tests are Public Sub Test_*() with no arguments in modules/*.bas; optional TestSetup/TestTeardown in the same module run around each test. Assertions: Assert.AreEqual expected, actual / Assert.IsTrue cond / Assert.AreClose / Assert.Contains / Assert.Fail. Returns pass/fail per test with messages; runtime errors include the source file/line. Import changed files first (tests run against the database).' ([ordered]@{
                filter = [ordered]@{ type = 'array'; items = [ordered]@{ type = 'string' }; description = 'Wildcards matched against "Module.Test_Name", the test name or the module name, e.g. ["Test_Tax*"].' }
            }) @() $false $false)
    (New-Tool 'access_report_snapshot' 'Report visual regression' 'Render reports to PNG and compare them with the approved baseline images in <src>/snapshots. Returns same/changed/new per report and the diff images (changed pixels in red) of changed pages. After an intended layout change, call again with update=true to accept the new look (only when the user agrees the change is intended).' ([ordered]@{
                reports   = [ordered]@{ type = 'array'; items = [ordered]@{ type = 'string' }; description = 'Report names or wildcards (default: all reports).' }
                update    = [ordered]@{ type = 'boolean'; description = 'Accept the current look as the new baseline.' }
                threshold = [ordered]@{ type = 'number'; description = 'Ignore pages whose changed pixels are at most this percentage (default 0). Use e.g. 0.1 when reports print the current date.' }
            }) @())
    (New-Tool 'access_build' 'Build release ACCDE' 'Build a release .accde from the database: copy, stamp AppVersion, relink linked tables to the production back end in access-deploy.json, compile and make the ACCDE into the output folder. Does NOT publish; publishing (deploy) is done by a person from VS Code or the CLI.' ([ordered]@{
                version = [ordered]@{ type = 'string'; description = 'Version, e.g. 1.2.0 (default: date and time).' }
            }) @() $false $false)
    (New-Tool 'access_lint' 'Lint source' 'Static checks of the source folder without Access: characters not representable in Shift-JIS, missing Option Explicit, leftover Stop, report/form controls overflowing width or section, empty SQL. Fast.' ([ordered]@{}) @() $true)
    (New-Tool 'access_format' 'Format VBA' 'Fix indentation (4 spaces per block: Sub/Function, If, For, Do, While, With, Select Case, Type) and trailing spaces of VBA in modules/*.bas|cls and the code part of forms/reports .txt. Changes only whitespace; files whose blocks do not match (missing End If etc.) are left unchanged and listed in "skipped". Does not start Access. Run it after editing VBA and before access_import.' ([ordered]@{
                files = [ordered]@{ type = 'array'; items = [ordered]@{ type = 'string' }; description = 'Only these files, relative to the source folder (e.g. "modules/modMain.bas"). Default: all.' }
                check = [ordered]@{ type = 'boolean'; description = 'Only report which files are not formatted; do not write.' }
            }) @() $false $false)
    (New-Tool 'access_setup_ci' 'Set up tests and CI' 'Add the Assert module and an example test module to the source folder, copy the tool scripts to tools/access-vsc, and create a GitHub Actions workflow (.github/workflows/access-ci.yml), a pre-commit lint hook and .gitattributes/.gitignore entries. Does not start Access.' ([ordered]@{
                root    = [ordered]@{ type = 'string'; description = 'Repository root (default: git root of the database folder, or the database folder).' }
                gitInit = [ordered]@{ type = 'boolean'; description = 'Run git init when the root is not a git repository.' }
                force   = [ordered]@{ type = 'boolean'; description = 'Overwrite an existing workflow / hook.' }
            }) @())
)

$instructions = @'
Access VSC lets you edit a Microsoft Access database as text files.
Workflow: access_status -> edit files in the source folder -> access_format (fix VBA indentation) -> access_import (imports changed files and compiles) -> fix compile errors at the returned file/fileLine -> access_run / access_sql to verify.
If objects were changed in the Access GUI, run access_export first (it refuses to overwrite unimported edits).
tables/*.json and relationships.json are read-only snapshots; change schema with access_sql write=true (DDL) and then access_export kinds=["tables"].
VBA modules are stored in Shift-JIS inside Access: avoid characters that Shift-JIS cannot represent (e.g. emoji).
Report layout work: access_report_layout (positions in twips, 1 cm = 567) -> edit Left/Top/Width/Height/Caption/FontSize... in reports/<name>.txt -> access_import -> access_report_preview to look at the rendered pages and fix until "issues" is empty and the images look right. Paper size / orientation / margins: access_report_pagesetup (they are binary in the text file; do not edit PrtDevMode/PrtMip lines).
Report visual regression: after changing reports or the code/queries behind them, run access_report_snapshot; if pages changed, look at the diff images and report to the user. Accept (update=true) only intended changes.
Release: access_build makes the .accde. Never publish (deploy) unless the user explicitly asks; publishing is done from VS Code or the CLI.
Tests: write Public Sub Test_*() in a standard module (e.g. modules/Test_Tax.bas) using Assert.AreEqual expected, actual -> access_import -> access_test. To test forms like a user (typing, clicking buttons, AfterUpdate/Click events fire), use the UI module: UI.OpenForm "frm", UI.TypeText "txt", "value", UI.Click "cmd", UI.Value("txt"), and UI.CloseForm in TestTeardown. Run access_lint before finishing.
'@

# ------------------------------------------------------------
# ツール呼び出し -> Invoke-AccessVsc
# ------------------------------------------------------------
function Invoke-Tool([string]$Name, $A) {
    if (-not $A) { $A = @{} }
    $P = @{ db = $(if ($A.database) { $A.database } else { $DbPath }); src = $Src }
    switch ($Name) {
        'access_status' { $cmd = 'status' }
        'access_list' { $cmd = 'list' }
        'access_export' { $cmd = 'export'; $P.force = [bool]$A.force; if ($A.kinds) { $P.kinds = @($A.kinds) } }
        'access_import' {
            $cmd = 'import'
            if ($A.files) { $P.files = @($A.files) }
            $P.all = [bool]$A.all; $P.deleteMissing = [bool]$A.deleteMissing
            $P.compile = $(if ($null -ne $A.compile) { [bool]$A.compile } else { $true })
        }
        'access_compile' { $cmd = 'compile' }
        'access_run' { $cmd = 'run'; $P.function = $A.function; $P.args = $(if ($null -ne $A.args) { @($A.args) } else { @() }) }
        'access_eval' { $cmd = 'eval'; $P.expression = $A.expression }
        'access_sql' { $cmd = 'sql'; $P.sql = $A.sql; $P.limit = $(if ($A.limit) { [int]$A.limit } else { 100 }); $P.write = [bool]$A.write }
        'access_report_layout' { $cmd = 'layout'; $P.report = $A.report; $P.objectType = $A.objectType }
        'access_report_preview' {
            $cmd = 'preview'; $P.report = $A.report; $P.objectType = $A.objectType; $P.where = $A.where
            $P.pages = $(if ($A.pages) { [int]$A.pages } else { 2 })
        }
        'access_report_pagesetup' {
            $cmd = 'pagesetup'; $P.report = $A.report
            foreach ($k in 'paperSize', 'orientation', 'marginTop', 'marginBottom', 'marginLeft', 'marginRight', 'widthCm') { if ($null -ne $A[$k]) { $P[$k] = $A[$k] } }
        }
        'access_test' { $cmd = 'test'; if ($A.filter) { $P.filter = @($A.filter) } }
        'access_lint' { $cmd = 'lint' }
        'access_format' { $cmd = 'format'; if ($A.files) { $P.files = @($A.files) }; $P.check = [bool]$A.check }
        'access_report_snapshot' {
            $cmd = 'snapshot'
            if ($A.reports) { $P.reports = @($A.reports) }
            $P.update = [bool]$A.update
            if ($null -ne $A.threshold) { $P.threshold = [double]$A.threshold }
        }
        'access_build' { $cmd = 'build'; $P.version = $A.version }
        'access_setup_ci' { $cmd = 'setup-ci'; $P.root = $A.root; $P.gitInit = [bool]$A.gitInit; $P.force = [bool]$A.force }
        default { throw "Unknown tool: $Name" }
    }
    return Invoke-AccessVsc $cmd $P
}

Write-Log "started (default database: $(if ($DbPath) { $DbPath } else { '(none)' }))"

while ($true) {
    $line = $stdin.ReadLine()
    if ($null -eq $line) { break }
    if (-not $line.Trim()) { continue }
    try { $msg = $line | ConvertFrom-Json }
    catch { Send-Error $null -32700 'Parse error'; continue }

    $hasId = $msg.PSObject.Properties.Name -contains 'id'
    $id = $msg.id
    try {
        switch ($msg.method) {
            'initialize' {
                $pv = '2025-06-18'
                if ($msg.params -and $msg.params.protocolVersion) { $pv = [string]$msg.params.protocolVersion }
                Send-Result $id ([ordered]@{
                        protocolVersion = $pv
                        capabilities    = [ordered]@{ tools = [ordered]@{ listChanged = $false } }
                        serverInfo      = [ordered]@{ name = 'access-vsc'; title = 'Access VSC'; version = $script:Version }
                        instructions    = $instructions
                    })
            }
            'ping' { Send-Result $id @{} }
            'tools/list' { Send-Result $id ([ordered]@{ tools = $tools }) }
            'tools/call' {
                $name = [string]$msg.params.name
                $toolArgs = ConvertTo-HashtableDeep $msg.params.arguments
                Write-Log "call $name"
                $res = Invoke-Tool $name $toolArgs
                $content = New-Object System.Collections.ArrayList
                [void]$content.Add([ordered]@{ type = 'text'; text = (ConvertTo-PrettyJson $res) })
                # プレビュー画像はそのまま AI に見せる
                $imgs = @()
                if ($res.ok -and $res.result -and $res.result.images) { $imgs = @($res.result.images) }
                if ($name -eq 'access_report_snapshot' -and $res.result -and $res.result.reports) {
                    # 変化したページの差分画像（違う所が赤）と、今の見た目
                    foreach ($r in @($res.result.reports)) {
                        foreach ($pg in @($r.pages | Where-Object { $_.status -in @('changed', 'added', 'new') })) {
                            if ($pg.diff) { $imgs += $pg.diff } elseif ($pg.current) { $imgs += $pg.current }
                        }
                    }
                    $imgs = @($imgs | Select-Object -First 4)
                }
                foreach ($img in $imgs) {
                    if ($img -and (Test-Path -LiteralPath $img)) {
                        [void]$content.Add([ordered]@{ type = 'image'; mimeType = 'image/png'; data = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($img)) })
                    }
                }
                Send-Result $id ([ordered]@{ content = $content.ToArray(); isError = (-not $res.ok) })
            }
            default {
                if ($hasId -and $msg.method) { Send-Error $id -32601 "Method not found: $($msg.method)" }
                # 通知（id なし）は返事をしない
            }
        }
    }
    catch {
        Write-Log "error: $($_.Exception.Message)"
        if ($hasId) {
            if ($msg.method -eq 'tools/call') {
                Send-Result $id ([ordered]@{ content = @([ordered]@{ type = 'text'; text = $_.Exception.Message }); isError = $true })
            }
            else { Send-Error $id -32603 $_.Exception.Message }
        }
    }
}
