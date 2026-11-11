<#
  Run-ExtensionTest.ps1 — 本物の VS Code の中で Access VSC の編集支援などを確かめる結合テスト。
  普段の VS Code とは別の設定・別の拡張機能フォルダで起動し、終わったら閉じる。
  使い方: powershell -NoProfile -ExecutionPolicy Bypass -File .\test\integration\Run-ExtensionTest.ps1 -Work <作業フォルダ>
#>
param([Parameter(Mandatory = $true)][string]$Work, [int]$TimeoutMinutes = 6)
$ErrorActionPreference = 'Stop'
$repo = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$Work = [System.IO.Path]::GetFullPath($Work)
$ud = Join-Path $Work 'vsc-itest-ud'; $ext = Join-Path $Work 'vsc-itest-ext'; $ws = Join-Path $Work 'shop'; $out = Join-Path $Work 'itest-out'
foreach ($d in $ud, $ext, $ws, $out) { if (Test-Path $d) { Remove-Item $d -Recurse -Force } }
New-Item -ItemType Directory -Path $ud, $ext, $ws, $out, (Join-Path $ud 'User') -Force | Out-Null
$code = (Get-Command code -ErrorAction Stop).Source
function Invoke-Code { $prev = $ErrorActionPreference; $ErrorActionPreference = 'Continue'; try { & $code @args 2>&1 | ForEach-Object { "$_" } | Where-Object { $_ -notmatch 'DeprecationWarning|trace-deprecation' } } finally { $ErrorActionPreference = $prev } }
$u8 = New-Object System.Text.UTF8Encoding($false)

# 拡張機能とテスト用ドライバー
& powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $repo 'build-vsix.ps1') | Write-Output
$vsix = Get-ChildItem $repo -Filter 'access-vscode-*.vsix' | Sort-Object LastWriteTime -Descending | Select-Object -First 1
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
$driverVsix = Join-Path $Work 'itest-driver.vsix'
if (Test-Path $driverVsix) { Remove-Item $driverVsix }
$zip = [System.IO.Compression.ZipFile]::Open($driverVsix, 'Create')
try {
    $add = { param($name, $text) $e = $zip.CreateEntry($name); $w = New-Object System.IO.StreamWriter($e.Open(), $u8); $w.Write($text); $w.Dispose() }
    & $add 'extension.vsixmanifest' '<?xml version="1.0" encoding="utf-8"?><PackageManifest Version="2.0.0" xmlns="http://schemas.microsoft.com/developer/vsx-schema/2011"><Metadata><Identity Language="en-US" Id="accessvsc-test-driver" Version="0.0.1" Publisher="local" /><DisplayName>test driver</DisplayName><Description xml:space="preserve">test only</Description><Properties><Property Id="Microsoft.VisualStudio.Code.Engine" Value="^1.101.0" /></Properties></Metadata><Installation><InstallationTarget Id="Microsoft.VisualStudio.Code" /></Installation><Dependencies /><Assets><Asset Type="Microsoft.VisualStudio.Code.Manifest" Path="extension/package.json" Addressable="true" /></Assets></PackageManifest>'
    & $add '[Content_Types].xml' '<?xml version="1.0" encoding="utf-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension=".json" ContentType="application/json" /><Default Extension=".vsixmanifest" ContentType="text/xml" /><Default Extension=".js" ContentType="application/javascript" /></Types>'
    foreach ($f in 'package.json', 'extension.js') { [void][System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, (Join-Path $PSScriptRoot "driver\$f"), "extension/$f") }
}
finally { $zip.Dispose() }
[System.IO.File]::WriteAllText((Join-Path $ud 'User\settings.json'), '{ "workbench.startupEditor": "none", "security.workspace.trust.enabled": false, "update.mode": "none", "extensions.autoUpdate": false, "chat.disableAIFeatures": true, "window.restoreWindows": "none" }', $u8)
Invoke-Code --user-data-dir $ud --extensions-dir $ext --install-extension $vsix.FullName --force
Invoke-Code --user-data-dir $ud --extensions-dir $ext --install-extension $driverVsix --force

# テスト用のサンプル（取り出してから、テスト用のファイルを足す）
$db = Join-Path $ws 'shop.accdb'
& powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $repo 'scripts\New-SampleDatabase.ps1') -Path $db | Write-Output
& powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $repo 'scripts\access.ps1') export -DbPath $db -Compress | Out-Null
$mods = Join-Path $ws 'shop.accdb.src\modules'
[System.IO.File]::WriteAllText((Join-Path $mods 'modTax.bas'), @"
Option Compare Database
Option Explicit

Public Const TAX_RATE As Double = 0.1

' 税込金額（1 円未満切り捨て）
Public Function TaxIncluded(ByVal price As Currency) As Currency
    Dim subtotal As Currency
    subtotal = Nz(price, 0)
    TaxIncluded = Int(subtotal * (1 + TAX_RATE))
End Function

' 合計に税を足す
Public Function 合計金額(ByVal amount As Currency) As Currency
    合計金額 = TaxIncluded(amount)
End Function

Private Function PrivateHelper() As Long
    PrivateHelper = 1
End Function
"@, $u8)
[System.IO.File]::WriteAllText((Join-Path $mods 'Test_Tax.bas'), @"
Option Compare Database
Option Explicit

Public Sub Test_Tax()
    ' TaxIncluded はコメントの中（参照に含めない）
    Debug.Print "TaxIncluded は文字列の中"
    Assert.AreEqual 110, TaxIncluded(100)
    Assert.AreEqual 110, 合計金額(100)
    Assert.AreEqual 110, modTax.TaxIncluded(100)
End Sub
"@, $u8)
[System.IO.File]::WriteAllText((Join-Path $ws '.itest.json'), (ConvertTo-Json @{ out = $out }), $u8)

# 起動（VS Code の中から実行しても動くよう、VSCODE_* などを外す）
$codeExe = Join-Path (Split-Path -Parent (Split-Path -Parent $code)) 'Code.exe'
$psi = New-Object System.Diagnostics.ProcessStartInfo
$psi.FileName = $codeExe
$psi.Arguments = ('--user-data-dir "{0}" --extensions-dir "{1}" --new-window "{2}"' -f $ud, $ext, $ws)
$psi.UseShellExecute = $false
foreach ($k in @($psi.EnvironmentVariables.Keys)) { if ($k -match '^(VSCODE_|ELECTRON_)') { $psi.EnvironmentVariables.Remove($k) } }
[void][System.Diagnostics.Process]::Start($psi)
$deadline = (Get-Date).AddMinutes($TimeoutMinutes)
while (-not (Test-Path (Join-Path $out 'done.txt')) -and (Get-Date) -lt $deadline) { Start-Sleep -Seconds 2 }
Get-CimInstance Win32_Process -Filter "Name='Code.exe'" | Where-Object { $_.CommandLine -match 'vsc-itest-ud' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }

if (Test-Path (Join-Path $out 'error.txt')) { Write-Output ('ERROR: ' + (Get-Content (Join-Path $out 'error.txt') -Raw)) }
$resultsFile = Join-Path $out 'results.json'
if (-not (Test-Path $resultsFile)) { Write-Output 'no results (timeout?)'; exit 1 }
$results = [System.IO.File]::ReadAllText($resultsFile) | ConvertFrom-Json
$fail = 0
foreach ($r in $results) {
    if ($r.pass) { Write-Output "PASS $($r.name)" } else { $fail++; Write-Output "FAIL $($r.name) -> $(ConvertTo-Json $r.detail -Compress -Depth 4)" }
}
if ($fail) { Write-Output "$fail failure(s)"; exit 1 }
Write-Output 'all passed'
