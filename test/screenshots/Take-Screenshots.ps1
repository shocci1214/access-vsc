<#
  Take-Screenshots.ps1 — マニュアル（docs/manual.html）用のスクリーンショットを撮り直す。
    1. 拡張機能（.vsix）と撮影用ドライバーを作る
    2. 普段の VS Code とは別の設定・別の拡張機能フォルダで VS Code を起動（日本語表示）
    3. ドライバーが操作し、場面ごとに docs/images/*.png を保存
    4. 撮影用の VS Code を閉じる
  使い方: powershell -NoProfile -ExecutionPolicy Bypass -File .\test\screenshots\Take-Screenshots.ps1 -Work <作業フォルダ>
  ※ 撮影中は VS Code のウィンドウが表示されます。触らずにお待ちください。
#>
param(
    [Parameter(Mandatory = $true)][string]$Work,
    [int]$TimeoutMinutes = 15
)
$ErrorActionPreference = 'Stop'
$repo = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$out = Join-Path $repo 'docs\images'
$Work = [System.IO.Path]::GetFullPath($Work)
$ud = Join-Path $Work 'vsc-shot-ud'
$ext = Join-Path $Work 'vsc-shot-ext'
$ws = Join-Path $Work 'shop'
$code = (Get-Command code -ErrorAction Stop).Source
function Invoke-Code {
    # code.cmd は標準エラーに警告を出す。Windows PowerShell 5.1 ではそれがエラー扱いになるので、ここだけ止めない
    $prev = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try { & $code @args 2>&1 | ForEach-Object { "$_" } | Where-Object { $_ -notmatch 'DeprecationWarning|trace-deprecation' } }
    finally { $ErrorActionPreference = $prev }
}

foreach ($d in $ud, $ext, $ws) { if (Test-Path $d) { Remove-Item $d -Recurse -Force } }
New-Item -ItemType Directory -Path $ud, $ext, $ws, (Join-Path $ud 'User') -Force | Out-Null
if (Test-Path $out) { Get-ChildItem $out -Filter '*.png' | Remove-Item -Force } else { New-Item -ItemType Directory -Path $out | Out-Null }
Remove-Item (Join-Path $out 'log.txt'), (Join-Path $out 'done.txt') -ErrorAction SilentlyContinue

# 1) パッケージ
& powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $repo 'build-vsix.ps1') | Write-Output
$vsix = Get-ChildItem $repo -Filter 'access-vscode-*.vsix' | Sort-Object LastWriteTime -Descending | Select-Object -First 1
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
$driverVsix = Join-Path $Work 'shot-driver.vsix'
if (Test-Path $driverVsix) { Remove-Item $driverVsix }
$zip = [System.IO.Compression.ZipFile]::Open($driverVsix, 'Create')
try {
    $u8 = New-Object System.Text.UTF8Encoding($false)
    $add = { param($name, $text) $e = $zip.CreateEntry($name); $w = New-Object System.IO.StreamWriter($e.Open(), $u8); $w.Write($text); $w.Dispose() }
    & $add 'extension.vsixmanifest' @'
<?xml version="1.0" encoding="utf-8"?>
<PackageManifest Version="2.0.0" xmlns="http://schemas.microsoft.com/developer/vsx-schema/2011">
  <Metadata><Identity Language="en-US" Id="accessvsc-shot-driver" Version="0.0.1" Publisher="local" /><DisplayName>shot driver</DisplayName><Description xml:space="preserve">test only</Description>
  <Properties><Property Id="Microsoft.VisualStudio.Code.Engine" Value="^1.101.0" /></Properties></Metadata>
  <Installation><InstallationTarget Id="Microsoft.VisualStudio.Code" /></Installation><Dependencies />
  <Assets><Asset Type="Microsoft.VisualStudio.Code.Manifest" Path="extension/package.json" Addressable="true" /></Assets>
</PackageManifest>
'@
    & $add '[Content_Types].xml' '<?xml version="1.0" encoding="utf-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension=".json" ContentType="application/json" /><Default Extension=".vsixmanifest" ContentType="text/xml" /><Default Extension=".js" ContentType="application/javascript" /></Types>'
    foreach ($f in 'package.json', 'extension.js') { [void][System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, (Join-Path $PSScriptRoot "driver\$f"), "extension/$f") }
}
finally { $zip.Dispose() }

# 2) 撮影用 VS Code の準備
@'
{
  "workbench.startupEditor": "none",
  "workbench.tips.enabled": false,
  "workbench.welcomePage.walkthroughs.openOnInstall": false,
  "workbench.secondarySideBar.defaultVisibility": "hidden",
  "update.mode": "none",
  "extensions.ignoreRecommendations": true,
  "extensions.autoUpdate": false,
  "security.workspace.trust.enabled": false,
  "telemetry.telemetryLevel": "off",
  "chat.disableAIFeatures": true,
  "editor.minimap.enabled": false,
  "git.openRepositoryInParentFolders": "never",
  "files.exclude": { "**/.shot.json": true, "**/.accessvsc": true, "**/.git": true },
  "window.restoreWindows": "none",
  "window.dialogStyle": "custom"
}
'@ | ForEach-Object { [System.IO.File]::WriteAllText((Join-Path $ud 'User\settings.json'), $_, (New-Object System.Text.UTF8Encoding($false))) }
Invoke-Code --user-data-dir $ud --extensions-dir $ext --install-extension $vsix.FullName --force
Invoke-Code --user-data-dir $ud --extensions-dir $ext --install-extension $driverVsix --force
Invoke-Code --user-data-dir $ud --extensions-dir $ext --install-extension MS-CEINTL.vscode-language-pack-ja

# 3) デモ用 DB
& powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $repo 'scripts\New-SampleDatabase.ps1') -Path (Join-Path $ws 'shop.accdb') | Write-Output
[ordered]@{
    out     = $out
    capture = (Join-Path $PSScriptRoot 'capture.ps1')
    marker  = 'vsc-shot-ud'
    cli     = (Join-Path $repo 'scripts\access.ps1')
} | ConvertTo-Json | ForEach-Object { [System.IO.File]::WriteAllText((Join-Path $ws '.shot.json'), $_, (New-Object System.Text.UTF8Encoding($false))) }

# 4) 起動して、撮影が終わるのを待つ
# VS Code の中（Claude Code など）から実行すると VSCODE_* / ELECTRON_RUN_AS_NODE が引き継がれ、
# 新しい VS Code が正しく起動しない。これらを取り除いて Code.exe を直接起動する
$codeExe = Join-Path (Split-Path -Parent (Split-Path -Parent $code)) 'Code.exe'
$psi = New-Object System.Diagnostics.ProcessStartInfo
$psi.FileName = $codeExe
$psi.Arguments = ('--user-data-dir "{0}" --extensions-dir "{1}" --locale=ja --new-window "{2}"' -f $ud, $ext, $ws)
$psi.UseShellExecute = $false
foreach ($k in @($psi.EnvironmentVariables.Keys)) { if ($k -match '^(VSCODE_|ELECTRON_)') { $psi.EnvironmentVariables.Remove($k) } }
[void][System.Diagnostics.Process]::Start($psi)
$deadline = (Get-Date).AddMinutes($TimeoutMinutes)
while (-not (Test-Path (Join-Path $out 'done.txt')) -and (Get-Date) -lt $deadline) { Start-Sleep -Seconds 3 }
Start-Sleep -Seconds 2

# 5) 撮影用の VS Code だけを閉じる
Get-CimInstance Win32_Process -Filter "Name='Code.exe'" | Where-Object { $_.CommandLine -match 'vsc-shot-ud' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
if (Test-Path (Join-Path $out 'log.txt')) { Get-Content (Join-Path $out 'log.txt') -Encoding UTF8 }
foreach ($n in 'log.txt', 'done.txt', 'error.txt') { $p = Join-Path $out $n; if (Test-Path $p) { Move-Item $p (Join-Path $Work "shot-$n") -Force } }
Get-ChildItem $out -Filter '*.png' | Select-Object Name, Length
