<#
  build-vsix.ps1
  npm / vsce を使わずに .vsix（VS Code 拡張機能のパッケージ）を作る。
  使い方:
    powershell -NoProfile -ExecutionPolicy Bypass -File .\build-vsix.ps1            # 作るだけ
    powershell -NoProfile -ExecutionPolicy Bypass -File .\build-vsix.ps1 -Install   # 作って VS Code にインストール
#>
param([switch]$Install)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$root = $PSScriptRoot
$pkg = Get-Content -Raw -Encoding UTF8 (Join-Path $root 'package.json') | ConvertFrom-Json
$out = Join-Path $root ("{0}-{1}.vsix" -f $pkg.name, $pkg.version)

# パッケージに入れるファイル
$include = @('package.json', 'extension.js', 'vba-language.js', 'vba-format.js', 'language-configuration.json', 'README.md', 'LICENSE') +
    @(Get-ChildItem (Join-Path $root 'snippets') -File | ForEach-Object { "snippets/$($_.Name)" }) +
    @(Get-ChildItem (Join-Path $root 'syntaxes') -File | ForEach-Object { "syntaxes/$($_.Name)" }) +
    @(Get-ChildItem (Join-Path $root 'scripts') -File -Recurse | ForEach-Object { $_.FullName.Substring($root.Length + 1).Replace('\', '/') }) +
    @(Get-ChildItem (Join-Path $root 'media') -File -Recurse | ForEach-Object { $_.FullName.Substring($root.Length + 1).Replace('\', '/') }) +
    @(Get-ChildItem (Join-Path $root 'docs') -File -Recurse | Where-Object { $_.Extension -in '.html', '.png', '.svg', '.css', '.js' } | ForEach-Object { $_.FullName.Substring($root.Length + 1).Replace('\', '/') })

function Esc([string]$s) { return [System.Security.SecurityElement]::Escape($s) }

$manifest = @"
<?xml version="1.0" encoding="utf-8"?>
<PackageManifest Version="2.0.0" xmlns="http://schemas.microsoft.com/developer/vsx-schema/2011" xmlns:d="http://schemas.microsoft.com/developer/vsx-schema-design/2011">
  <Metadata>
    <Identity Language="en-US" Id="$(Esc $pkg.name)" Version="$(Esc $pkg.version)" Publisher="$(Esc $pkg.publisher)" />
    <DisplayName>$(Esc $pkg.displayName)</DisplayName>
    <Description xml:space="preserve">$(Esc $pkg.description)</Description>
    <Tags>$(Esc (($pkg.keywords) -join ','))</Tags>
    <Categories>$(Esc (($pkg.categories) -join ','))</Categories>
    <GalleryFlags>Public</GalleryFlags>
    <Properties>
      <Property Id="Microsoft.VisualStudio.Code.Engine" Value="$(Esc $pkg.engines.vscode)" />
      <Property Id="Microsoft.VisualStudio.Code.ExtensionDependencies" Value="" />
      <Property Id="Microsoft.VisualStudio.Code.ExtensionPack" Value="" />
      <Property Id="Microsoft.VisualStudio.Code.ExtensionKind" Value="workspace" />
      <Property Id="Microsoft.VisualStudio.Code.LocalizedLanguages" Value="" />
    </Properties>
    <Icon>extension/media/icon.png</Icon>
  </Metadata>
  <Installation>
    <InstallationTarget Id="Microsoft.VisualStudio.Code" />
  </Installation>
  <Dependencies />
  <Assets>
    <Asset Type="Microsoft.VisualStudio.Code.Manifest" Path="extension/package.json" Addressable="true" />
    <Asset Type="Microsoft.VisualStudio.Services.Content.Details" Path="extension/README.md" Addressable="true" />
    <Asset Type="Microsoft.VisualStudio.Services.Icons.Default" Path="extension/media/icon.png" Addressable="true" />
  </Assets>
</PackageManifest>
"@

$contentTypes = @'
<?xml version="1.0" encoding="utf-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension=".json" ContentType="application/json" />
  <Default Extension=".vsixmanifest" ContentType="text/xml" />
  <Default Extension=".js" ContentType="application/javascript" />
  <Default Extension=".md" ContentType="text/markdown" />
  <Default Extension=".ps1" ContentType="text/plain" />
  <Default Extension=".cs" ContentType="text/plain" />
  <Default Extension=".svg" ContentType="image/svg+xml" />
  <Default Extension=".bas" ContentType="text/plain" />
  <Default Extension=".yml" ContentType="text/plain" />
  <Default Extension=".sh" ContentType="text/plain" />
  <Default Extension=".html" ContentType="text/html" />
  <Default Extension=".css" ContentType="text/css" />
  <Default Extension=".png" ContentType="image/png" />
</Types>
'@

if (Test-Path $out) { Remove-Item $out -Force }
$utf8 = New-Object System.Text.UTF8Encoding($false)
$zip = [System.IO.Compression.ZipFile]::Open($out, 'Create')
try {
    function Add-Text($name, $text) {
        $e = $zip.CreateEntry($name)
        $w = New-Object System.IO.StreamWriter($e.Open(), $utf8)
        try { $w.Write($text) } finally { $w.Dispose() }
    }
    Add-Text 'extension.vsixmanifest' $manifest
    Add-Text '[Content_Types].xml' $contentTypes
    foreach ($rel in $include) {
        $full = Join-Path $root ($rel.Replace('/', '\'))
        $entry = if ($rel -eq 'LICENSE') { 'LICENSE.txt' } else { $rel }   # vsce と同じく拡張子を付ける
        [void][System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $full, "extension/$entry")
    }
}
finally { $zip.Dispose() }
Write-Output "built: $out ($([math]::Round((Get-Item $out).Length / 1KB)) KB, $($include.Count) files)"

if ($Install) {
    $code = (Get-Command code -ErrorAction SilentlyContinue).Source
    if (-not $code) { throw 'code コマンドが見つかりません（VS Code の PATH 設定を確認してください）' }
    & $code --install-extension $out --force
}
