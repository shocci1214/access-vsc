<#
  Access VSC ランチャー（配布先の共有フォルダに置かれ、利用者が「Start-○○.cmd」から起動する）
    1. 共有フォルダの version.json を読む
    2. 利用者の PC（%LOCALAPPDATA%\AccessApps\<アプリ名>）の版が古ければ、最新版をコピーする
    3. ローカルのコピーを Access で開く
  ネットワーク越しに直接開かないので速く、ファイル破損や「使用中で更新できない」問題も起きにくい。

  ※ .accde を開いたときにセキュリティの警告が出る場合は、管理者に
     %LOCALAPPDATA%\AccessApps を Access の「信頼できる場所」に登録してもらってください
     （グループポリシーまたは [ファイル] → [オプション] → [トラストセンター]）。
#>
param([switch]$NoStart)   # -NoStart: 更新だけして起動しない（動作確認用）
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms

function Show-Message([string]$Text, [string]$Icon = 'Information') {
    [void][System.Windows.Forms.MessageBox]::Show($Text, 'アプリの起動', 'OK', $Icon)
}

try {
    $share = $PSScriptRoot
    $info = [System.IO.File]::ReadAllText((Join-Path $share 'version.json'), [System.Text.Encoding]::UTF8) | ConvertFrom-Json
    # 安全確認：version.json が書き換えられても、共有フォルダの外のファイルや .exe などは開かない
    if ([string]$info.appName -notmatch '^[\w\-. \p{L}]+$' -or [string]$info.appName -match '\.\.') { throw "version.json の appName が不正です: $($info.appName)" }
    if ([string]$info.file -notmatch '(?i)^[\w\-. /\\\p{L}]+\.(accde|accdb|mde|mdb)$' -or [string]$info.file -match '\.\.') { throw "version.json の file が不正です（.accde / .accdb 以外、または共有フォルダの外）: $($info.file)" }
    $source = [System.IO.Path]::GetFullPath((Join-Path $share $info.file))
    if (-not $source.StartsWith([System.IO.Path]::GetFullPath($share).TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) { throw '配布ファイルが共有フォルダの外を指しています。' }
    if (-not $info.sha256) { throw 'version.json にハッシュ（sha256）がありません。' }
    $localDir = Join-Path $env:LOCALAPPDATA ('AccessApps\' + $info.appName)
    $local = Join-Path $localDir ([System.IO.Path]::GetFileName($info.file))
    $localInfo = Join-Path $localDir 'version.json'
    $current = ''
    if (Test-Path -LiteralPath $localInfo) { $current = ([System.IO.File]::ReadAllText($localInfo, [System.Text.Encoding]::UTF8) | ConvertFrom-Json).version }

    if ($current -ne $info.version -or -not (Test-Path -LiteralPath $local)) {
        $lock = [System.IO.Path]::ChangeExtension($local, '.laccdb')
        if ((Test-Path -LiteralPath $local) -and (Test-Path -LiteralPath $lock)) {
            # 古い版を使用中。いったん閉じてもらう
            try { $fs = [System.IO.File]::Open($local, 'Open', 'ReadWrite', 'None'); $fs.Close() }
            catch {
                Show-Message ("新しいバージョン {0} があります。`n開いているアプリを閉じてから、もう一度起動してください。" -f $info.version) 'Warning'
                exit 1
            }
        }
        if (-not (Test-Path -LiteralPath $localDir)) { [void](New-Item -ItemType Directory -Path $localDir -Force) }
        $tmp = "$local.download"
        Copy-Item -LiteralPath $source -Destination $tmp -Force
        $sha = [System.Security.Cryptography.SHA256]::Create()
        try { $hash = ([BitConverter]::ToString($sha.ComputeHash([System.IO.File]::ReadAllBytes($tmp))) -replace '-', '').ToLowerInvariant() } finally { $sha.Dispose() }
        if ($hash -ne [string]$info.sha256) {
            Remove-Item -LiteralPath $tmp -Force
            Show-Message 'ダウンロードしたファイルが壊れています。時間をおいてもう一度起動してください。' 'Error'
            exit 1
        }
        Move-Item -LiteralPath $tmp -Destination $local -Force
        Copy-Item -LiteralPath (Join-Path $share 'version.json') -Destination $localInfo -Force
    }
    if ($NoStart) { Write-Output "ready: $local (version $($info.version))"; exit 0 }
    Start-Process -FilePath $local
}
catch {
    Show-Message ("起動できませんでした。`n" + $_.Exception.Message) 'Error'
    exit 1
}
