<#
  New-Mascot.ps1 — マスコット「つなぎ」のドット絵を、設計図（文字の格子）から SVG と PNG に書き出す。
    つなぎ：VS Code（左・青）と Access のデータベース（右・えんじ）を「つなぐ」子。頭の 2 本の角はプラグ。
  使い方: powershell -NoProfile -ExecutionPolicy Bypass -File media\mascot\New-Mascot.ps1
  出力: media/mascot/tsunagi-*.svg（ドキュメント・README・VS Code の画面用）、media/icon.png（拡張機能のアイコン 128x128）
#>
$ErrorActionPreference = 'Stop'
$here = $PSScriptRoot
$repo = (Resolve-Path (Join-Path $here '..\..')).Path

# 大文字と小文字を区別する（PowerShell の @{} は区別しないため）
$palette = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([System.StringComparer]::Ordinal)
[void]$palette.Add('K', '#2E2A3B')   # 輪郭・目
[void]$palette.Add('G', '#D5DAE1')   # プラグの金具
[void]$palette.Add('g', '#9AA3AE')   # 金具の影
[void]$palette.Add('B', '#4FA3F0')   # 左半分（青）
[void]$palette.Add('b', '#A9D6FB')   # 青のハイライト
[void]$palette.Add('d', '#2F7FD0')   # 青の影
[void]$palette.Add('R', '#C2475E')   # 右半分（えんじ）
[void]$palette.Add('r', '#EE93A3')   # えんじのハイライト
[void]$palette.Add('e', '#97304A')   # えんじの影
[void]$palette.Add('W', '#FFFFFF')   # 目の光
[void]$palette.Add('C', '#FF8FA8')   # ほっぺ
[void]$palette.Add('S', '#8FD3FF')   # 汗
[void]$palette.Add('Y', '#FFD54A')   # きらきら

# 20 x 18 の格子。. は透明
$base = @(
    '......KK....KK......',
    '......KG....GK......',
    '......Kg....gK......',
    '....KKKKKKKKKKKK....',
    '...KbbBBBBRRRRrrK...',
    '..KbBBBBBBRRRRRRrK..',
    '..KBBBBBBBRRRRRRRK..',
    '.KBBBBBBBBRRRRRRRRK.',
    '.KBBEEBBBBRRRRFFRRK.',
    '.KBBEEBBBBRRRRFFRRK.',
    '.KBBBBBBBBRRRRRRRRK.',
    '.KBCCBBMMMMMMRRCCRK.',
    '.KdBBBBBBBRRRRRRReK.',
    '..KdBBBBBBRRRRRReK..',
    '..KddBBBBBRRRRReeK..',
    '...KKddBBBRRReeKK...',
    '.....KKKKKKKKKK.....',
    '.....KK......KK.....'
)

# 表情ごとに、目（E/F）と口（M）の描き方を変える
function Get-Face([string]$Mood) {
    switch ($Mood) {
        'happy' {
            # にっこり目（^ ^）と大きな口
            return @{ eyes = @('..K.', '.K.K'); mouth = 'w'; extra = 'sparkle' }
        }
        'worried' { return @{ eyes = @('KK..', 'KW..'); mouth = 'flat'; extra = 'sweat' } }
        default { return @{ eyes = @('KK..', 'KW..'); mouth = 'smile'; extra = $null } }
    }
}

function New-Pixels([string]$Mood) {
    $face = Get-Face $Mood
    $h = $base.Count; $w = $base[0].Length
    $px = New-Object 'string[,]' $h, $w
    for ($y = 0; $y -lt $h; $y++) { for ($x = 0; $x -lt $w; $x++) { $px[$y, $x] = [string]$base[$y][$x] } }
    $set = { param([int]$y, [int]$x, [string]$c) $px[$y, $x] = $c }
    # 目（左は 4 列目、右は 14 列目。7〜9 行目の縦長）。背景色で塗ってから描く
    foreach ($eye in @(@{ x = 4; bg = 'B' }, @{ x = 14; bg = 'R' })) {
        $x0 = $eye.x
        foreach ($dy in 0, 1, 2) { foreach ($dx in 0, 1) { & $set (7 + $dy) ($x0 + $dx) $eye.bg } }
        if ($Mood -eq 'happy') {
            # にっこり目（^）
            & $set 8 $x0 'K'; & $set 8 ($x0 + 1) 'K'; & $set 9 ($x0 - 1) 'K'; & $set 9 ($x0 + 2) 'K'
        }
        else {
            # 縦長の目。右上に光
            & $set 7 $x0 'K'; & $set 7 ($x0 + 1) 'W'
            & $set 8 $x0 'K'; & $set 8 ($x0 + 1) 'K'
            & $set 9 $x0 'K'; & $set 9 ($x0 + 1) 'K'
            if ($Mood -eq 'worried') { & $set 6 ($x0 + $(if ($x0 -lt 10) { 1 } else { 0 })) 'K' }   # 困り眉
        }
    }
    # 口（11 行目の 7〜12 列）。背景色で塗ってから描く
    foreach ($x in 7..12) { & $set 11 $x $(if ($x -lt 10) { 'B' } else { 'R' }) }
    switch ($face.mouth) {
        'w' { foreach ($p in @(@(11, 8), @(12, 9), @(11, 10), @(12, 11), @(11, 12))) { & $set $p[0] $p[1] 'K' } }   # 大きな ω
        'flat' { & $set 11 9 'K'; & $set 11 10 'K' }
        default { & $set 11 9 'K'; & $set 12 10 'K'; & $set 11 11 'K' }   # 小さな口
    }
    if ($face.extra -eq 'sweat') { & $set 5 18 'S'; & $set 6 18 'S'; & $set 6 19 'S' }
    if ($face.extra -eq 'sparkle') { foreach ($p in @(@(0, 1), @(1, 0), @(1, 1), @(1, 2), @(2, 1), @(2, 18), @(3, 17), @(3, 18), @(3, 19), @(4, 18))) { & $set $p[0] $p[1] 'Y' } }
    return , $px
}

function ConvertTo-Svg($Px, [string]$Title) {
    $h = $Px.GetLength(0); $w = $Px.GetLength(1)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<svg xmlns=""http://www.w3.org/2000/svg"" viewBox=""0 0 $w $h"" width=""$($w * 8)"" height=""$($h * 8)"" shape-rendering=""crispEdges"" role=""img"" aria-label=""$Title""><title>$Title</title>")
    for ($y = 0; $y -lt $h; $y++) {
        $x = 0
        while ($x -lt $w) {
            $c = $Px[$y, $x]
            if ($c -eq '.' -or -not $palette.ContainsKey($c)) { $x++; continue }
            $run = 1
            while ($x + $run -lt $w -and $Px[$y, ($x + $run)] -eq $c) { $run++ }
            [void]$sb.Append("<rect x=""$x"" y=""$y"" width=""$run"" height=""1"" fill=""$($palette[$c])""/>")
            $x += $run
        }
    }
    [void]$sb.Append('</svg>')
    return $sb.ToString()
}

function Save-Png($Px, [string]$Path, [int]$Size) {
    Add-Type -AssemblyName System.Drawing
    $h = $Px.GetLength(0); $w = $Px.GetLength(1)
    $scale = [Math]::Floor(($Size - 8) / [Math]::Max($w, $h))
    $ox = [Math]::Floor(($Size - $w * $scale) / 2); $oy = [Math]::Floor(($Size - $h * $scale) / 2)
    $bmp = New-Object System.Drawing.Bitmap $Size, $Size
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.Clear([System.Drawing.Color]::Transparent)
    for ($y = 0; $y -lt $h; $y++) {
        for ($x = 0; $x -lt $w; $x++) {
            $c = $Px[$y, $x]
            if (-not $palette.ContainsKey($c)) { continue }
            $brush = New-Object System.Drawing.SolidBrush ([System.Drawing.ColorTranslator]::FromHtml($palette[$c]))
            $g.FillRectangle($brush, $ox + $x * $scale, $oy + $y * $scale, $scale, $scale)
            $brush.Dispose()
        }
    }
    $g.Dispose()
    $bmp.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
}

$u8 = New-Object System.Text.UTF8Encoding($false)
foreach ($mood in 'normal', 'happy', 'worried') {
    $px = New-Pixels $mood
    $title = switch ($mood) { 'happy' { 'つなぎ（よろこび）' } 'worried' { 'つなぎ（こまり）' } default { 'つなぎ' } }
    [System.IO.File]::WriteAllText((Join-Path $here "tsunagi-$mood.svg"), (ConvertTo-Svg $px $title), $u8)
    if ($mood -eq 'normal') { Save-Png $px (Join-Path $repo 'media\icon.png') 128 }
    Save-Png $px (Join-Path $here "tsunagi-$mood.png") 256
}
'written: ' + ((Get-ChildItem $here -Filter 'tsunagi-*').Name -join ', ') + ', media/icon.png'
