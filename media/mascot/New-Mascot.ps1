<#
  New-Mascot.ps1 — マスコット「つなぎ」の SVG（3 つの表情）と、拡張機能のアイコン（PNG）を作る。
    つなぎ：おもちのような丸いからだ。水色（エディタ）からピンク（データベース）へ、やわらかく混ざり合う色で
    「つながる」を表す。頭のアホ毛の先は小さなハート。
  使い方: powershell -NoProfile -ExecutionPolicy Bypass -File media\mascot\New-Mascot.ps1
  出力: media/mascot/tsunagi-{normal,happy,worried}.svg、docs/images/ に同じもの、media/icon.png（128x128）
  PNG は Microsoft Edge（ヘッドレス）で SVG を描画して作る。
#>
$ErrorActionPreference = 'Stop'
$here = $PSScriptRoot
$repo = (Resolve-Path (Join-Path $here '..\..')).Path
$u8 = New-Object System.Text.UTF8Encoding($false)

$line = '#5B4566'   # 輪郭（あたたかみのある濃い紫）
$ink = '#3B3350'    # 目・口

function New-MascotSvg([string]$Mood, [string]$Title) {
    $face = switch ($Mood) {
        'happy' {
            @"
  <path d="M65 119 Q76 105 87 119" fill="none" stroke="$ink" stroke-width="5" stroke-linecap="round"/>
  <path d="M113 119 Q124 105 135 119" fill="none" stroke="$ink" stroke-width="5" stroke-linecap="round"/>
  <path d="M88 130 Q100 150 112 130 Z" fill="#8C3A55" stroke="$ink" stroke-width="3" stroke-linejoin="round"/>
  <path d="M94 139 Q100 145 106 139 Q100 135 94 139 Z" fill="#FF8FA8"/>
  <path d="M38 50 L42 60 L52 64 L42 68 L38 78 L34 68 L24 64 L34 60 Z" fill="#FFD86B" stroke="#E7B23C" stroke-width="1.5" stroke-linejoin="round"/>
  <path d="M166 40 L169 47 L176 50 L169 53 L166 60 L163 53 L156 50 L163 47 Z" fill="#FFD86B" stroke="#E7B23C" stroke-width="1.5" stroke-linejoin="round"/>
  <circle cx="176" cy="78" r="3.5" fill="#FFD86B"/>
"@
        }
        'worried' {
            @"
  <ellipse cx="76" cy="118" rx="8.5" ry="11" fill="$ink"/>
  <ellipse cx="124" cy="118" rx="8.5" ry="11" fill="$ink"/>
  <circle cx="79" cy="113" r="3.2" fill="#fff"/><circle cx="127" cy="113" r="3.2" fill="#fff"/>
  <path d="M64 101 Q74 96 84 98" fill="none" stroke="$ink" stroke-width="3.5" stroke-linecap="round"/>
  <path d="M116 98 Q126 96 136 101" fill="none" stroke="$ink" stroke-width="3.5" stroke-linecap="round"/>
  <path d="M89 138 Q94.5 133 100 138 Q105.5 143 111 138" fill="none" stroke="$ink" stroke-width="3.2" stroke-linecap="round"/>
  <path d="M160 70 C155 80 152 85 152 90 A8 8 0 0 0 168 90 C168 85 165 80 160 70 Z" fill="#A8DDFF" stroke="#6FB6E8" stroke-width="2"/>
  <ellipse cx="157" cy="88" rx="2" ry="3" fill="#fff" opacity=".8"/>
"@
        }
        default {
            @"
  <ellipse cx="76" cy="116" rx="9.5" ry="12.5" fill="$ink"/>
  <ellipse cx="124" cy="116" rx="9.5" ry="12.5" fill="$ink"/>
  <circle cx="79.5" cy="110.5" r="3.8" fill="#fff"/><circle cx="73" cy="121" r="1.9" fill="#fff"/>
  <circle cx="127.5" cy="110.5" r="3.8" fill="#fff"/><circle cx="121" cy="121" r="1.9" fill="#fff"/>
  <path d="M91 132 Q95.5 138.5 100 133 Q104.5 138.5 109 132" fill="none" stroke="$ink" stroke-width="3.2" stroke-linecap="round" stroke-linejoin="round"/>
"@
        }
    }
    return @"
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 200" width="200" height="200" role="img" aria-label="$Title">
  <title>$Title</title>
  <defs>
    <linearGradient id="tsunagi-body" x1="22" y1="0" x2="178" y2="0" gradientUnits="userSpaceOnUse">
      <stop offset="0" stop-color="#8CCBFF"/>
      <stop offset=".5" stop-color="#C9B6F7"/>
      <stop offset="1" stop-color="#FF9DBA"/>
    </linearGradient>
    <radialGradient id="tsunagi-shine" cx=".32" cy=".25" r=".7">
      <stop offset="0" stop-color="#fff" stop-opacity=".55"/>
      <stop offset=".45" stop-color="#fff" stop-opacity=".12"/>
      <stop offset="1" stop-color="#fff" stop-opacity="0"/>
    </radialGradient>
    <linearGradient id="tsunagi-heart" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0" stop-color="#FFB3C8"/>
      <stop offset="1" stop-color="#FF6F98"/>
    </linearGradient>
  </defs>
  <ellipse cx="100" cy="187" rx="56" ry="7" fill="$line" opacity=".14"/>
  <!-- アホ毛と、先のハート -->
  <path d="M100 47 C96 36 104 28 112 30" fill="none" stroke="$line" stroke-width="3.5" stroke-linecap="round"/>
  <path d="M114 32 C110 25 101 26 102 33 C103 39 114 44 116 46 C118 42 127 34 125 28 C123 22 116 24 114 32 Z" fill="url(#tsunagi-heart)" stroke="$line" stroke-width="3" stroke-linejoin="round"/>
  <!-- 手 -->
  <ellipse cx="27" cy="132" rx="13" ry="10" transform="rotate(-25 27 132)" fill="#9FD3FF" stroke="$line" stroke-width="3.5"/>
  <ellipse cx="173" cy="132" rx="13" ry="10" transform="rotate(25 173 132)" fill="#FFAEC6" stroke="$line" stroke-width="3.5"/>
  <!-- 足 -->
  <ellipse cx="76" cy="176" rx="17" ry="10" fill="#A9D7FF" stroke="$line" stroke-width="3.5"/>
  <ellipse cx="124" cy="176" rx="17" ry="10" fill="#FFB6CB" stroke="$line" stroke-width="3.5"/>
  <!-- からだ（おもち） -->
  <path d="M100 46 C147 46 179 80 179 122 C179 160 147 180 100 180 C53 180 21 160 21 122 C21 80 53 46 100 46 Z" fill="url(#tsunagi-body)" stroke="$line" stroke-width="3.5"/>
  <path d="M100 46 C147 46 179 80 179 122 C179 160 147 180 100 180 C53 180 21 160 21 122 C21 80 53 46 100 46 Z" fill="url(#tsunagi-shine)"/>
  <ellipse cx="68" cy="76" rx="20" ry="11" transform="rotate(-22 68 76)" fill="#fff" opacity=".55"/>
  <ellipse cx="90" cy="64" rx="5" ry="3.5" transform="rotate(-22 90 64)" fill="#fff" opacity=".6"/>
  <!-- ほっぺ -->
  <ellipse cx="58" cy="134" rx="12" ry="7" fill="#FF7FA3" opacity=".5"/>
  <ellipse cx="142" cy="134" rx="12" ry="7" fill="#FF7FA3" opacity=".5"/>
  <!-- かお -->
$face
</svg>
"@
}

$titles = @{ normal = 'つなぎ'; happy = 'つなぎ（よろこび）'; worried = 'つなぎ（こまり）' }
foreach ($mood in 'normal', 'happy', 'worried') {
    $svg = New-MascotSvg $mood $titles[$mood]
    [System.IO.File]::WriteAllText((Join-Path $here "tsunagi-$mood.svg"), $svg, $u8)
    [System.IO.File]::WriteAllText((Join-Path $repo "docs\images\tsunagi-$mood.svg"), $svg, $u8)
}
Remove-Item (Join-Path $here 'tsunagi-*.png') -ErrorAction SilentlyContinue

# アイコン（PNG）：Edge で 512px に描画して、128px に縮小（なめらかに）
$edge = @("${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe", "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe") | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $edge) { Write-Warning 'Microsoft Edge が見つからないため、アイコン（PNG）は作りませんでした'; return }
$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('tsunagi_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmp | Out-Null
try {
    $html = Join-Path $tmp 'icon.html'
    [System.IO.File]::WriteAllText($html, "<!DOCTYPE html><html><body style='margin:0;background:transparent'><img src='file:///$((Join-Path $here 'tsunagi-normal.svg').Replace('\', '/'))' style='display:block;width:512px;height:512px'></body></html>", $u8)
    $shot = Join-Path $tmp 'icon512.png'
    $p = Start-Process -FilePath $edge -ArgumentList @('--headless=new', '--disable-gpu', '--hide-scrollbars', "--user-data-dir=`"$tmp\ud`"", '--default-background-color=00000000', '--window-size=512,512', "--screenshot=`"$shot`"", "`"file:///$($html.Replace('\', '/'))`"") -PassThru -WindowStyle Hidden
    if (-not $p.WaitForExit(30000)) { $p.Kill() }
    Add-Type -AssemblyName System.Drawing
    $src = [System.Drawing.Image]::FromFile($shot)
    $bmp = New-Object System.Drawing.Bitmap 128, 128
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
    $g.Clear([System.Drawing.Color]::Transparent)
    $g.DrawImage($src, 0, 0, 128, 128)
    $g.Dispose(); $src.Dispose()
    $bmp.Save((Join-Path $repo 'media\icon.png'), [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
}
finally { Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue }
'written: media/mascot/tsunagi-{normal,happy,worried}.svg, docs/images/tsunagi-*.svg, media/icon.png'
