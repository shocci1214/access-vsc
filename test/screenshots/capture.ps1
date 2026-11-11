<#
  capture.ps1 — マニュアル用スクリーンショット。
  -Marker を起動引数に含む VS Code のウィンドウを、決まった大きさにしてから画像にする（前面に出さずに PrintWindow で撮る）。
#>
param(
    [Parameter(Mandatory = $true)][string]$Out,
    [Parameter(Mandatory = $true)][string]$Marker,
    [int]$Width = 1440,
    [int]$Height = 900
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
Add-Type @'
using System; using System.Runtime.InteropServices;
public static class Shot {
  [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
  [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int w, int hgt, uint flags);
  [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint flags);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr h);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int cmd);
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
}
'@
[void][Shot]::SetProcessDPIAware()

$main = Get-CimInstance Win32_Process -Filter "Name='Code.exe'" | Where-Object { $_.CommandLine -match [regex]::Escape($Marker) -and $_.CommandLine -notmatch '--type=' } | Select-Object -First 1
if (-not $main) { throw "VS Code window not found ($Marker)" }
$h = (Get-Process -Id $main.ProcessId).MainWindowHandle
if ($h -eq [IntPtr]::Zero) { throw 'No main window handle' }
if ([Shot]::IsIconic($h)) { [void][Shot]::ShowWindow($h, 4) }   # SW_SHOWNOACTIVATE
[void][Shot]::SetWindowPos($h, [IntPtr]::Zero, 30, 30, $Width, $Height, 0x0014)   # NOZORDER | NOACTIVATE
Start-Sleep -Milliseconds 400
$r = New-Object Shot+RECT
[void][Shot]::GetWindowRect($h, [ref]$r)
$w = $r.Right - $r.Left; $hh = $r.Bottom - $r.Top
$bmp = New-Object System.Drawing.Bitmap $w, $hh
$g = [System.Drawing.Graphics]::FromImage($bmp)
$hdc = $g.GetHdc()
$ok = [Shot]::PrintWindow($h, $hdc, 2)   # PW_RENDERFULLCONTENT
$g.ReleaseHdc($hdc); $g.Dispose()
$dir = Split-Path -Parent $Out
if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
$bmp.Save($Out, [System.Drawing.Imaging.ImageFormat]::Png)
$bmp.Dispose()
"captured $Out ($w x $hh, printwindow=$ok)"
