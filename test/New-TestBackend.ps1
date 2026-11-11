<#
  テスト用：バックエンド DB（開発用・本番用）を作り、フロントエンドに Products をリンクする
  使い方: powershell -File New-TestBackend.ps1 -FrontEnd C:\x\shop.accdb -Dir C:\x
#>
param([Parameter(Mandatory = $true)][string]$FrontEnd, [Parameter(Mandatory = $true)][string]$Dir)
$ErrorActionPreference = 'Stop'
$dev = Join-Path $Dir 'shop_be_dev.accdb'
$prod = Join-Path $Dir 'shop_be_prod.accdb'
foreach ($p in $dev, $prod) { if (Test-Path -LiteralPath $p) { Remove-Item -LiteralPath $p -Force } }

$engine = New-Object -ComObject DAO.DBEngine.120
$d = $engine.CreateDatabase($dev, ';LANGID=0x0411;CP=932;COUNTRY=0')
$d.Execute('CREATE TABLE Products (ID COUNTER PRIMARY KEY, ProductName TEXT(50))')
$d.Execute("INSERT INTO Products (ProductName) VALUES ('dev item')")
$d.Close()
[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($engine)
Copy-Item -LiteralPath $dev -Destination $prod
$engine = New-Object -ComObject DAO.DBEngine.120
$d = $engine.OpenDatabase($prod)
$d.Execute("INSERT INTO Products (ProductName) VALUES ('prod item 2')")
$d.Execute("INSERT INTO Products (ProductName) VALUES ('prod item 3')")
$d.Close()
[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($engine)

$app = New-Object -ComObject Access.Application
$hwnd = $app.hWndAccessApp()
try {
    $app.Visible = $false
    $app.AutomationSecurity = 1
    $app.OpenCurrentDatabase($FrontEnd)
    $app.DoCmd.TransferDatabase(2, 'Microsoft Access', $dev, 0, 'Products', 'Products')   # acLink, acTable
    "linked Products -> $dev"
}
finally {
    try { $app.CloseCurrentDatabase() } catch {}
    try { $app.Quit() } catch {}
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
    $app = $null
    [GC]::Collect()
    Add-Type -Namespace T3 -Name U -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr h, out uint p);'
    $procId = 0; [void][T3.U]::GetWindowThreadProcessId([IntPtr][long]$hwnd, [ref]$procId)
    if ($procId -gt 0) {
        $p = Get-Process -Id $procId -ErrorAction SilentlyContinue
        if ($p -and -not $p.WaitForExit(8000)) { Stop-Process -Id $procId -Force }
    }
}
