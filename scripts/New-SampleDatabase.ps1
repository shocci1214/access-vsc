<#
  New-SampleDatabase.ps1 — チュートリアル用のサンプル データベース（小さな受注管理）を作る。
  使い方: powershell -NoProfile -ExecutionPolicy Bypass -File New-SampleDatabase.ps1 -Path C:\work\tutorial\shop.accdb
  中身:
    テーブル  Customers（顧客）/ Products（商品）/ Orders（受注）
    クエリ    qryOrderList（受注一覧）
    フォーム  frmCustomers（顧客）
    レポート  rptOrders（受注一覧の帳票）
    モジュール modMain（あいさつの関数）
#>
param([Parameter(Mandatory = $true)][string]$Path)
$ErrorActionPreference = 'Stop'
$Path = [System.IO.Path]::GetFullPath($Path)
if (Test-Path -LiteralPath $Path) { throw "既にファイルがあります: $Path" }
$dir = Split-Path -Parent $Path
if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

$app = New-Object -ComObject Access.Application
$hwnd = $app.hWndAccessApp()
try {
    $app.Visible = $false
    $app.AutomationSecurity = 1
    $app.NewCurrentDatabase($Path)
    $db = $app.CurrentDb()

    $db.Execute('CREATE TABLE Customers (ID COUNTER CONSTRAINT PK_Customers PRIMARY KEY, CustomerName TEXT(100) NOT NULL, City TEXT(50))')
    foreach ($c in @(@('山田商店', '東京'), @('佐藤工業', '大阪'), @('鈴木食品', '名古屋'), @('田中書店', '東京'), @('高橋電機', '福岡'))) {
        $db.Execute("INSERT INTO Customers (CustomerName, City) VALUES ('$($c[0])', '$($c[1])')")
    }
    $db.Execute('CREATE TABLE Products (ID COUNTER CONSTRAINT PK_Products PRIMARY KEY, ProductName TEXT(100) NOT NULL, Price CURRENCY NOT NULL)')
    foreach ($p in @(@('ノート', 150), @('ボールペン', 99), @('ファイル', 230), @('電卓', 1980))) {
        $db.Execute("INSERT INTO Products (ProductName, Price) VALUES ('$($p[0])', $($p[1]))")
    }
    $db.Execute('CREATE TABLE Orders (ID COUNTER CONSTRAINT PK_Orders PRIMARY KEY, OrderDate DATETIME NOT NULL, CustomerID LONG NOT NULL, ProductID LONG NOT NULL, Qty LONG NOT NULL, CONSTRAINT FK_Orders_Customers FOREIGN KEY (CustomerID) REFERENCES Customers (ID), CONSTRAINT FK_Orders_Products FOREIGN KEY (ProductID) REFERENCES Products (ID))')
    foreach ($o in @(@('2026/04/01', 1, 1, 10), @('2026/04/03', 2, 4, 2), @('2026/04/05', 1, 2, 30), @('2026/04/08', 3, 3, 5), @('2026/04/10', 4, 1, 12), @('2026/04/12', 5, 4, 1), @('2026/04/15', 2, 2, 50))) {
        $db.Execute("INSERT INTO Orders (OrderDate, CustomerID, ProductID, Qty) VALUES (#$($o[0])#, $($o[1]), $($o[2]), $($o[3]))")
    }
    [void]$db.CreateQueryDef('qryOrderList', 'SELECT o.ID, o.OrderDate, c.CustomerName, p.ProductName, p.Price, o.Qty, p.Price * o.Qty AS Amount FROM (Orders AS o INNER JOIN Customers AS c ON o.CustomerID = c.ID) INNER JOIN Products AS p ON o.ProductID = p.ID ORDER BY o.OrderDate')

    # フォーム
    $frm = $app.CreateForm()
    $tmp = $frm.Name
    $frm.RecordSource = 'Customers'
    $frm.Caption = '顧客'
    $t = $app.CreateControl($tmp, 109, 0, '', 'CustomerName', 1000, 500, 3000, 300); $t.Name = 'txtCustomerName'
    $t = $app.CreateControl($tmp, 109, 0, '', 'City', 1000, 900, 2000, 300); $t.Name = 'txtCity'
    $app.DoCmd.Close(2, $tmp, 1)
    $app.DoCmd.Rename('frmCustomers', 2, $tmp)

    # レポート（受注一覧）
    $rpt = $app.CreateReport()
    $tmp = $rpt.Name
    $rpt.RecordSource = 'qryOrderList'
    $l = $app.CreateReportControl($tmp, 100, 3, '', '', 300, 100, 4000, 450); $l.Name = 'lblTitle'; $l.Caption = '受注一覧'; $l.FontSize = 14
    $cols = @(@('OrderDate', '日付', 300, 1300), @('CustomerName', '顧客', 1700, 2200), @('ProductName', '商品', 4000, 1700), @('Qty', '数量', 5800, 800), @('Amount', '金額', 6700, 1300))
    foreach ($c in $cols) {
        $h = $app.CreateReportControl($tmp, 100, 3, '', '', $c[2], 700, $c[3], 300); $h.Name = 'lbl' + $c[0]; $h.Caption = $c[1]
        $x = $app.CreateReportControl($tmp, 109, 0, '', $c[0], $c[2], 60, $c[3], 300); $x.Name = 'txt' + $c[0]
    }
    $rpt.Section(0).Height = 420
    $rpt.Section(3).Height = 1100
    $app.DoCmd.Close(3, $tmp, 1)
    $app.DoCmd.Rename('rptOrders', 3, $tmp)

    # 標準モジュール
    $code = "Option Compare Database`r`nOption Explicit`r`n`r`n' あいさつの文章を返す`r`nPublic Function Greet(ByVal name As String) As String`r`n    Greet = ""こんにちは、"" & name & ""さん""`r`nEnd Function`r`n"
    $f = Join-Path ([IO.Path]::GetTempPath()) ('mod_' + [guid]::NewGuid().ToString('N') + '.txt')
    [IO.File]::WriteAllText($f, $code, [Text.Encoding]::Default)
    $app.LoadFromText(5, 'modMain', $f)
    Remove-Item $f
    $app.RunCommand(126)
    "created: $Path"
}
finally {
    try { $app.CloseCurrentDatabase() } catch {}
    try { $app.Quit() } catch {}
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
    $app = $null
    [GC]::Collect()
    Add-Type -Namespace AccessVscSample -Name U -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr h, out uint p);'
    $procId = 0; [void][AccessVscSample.U]::GetWindowThreadProcessId([IntPtr][long]$hwnd, [ref]$procId)
    if ($procId -gt 0) {
        $p = Get-Process -Id $procId -ErrorAction SilentlyContinue
        if ($p -and -not $p.WaitForExit(8000)) { Stop-Process -Id $procId -Force }
    }
}
