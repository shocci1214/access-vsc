<#
  New-SampleDatabase.ps1 — チュートリアル用のサンプル データベース「カフェ こもれび」を作る。
  使い方: powershell -NoProfile -ExecutionPolicy Bypass -File New-SampleDatabase.ps1 -Path C:\work\tutorial\cafe.accdb
  中身:
    テーブル   Products（メニュー）/ Members（会員・スタンプ）/ Orders（注文）/ OrderDetails（注文明細。行ごとに店内・テイクアウト）
    クエリ     qrySalesDetail（売上明細。金額・税率つき）
    フォーム   frmOrder（注文入力。テイクアウトでも税率が 10% のままという「不具合」入り。中級編で直す）
    レポート   rptReceipt（レシート）/ rptMonthlySales（月次売上。分類ごとの小計）
    モジュール modCafe（税込金額・今日のラッキーメニュー）
#>
param([Parameter(Mandatory = $true)][string]$Path)
$ErrorActionPreference = 'Stop'
$Path = [System.IO.Path]::GetFullPath($Path)
if (Test-Path -LiteralPath $Path) { throw "既にファイルがあります: $Path" }
$dir = Split-Path -Parent $Path
if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

function Set-Props($obj, [hashtable]$props) { foreach ($k in $props.Keys) { $obj.$k = $props[$k] } }

$app = New-Object -ComObject Access.Application
$hwnd = $app.hWndAccessApp()
try {
    $app.Visible = $false
    $app.AutomationSecurity = 1
    $app.NewCurrentDatabase($Path)
    $db = $app.CurrentDb()

    # ---------------- テーブルとデータ ----------------
    $db.Execute('CREATE TABLE Products (ID COUNTER CONSTRAINT PK_Products PRIMARY KEY, ProductName TEXT(50) NOT NULL, Category TEXT(20) NOT NULL, Price CURRENCY NOT NULL)')
    foreach ($p in @(
            @('ブレンドコーヒー', 'ドリンク', 450), @('カフェラテ', 'ドリンク', 500), @('抹茶ラテ', 'ドリンク', 550), @('クリームソーダ', 'ドリンク', 600),
            @('たまごサンド', 'フード', 650), @('ナポリタン', 'フード', 900), @('チーズケーキ', 'デザート', 480), @('季節のパフェ', 'デザート', 980))) {
        $db.Execute("INSERT INTO Products (ProductName, Category, Price) VALUES ('$($p[0])', '$($p[1])', $($p[2]))")
    }
    $db.Execute('CREATE TABLE Members (ID COUNTER CONSTRAINT PK_Members PRIMARY KEY, MemberName TEXT(50) NOT NULL, Stamps LONG NOT NULL)')
    foreach ($m in @(@('山田 花子', 8), @('佐藤 健', 3), @('鈴木 愛', 10), @('田中 誠', 0), @('高橋 美咲', 5))) {
        $db.Execute("INSERT INTO Members (MemberName, Stamps) VALUES ('$($m[0])', $($m[1]))")
    }
    $db.Execute('CREATE TABLE Orders (ID COUNTER CONSTRAINT PK_Orders PRIMARY KEY, OrderDate DATETIME NOT NULL, MemberID LONG, CONSTRAINT FK_Orders_Members FOREIGN KEY (MemberID) REFERENCES Members (ID))')
    $db.Execute('CREATE TABLE OrderDetails (ID COUNTER CONSTRAINT PK_OrderDetails PRIMARY KEY, OrderID LONG NOT NULL, ProductID LONG NOT NULL, Qty LONG NOT NULL, Takeout YESNO, CONSTRAINT FK_Details_Orders FOREIGN KEY (OrderID) REFERENCES Orders (ID), CONSTRAINT FK_Details_Products FOREIGN KEY (ProductID) REFERENCES Products (ID))')
    $orders = @(
        @('2026/04/01', 1), @('2026/04/01', $null), @('2026/04/02', 2), @('2026/04/03', 3), @('2026/04/05', $null), @('2026/04/06', 1),
        @('2026/04/08', 4), @('2026/04/10', 5), @('2026/04/12', $null), @('2026/04/15', 2), @('2026/04/18', 3), @('2026/04/20', 1))
    foreach ($o in $orders) {
        $mid = if ($null -eq $o[1]) { 'Null' } else { $o[1] }
        $db.Execute("INSERT INTO Orders (OrderDate, MemberID) VALUES (#$($o[0])#, $mid)")
    }
    # 注文 ID, 商品 ID, 数量, テイクアウト
    $details = @(
        @(1, 3, 1, 0), @(1, 8, 1, 0), @(1, 7, 2, 1),
        @(2, 2, 2, 1), @(2, 5, 1, 1),
        @(3, 6, 1, 0), @(3, 4, 1, 0),
        @(4, 1, 1, 0), @(4, 7, 1, 0),
        @(5, 1, 3, 1),
        @(6, 2, 1, 0), @(6, 5, 1, 0),
        @(7, 8, 2, 0),
        @(8, 3, 2, 1), @(8, 7, 1, 1),
        @(9, 1, 1, 0), @(9, 6, 2, 0),
        @(10, 4, 1, 0), @(10, 7, 1, 0),
        @(11, 2, 1, 0), @(11, 8, 1, 0),
        @(12, 1, 2, 0), @(12, 7, 1, 1))
    foreach ($d in $details) {
        $db.Execute("INSERT INTO OrderDetails (OrderID, ProductID, Qty, Takeout) VALUES ($($d[0]), $($d[1]), $($d[2]), $(if ($d[3]) { 'True' } else { 'False' }))")
    }
    [void]$db.CreateQueryDef('qrySalesDetail', 'SELECT d.ID, d.OrderID, o.OrderDate, m.MemberName, p.ProductName, p.Category, p.Price, d.Qty, p.Price * d.Qty AS Amount, d.Takeout, IIf(d.Takeout, 0.08, 0.1) AS TaxRate FROM ((OrderDetails AS d INNER JOIN Orders AS o ON d.OrderID = o.ID) INNER JOIN Products AS p ON d.ProductID = p.ID) LEFT JOIN Members AS m ON o.MemberID = m.ID ORDER BY d.OrderID, d.ID')

    # ---------------- 標準モジュール ----------------
    $code = @'
Option Compare Database
Option Explicit

' 消費税率（店内 10%、テイクアウトは軽減税率 8%）
Public Function TaxRate(ByVal takeout As Boolean) As Double
    If takeout Then
        TaxRate = 0.08
    Else
        TaxRate = 0.1
    End If
End Function

' 税込金額（1 円未満切り捨て）
Public Function TaxIncluded(ByVal price As Currency, ByVal takeout As Boolean) As Currency
    TaxIncluded = Int(price * (1 + TaxRate(takeout)))
End Function

' 今日のラッキーメニュー（日付で決まるおみくじ）。日付を省略すると今日
Public Function PickOfTheDay(Optional ByVal d As Date = 0) As String
    Dim fortunes As Variant
    Dim n As Long
    Dim menuCount As Long
    If d = 0 Then d = Date
    fortunes = Array("大吉", "中吉", "小吉", "吉", "末吉")
    n = CLng(Format(d, "yyyymmdd"))
    menuCount = DCount("*", "Products")
    PickOfTheDay = "【" & fortunes(n Mod 5) & "】今日のラッキーメニューは「" & DLookup("ProductName", "Products", "ID=" & ((n Mod menuCount) + 1)) & "」！"
End Function
'@
    $f = Join-Path ([IO.Path]::GetTempPath()) ('mod_' + [guid]::NewGuid().ToString('N') + '.txt')
    [IO.File]::WriteAllText($f, $code.Replace("`r`n", "`n").Replace("`n", "`r`n"), [Text.Encoding]::Default)
    $app.LoadFromText(5, 'modCafe', $f)
    Remove-Item $f

    # ---------------- フォーム：注文入力 ----------------
    $frm = $app.CreateForm()
    $tmp = $frm.Name
    Set-Props $frm @{ Caption = '注文入力（カフェ こもれび）'; RecordSelectors = $false; NavigationButtons = $false; DividingLines = $false; ScrollBars = 0; Width = 6800 }
    $frm.Section(0).Height = 4700
    $lbl = { param($name, $caption, $left, $top, $w) $c = $app.CreateControl($tmp, 100, 0, '', '', $left, $top, $w, 330); $c.Name = $name; $c.Caption = $caption; $c }
    [void](& $lbl 'lblProduct' '商品' 300 300 1200)
    $c = $app.CreateControl($tmp, 111, 0, '', '', 1700, 300, 2800, 330)
    Set-Props $c @{ Name = 'cboProduct'; RowSource = 'SELECT ID, ProductName, Price FROM Products ORDER BY ID'; ColumnCount = 3; ColumnWidths = '0;2268;850'; BoundColumn = 1; LimitToList = $true; ListWidth = 3300 }
    [void](& $lbl 'lblQty' '数量' 300 750 1200)
    $c = $app.CreateControl($tmp, 109, 0, '', '', 1700, 750, 900, 330); $c.Name = 'txtQty'; $c.TextAlign = 3
    $c = $app.CreateControl($tmp, 106, 0, '', '', 1700, 1230, 260, 260); $c.Name = 'chkTakeout'
    [void](& $lbl 'lblTakeout' 'テイクアウト（持ち帰り）' 2050 1180 2600)
    $c = $app.CreateControl($tmp, 104, 0, '', '', 1700, 1650, 1400, 450); $c.Name = 'cmdCalc'; $c.Caption = '計算'; $c.OnClick = '[Event Procedure]'
    [void](& $lbl 'lblSubtotal' '税抜金額' 300 2300 1200)
    $c = $app.CreateControl($tmp, 109, 0, '', '', 1700, 2300, 1600, 330); Set-Props $c @{ Name = 'txtSubtotal'; Format = '#,##0'; TextAlign = 3; Locked = $true }
    [void](& $lbl 'lblTotal' '税込金額' 300 2750 1200)
    $c = $app.CreateControl($tmp, 109, 0, '', '', 1700, 2750, 1600, 330); Set-Props $c @{ Name = 'txtTotal'; Format = '#,##0'; TextAlign = 3; Locked = $true; FontWeight = 700 }
    $c = $app.CreateControl($tmp, 104, 0, '', '', 300, 3500, 3000, 450); $c.Name = 'cmdPick'; $c.Caption = '今日のラッキーメニュー'; $c.OnClick = '[Event Procedure]'
    [void](& $lbl 'lblPick' 'ボタンを押してね' 300 4050 6200)
    $frm.OnLoad = '[Event Procedure]'
    $formCode = @'
Option Compare Database
Option Explicit

Private Sub Form_Load()
    Me.txtQty = 1
End Sub

' 税抜金額と税込金額を計算する
Private Sub UpdateTotals()
    Dim price As Currency
    If IsNull(Me.cboProduct) Then
        Me.txtSubtotal = Null
        Me.txtTotal = Null
        Exit Sub
    End If
    price = DLookup("Price", "Products", "ID=" & Me.cboProduct)
    Me.txtSubtotal = price * Nz(Me.txtQty, 0)
    Me.txtTotal = TaxIncluded(Me.txtSubtotal, False)
End Sub

Private Sub cmdCalc_Click()
    UpdateTotals
End Sub

Private Sub cmdPick_Click()
    Me.lblPick.Caption = PickOfTheDay()
End Sub
'@
    # 新しいフォームのモジュールには Option 行が最初から入っているので、中身を入れ替える
    $mod = $frm.Module
    if ($mod.CountOfLines -gt 0) { $mod.DeleteLines(1, $mod.CountOfLines) }
    $mod.InsertText($formCode.Replace("`r`n", "`n").Replace("`n", "`r`n"))
    $app.DoCmd.Close(2, $tmp, 1)
    $app.DoCmd.Rename('frmOrder', 2, $tmp)

    $rc = {
        param($type, $sec, $name, $source, $left, $top, $w, $h)
        $c = $app.CreateReportControl($tmp, $type, $sec, '', $source, $left, $top, $w, $h); $c.Name = $name
        if ($type -eq 109) { $c.BorderStyle = 0 }   # 帳票の項目は枠なし
        $c
    }
    # 帳票の行の縞模様（1 行おきの灰色）をなくす
    $plain = { param($r, [int[]]$secs) foreach ($s in $secs) { $r.Section($s).AlternateBackColor = 16777215 } }

    # ---------------- レポート：レシート ----------------
    $rpt = $app.CreateReport()
    $tmp = $rpt.Name
    Set-Props $rpt @{ RecordSource = 'qrySalesDetail'; Caption = 'レシート'; Width = 4800 }
    [void]$app.CreateGroupLevel($tmp, 'OrderID', $true, $true)
    $rpt.Section(3).Height = 0; $rpt.Section(3).Visible = $false
    $rpt.Section(4).Height = 0; $rpt.Section(4).Visible = $false
    $c = & $rc 100 5 'lblShop' '' 0 100 4800 450; Set-Props $c @{ Caption = 'カフェ こもれび'; FontSize = 14; FontWeight = 700; TextAlign = 2 }
    $c = & $rc 100 5 'lblTitle' '' 0 600 4800 350; Set-Props $c @{ Caption = '領収書'; TextAlign = 2 }
    $c = & $rc 109 5 'txtOrderDate' 'OrderDate' 100 1050 2200 330; $c.Format = 'yyyy/mm/dd'
    $c = & $rc 109 5 'txtOrderNo' '="No. " & [OrderID]' 2500 1050 2200 330; $c.TextAlign = 3
    $rpt.Section(5).Height = 1450
    [void](& $rc 109 0 'txtProductName' 'ProductName' 100 30 2400 300)
    $c = & $rc 109 0 'txtQty' 'Qty' 2550 30 500 300; $c.TextAlign = 3
    $c = & $rc 109 0 'txtAmount' 'Amount' 3100 30 1600 300; Set-Props $c @{ Format = '#,##0'; TextAlign = 3 }
    $rpt.Section(0).Height = 360
    $c = & $rc 100 6 'lblSubtotal' '' 100 100 2000 330; $c.Caption = '小計（税抜）'
    $c = & $rc 109 6 'txtSubtotal' '=Sum([Amount])' 2600 100 2100 330; Set-Props $c @{ Format = '#,##0'; TextAlign = 3 }
    $c = & $rc 100 6 'lblTotal' '' 100 500 2000 330; Set-Props $c @{ Caption = '合計（税込）'; FontWeight = 700 }
    $c = & $rc 109 6 'txtTotal' '=Sum([Amount])+Int(Sum(IIf([Takeout],[Amount],0))*0.08)+Int(Sum(IIf([Takeout],0,[Amount]))*0.1)' 2600 500 2100 330; Set-Props $c @{ Format = '#,##0'; TextAlign = 3; FontWeight = 700 }
    $rpt.Section(6).Height = 950
    $rpt.Section(6).ForceNewPage = 2
    [void](& $rc 102 5 'linHeader' '' 100 1420 4600 0)
    [void](& $rc 102 6 'linFooter' '' 100 40 4600 0)
    & $plain $rpt @(0, 5, 6)
    $app.DoCmd.Close(3, $tmp, 1)
    $app.DoCmd.Rename('rptReceipt', 3, $tmp)

    # 続けて CreateReport すると「名前を付けて保存」で止まることがあるため、いったん開き直す
    $app.CloseCurrentDatabase()
    $app.OpenCurrentDatabase($Path)

    # ---------------- レポート：月次売上 ----------------
    $rpt = $app.CreateReport()
    $tmp = $rpt.Name
    Set-Props $rpt @{ RecordSource = 'qrySalesDetail'; Caption = '月次売上'; Width = 6300 }
    [void]$app.CreateGroupLevel($tmp, 'Category', $true, $true)
    [void]$app.CreateGroupLevel($tmp, 'ProductName', $false, $true)
    $c = & $rc 100 3 'lblTitle' '' 300 100 5000 450; Set-Props $c @{ Caption = '月次売上（2026年4月）'; FontSize = 14; FontWeight = 700 }
    $c = & $rc 100 3 'lblProductName' '' 500 700 3000 300; $c.Caption = '商品'
    $c = & $rc 100 3 'lblQty' '' 3600 700 800 300; Set-Props $c @{ Caption = '数量'; TextAlign = 3 }
    $c = & $rc 100 3 'lblAmount' '' 4600 700 1400 300; Set-Props $c @{ Caption = '売上（税抜）'; TextAlign = 3 }
    $rpt.Section(3).Height = 1050
    $rpt.Section(4).Height = 0
    $c = & $rc 109 5 'txtCategory' 'Category' 300 60 3000 330; $c.FontWeight = 700
    $rpt.Section(5).Height = 420
    $rpt.Section(0).Height = 0; $rpt.Section(0).Visible = $false
    [void](& $rc 109 8 'txtProductName' 'ProductName' 500 30 3000 300)
    $c = & $rc 109 8 'txtQty' '=Sum([Qty])' 3600 30 800 300; $c.TextAlign = 3
    $c = & $rc 109 8 'txtAmount' '=Sum([Amount])' 4600 30 1400 300; Set-Props $c @{ Format = '#,##0'; TextAlign = 3 }
    $rpt.Section(8).Height = 360
    $c = & $rc 100 6 'lblCategoryTotal' '' 2600 60 900 330; $c.Caption = '小計'
    $c = & $rc 109 6 'txtCategoryQty' '=Sum([Qty])' 3600 60 800 330; Set-Props $c @{ TextAlign = 3; FontWeight = 700 }
    $c = & $rc 109 6 'txtCategoryAmount' '=Sum([Amount])' 4600 60 1400 330; Set-Props $c @{ Format = '#,##0'; TextAlign = 3; FontWeight = 700 }
    $rpt.Section(6).Height = 500
    & $plain $rpt @(5, 6, 8)
    $app.DoCmd.Close(3, $tmp, 1)
    $app.DoCmd.Rename('rptMonthlySales', 3, $tmp)

    $app.RunCommand(126)
    if (-not $app.IsCompiled) { throw 'サンプルの VBA がコンパイルできませんでした' }
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
