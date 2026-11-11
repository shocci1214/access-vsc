<#
  テスト用の Access データベースを作る（テーブル・クエリ・フォーム・モジュール入り）
  使い方: powershell -NoProfile -ExecutionPolicy Bypass -File New-TestDatabase.ps1 -Path C:\temp\test.accdb
#>
param([Parameter(Mandatory = $true)][string]$Path)
$ErrorActionPreference = 'Stop'
$Path = [System.IO.Path]::GetFullPath($Path)
if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }

$app = New-Object -ComObject Access.Application
$hwnd = $app.hWndAccessApp()
try {
    $app.Visible = $false
    $app.AutomationSecurity = 1
    $app.NewCurrentDatabase($Path)
    $db = $app.CurrentDb()
    $db.Execute('CREATE TABLE Customers (ID COUNTER CONSTRAINT PK_Customers PRIMARY KEY, CustomerName TEXT(100) NOT NULL, City TEXT(50), CreatedAt DATETIME)')
    $db.Execute("INSERT INTO Customers (CustomerName, City, CreatedAt) VALUES ('山田商店', '東京', #2026/01/15#)")
    $db.Execute("INSERT INTO Customers (CustomerName, City, CreatedAt) VALUES ('佐藤工業', '大阪', #2026/02/01 10:30:00#)")
    $db.Execute("INSERT INTO Customers (CustomerName, City) VALUES ('Suzuki & Co.', NULL)")
    $db.Execute('CREATE TABLE Orders (OrderID COUNTER PRIMARY KEY, CustomerID LONG, Amount CURRENCY, CONSTRAINT FK_Orders_Customers FOREIGN KEY (CustomerID) REFERENCES Customers (ID))')
    $db.Execute('INSERT INTO Orders (CustomerID, Amount) VALUES (1, 1200)')
    [void]$db.CreateQueryDef('qryTokyoCustomers', "SELECT * FROM Customers WHERE City = '東京'")

    # フォーム（コード付き）
    $frm = $app.CreateForm()
    $frm.RecordSource = 'Customers'
    $frm.Caption = '顧客'
    $tmpName = $frm.Name
    $c = $app.CreateControl($tmpName, 109, 0, '', 'CustomerName', 1000, 500, 3000, 300)   # acTextBox
    $c.Name = 'txtCustomerName'
    $b = $app.CreateControl($tmpName, 104, 0, '', '', 1000, 1000, 1500, 400)               # acCommandButton
    $b.Name = 'cmdHello'
    $b.Caption = 'Hello'
    $b.OnClick = '[Event Procedure]'
    $frm.HasModule = $true
    $frm.Module.AddFromString("Private Sub cmdHello_Click()`r`n    MsgBox ""こんにちは""`r`nEnd Sub")
    $app.DoCmd.Close(2, $tmpName, 1)   # acSaveYes
    $app.DoCmd.Rename('frmCustomers', 2, $tmpName)

    # レポート（ページヘッダーにタイトル、詳細に 2 列）
    $rpt = $app.CreateReport()
    $tmpRpt = $rpt.Name
    $rpt.RecordSource = 'Customers'
    $l = $app.CreateReportControl($tmpRpt, 100, 3, '', '', 500, 100, 3000, 400)   # acLabel, acPageHeader
    $l.Name = 'lblTitle'
    $l.Caption = '顧客一覧'
    $t1 = $app.CreateReportControl($tmpRpt, 109, 0, '', 'CustomerName', 500, 100, 4000, 300)   # acTextBox, acDetail
    $t1.Name = 'txtName'
    $t2 = $app.CreateReportControl($tmpRpt, 109, 0, '', 'City', 4700, 100, 2000, 300)
    $t2.Name = 'txtCity'
    $app.DoCmd.Close(3, $tmpRpt, 1)
    $app.DoCmd.Rename('rptCustomers', 3, $tmpRpt)

    # 標準モジュール
    $code = @"
Option Compare Database
Option Explicit

Public Function Greet(ByVal name As String) As String
    Greet = "こんにちは、" & name
End Function

Public Function AddNumbers(ByVal a As Long, ByVal b As Long) As Long
    AddNumbers = a + b
End Function
"@
    $tmp = Join-Path ([IO.Path]::GetTempPath()) ('mod_' + [guid]::NewGuid().ToString('N') + '.txt')
    [IO.File]::WriteAllText($tmp, $code, [Text.Encoding]::Default)
    $app.LoadFromText(5, 'modMain', $tmp)
    Remove-Item $tmp
    $app.RunCommand(126)
    "created: $Path (compiled=$($app.IsCompiled))"
}
finally {
    try { $app.CloseCurrentDatabase() } catch {}
    try { $app.Quit() } catch {}
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
    $app = $null
    [GC]::Collect()
    Add-Type -Namespace T -Name U -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr h, out uint p);'
    $procId = 0; [void][T.U]::GetWindowThreadProcessId([IntPtr][long]$hwnd, [ref]$procId)
    if ($procId -gt 0) {
        $p = Get-Process -Id $procId -ErrorAction SilentlyContinue
        if ($p -and -not $p.WaitForExit(8000)) { Stop-Process -Id $procId -Force }
    }
}
