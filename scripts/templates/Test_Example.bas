Option Compare Database
Option Explicit

' ============================================================
' テストの書き方の例（Access VSC）
'   - 名前が Test_ で始まる、引数なしの Public Sub がテストです
'   - TestSetup / TestTeardown があれば、各テストの前後に呼ばれます
'   - テストは DB のコピーの上で実行されるので、データを書き換えても元の DB は変わりません
'     （ただしリンクテーブルの先＝バックエンドは本物なので注意）
'   - 確認には Assert モジュールを使います
'       Assert.AreEqual 期待値, 実際の値
'       Assert.IsTrue 条件
'   - 実行: サイドバーの「テストを実行」、VS Code の「テスト」ビュー、または
'           access.ps1 test -DbPath <DB>
' ============================================================

Public Sub TestSetup()
    ' 各テストの前に呼ばれる（テスト用データの準備など）
End Sub

Public Sub TestTeardown()
    ' 各テストの後に呼ばれる（後片付け）
End Sub

Public Sub Test_Addition()
    Assert.AreEqual 3, 1 + 2
End Sub

Public Sub Test_StringConcat()
    Assert.AreEqual "Access VSC", "Access" & " " & "VSC"
End Sub

Public Sub Test_Rounding()
    ' 銀行型丸め（偶数丸め）になることの確認
    Assert.AreEqual 2, Round(2.5)
    Assert.AreClose 0.3, 0.1 + 0.2
End Sub
