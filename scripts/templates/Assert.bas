Option Compare Database
Option Explicit

' ============================================================
' Assert - Access VSC のテスト用アサーション
'   テストの中で  Assert.AreEqual 期待値, 実際の値  のように使います。
'   条件を満たさないと、エラー番号 vbObjectError + 9001 / Source "Assert" の
'   エラーを発生させ、テストは「失敗」として報告されます。
' ============================================================

Private Const ASSERT_ERROR As Long = vbObjectError + 9001

Private Sub Raise(ByVal detail As String, ByVal message As String)
    If Len(message) > 0 Then detail = message & " - " & detail
    Err.Raise ASSERT_ERROR, "Assert", detail
End Sub

Private Function Show(ByVal v As Variant) As String
    If IsObject(v) Then
        If v Is Nothing Then Show = "Nothing" Else Show = "<" & TypeName(v) & ">"
    ElseIf IsNull(v) Then
        Show = "Null"
    ElseIf IsEmpty(v) Then
        Show = "Empty"
    ElseIf IsArray(v) Then
        Show = "<配列>"
    ElseIf VarType(v) = vbString Then
        Show = """" & v & """"
    Else
        Show = CStr(v) & " (" & TypeName(v) & ")"
    End If
End Function

Private Function Same(ByVal a As Variant, ByVal b As Variant) As Boolean
    If IsObject(a) Or IsObject(b) Then
        If IsObject(a) And IsObject(b) Then Same = (a Is b)
    ElseIf IsNull(a) Or IsNull(b) Then
        Same = IsNull(a) And IsNull(b)
    ElseIf IsArray(a) Or IsArray(b) Then
        Same = False
    ElseIf VarType(a) = vbString Xor VarType(b) = vbString Then
        Same = False   ' "1" と 1 は別物として扱う
    Else
        Same = (a = b)
    End If
End Function

' 等しいこと
Public Sub AreEqual(ByVal expected As Variant, ByVal actual As Variant, Optional ByVal message As String = "")
    If Not Same(expected, actual) Then Raise "期待値 " & Show(expected) & " / 実際 " & Show(actual), message
End Sub

' 等しくないこと
Public Sub AreNotEqual(ByVal notExpected As Variant, ByVal actual As Variant, Optional ByVal message As String = "")
    If Same(notExpected, actual) Then Raise Show(actual) & " ではないはず", message
End Sub

' 数値が誤差 tolerance 以内で等しいこと（小数の計算に）
Public Sub AreClose(ByVal expected As Double, ByVal actual As Double, Optional ByVal tolerance As Double = 0.000001, Optional ByVal message As String = "")
    If Abs(expected - actual) > tolerance Then Raise "期待値 " & expected & " ± " & tolerance & " / 実際 " & actual, message
End Sub

Public Sub IsTrue(ByVal condition As Boolean, Optional ByVal message As String = "")
    If Not condition Then Raise "True のはずが False", message
End Sub

Public Sub IsFalse(ByVal condition As Boolean, Optional ByVal message As String = "")
    If condition Then Raise "False のはずが True", message
End Sub

Public Sub IsNullValue(ByVal actual As Variant, Optional ByVal message As String = "")
    If Not IsNull(actual) Then Raise "Null のはずが " & Show(actual), message
End Sub

Public Sub IsNotNullValue(ByVal actual As Variant, Optional ByVal message As String = "")
    If IsNull(actual) Then Raise "Null ではないはず", message
End Sub

Public Sub IsNothing(ByVal actual As Object, Optional ByVal message As String = "")
    If Not actual Is Nothing Then Raise "Nothing のはずが <" & TypeName(actual) & ">", message
End Sub

Public Sub IsNotNothing(ByVal actual As Object, Optional ByVal message As String = "")
    If actual Is Nothing Then Raise "Nothing ではないはず", message
End Sub

' 文字列 text に part が含まれること
Public Sub Contains(ByVal text As String, ByVal part As String, Optional ByVal message As String = "")
    If InStr(1, text, part, vbBinaryCompare) = 0 Then Raise """" & text & """ に """ & part & """ が含まれていない", message
End Sub

' 無条件に失敗させる
Public Sub Fail(Optional ByVal message As String = "失敗")
    Err.Raise ASSERT_ERROR, "Assert", message
End Sub
