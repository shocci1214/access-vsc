Attribute VB_Name = "modSample"
Option Compare Database
Option Explicit

Private Const RATE As Double = 0.1   

' 税込金額
Public Function TaxIncluded(ByVal price As Currency) As Currency
TaxIncluded = Int(price * (1 + RATE))
End Function

Public Sub Sample(ByVal n As Long)
Dim i As Long, s As String
    On Error GoTo ErrHandler
If n > 0 Then ' 正の数
For i = 1 To n
s = s & "If x Then"
Next i
ElseIf n = 0 Then
Debug.Print "zero"
Else
Select Case n
Case -1
Debug.Print "minus one"
Case Else
Debug.Print "other"
End Select
End If
If n = 5 Then Exit Sub
With CurrentDb
Do While n < 10
n = n + 1
Loop
End With
s = "a" & _
      "b" & _
      "c"
#If VBA7 Then
  Debug.Print "vba7"
#End If
    ExitProc:
Exit Sub
ErrHandler:
Resume ExitProc
End Sub

Private Type Pair
Key As String
Value As Long
End Type
