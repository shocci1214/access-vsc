Option Compare Database
Option Explicit

' ============================================================
' UI - Access VSC のフォーム操作テスト用
'   キーボードで入力したのと同じようにフォームを操作します。
'   そのため、更新後処理（AfterUpdate）やボタンのクリック（Click）などのイベントも実際に動きます。
'
'   UI.OpenForm "frmOrder"              フォームを開く（以降の操作の対象になる）
'   UI.TypeText "txtQty", "3"           テキストボックスなどに入力して Tab で確定
'   UI.Click "cmdCalc"                  ボタンを押す（チェックボックスなら切り替え）
'   UI.PressKey UI_KEY_ENTER            キーを押す
'   UI.Value("txtTotal")                コントロールの値
'   UI.CloseForm                        フォームを閉じる（保存の確認は「いいえ」）
'
'   ・ボタンで MsgBox が出た場合は、Access VSC が自動で答えます（OK だけなら OK、確認は「キャンセル」/「いいえ」）
'   ・テストは DB のコピーで動くので、フォームから登録したデータは本物の DB には残りません
'   ・InputBox やモーダルフォーム（acDialog）は止まってしまうため、テストでは開かないようにしてください
' ============================================================

Private Declare PtrSafe Function GetFocus Lib "user32" () As LongPtr
Private Declare PtrSafe Function PostMessageW Lib "user32" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)
Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal flags As Long) As Long

Private Const WM_KEYDOWN As Long = &H100
Private Const WM_KEYUP As Long = &H101
Private Const WM_CHAR As Long = &H102
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_NOACTIVATE As Long = &H10

Public Const UI_KEY_TAB As Long = &H9
Public Const UI_KEY_ENTER As Long = &HD
Public Const UI_KEY_ESCAPE As Long = &H1B
Public Const UI_KEY_SPACE As Long = &H20
Public Const UI_KEY_DOWN As Long = &H28
Public Const UI_KEY_UP As Long = &H26
Public Const UI_KEY_F2 As Long = &H71

Private m_form As String

' 送ったキーを Access に処理させる
Private Sub Pump()
    Dim i As Long
    For i = 1 To 5
        DoEvents
        Sleep 10
    Next i
End Sub

Private Function Target(ByVal controlName As String) As Access.Control
    If Len(m_form) = 0 Then Err.Raise vbObjectError + 9101, "UI", "フォームが開かれていません。先に UI.OpenForm を呼んでください。"
    If Not CurrentProject.AllForms(m_form).IsLoaded Then Err.Raise vbObjectError + 9102, "UI", "フォーム " & m_form & " は閉じられています。"
    On Error GoTo NotFound
    Set Target = Forms(m_form).Controls(controlName)
    Exit Function
NotFound:
    Err.Raise vbObjectError + 9103, "UI", "フォーム " & m_form & " にコントロール " & controlName & " がありません。"
End Function

' フォームを開く。where は「顧客ID=3」のような条件（省略可）
Public Sub OpenForm(ByVal formName As String, Optional ByVal where As String = "")
    If Not Application.Visible Then
        ' 画面に出ていない（または最小化した）Access ではフォームにフォーカスを移せないため、
        ' ウィンドウを画面の外へ移してから表示する（テストの間、タスクバーに Access が出ます）
        SetWindowPos Application.hWndAccessApp, 0, -32000, -32000, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
        Application.Visible = True
    End If
    If Len(where) > 0 Then
        DoCmd.OpenForm formName, acNormal, , where
    Else
        DoCmd.OpenForm formName, acNormal
    End If
    m_form = formName
    Pump
End Sub

' 操作の対象にするフォームを切り替える（ボタンで別のフォームが開いたときなど）
Public Sub UseForm(ByVal formName As String)
    m_form = formName
End Sub

' 開いているフォーム（Me と同じもの）
Public Property Get Form() As Access.Form
    Set Form = Forms(m_form)
End Property

' フォームが開いているか
Public Function IsOpen(ByVal formName As String) As Boolean
    IsOpen = CurrentProject.AllForms(formName).IsLoaded
End Function

' コントロールにフォーカスを移す
Public Sub Focus(ByVal controlName As String)
    Target(controlName).SetFocus
    Pump
End Sub

' キーを 1 回押す（UI_KEY_ENTER など）
Public Sub PressKey(ByVal keyCode As Long)
    Dim h As LongPtr
    h = GetFocus()
    PostMessageW h, WM_KEYDOWN, keyCode, 1
    PostMessageW h, WM_KEYUP, keyCode, &HC0000001
    Pump
End Sub

' 入力する。今の内容は置き換える。commit = True なら最後に Tab で確定する（更新前・更新後処理が動く）
Public Sub TypeText(ByVal controlName As String, ByVal text As String, Optional ByVal commit As Boolean = True)
    Dim c As Access.Control
    Dim h As LongPtr
    Dim i As Long
    Set c = Target(controlName)
    c.SetFocus
    Pump
    On Error Resume Next
    c.SelStart = 0
    c.SelLength = Len(Nz(c.text, ""))
    On Error GoTo 0
    h = GetFocus()
    For i = 1 To Len(text)
        PostMessageW h, WM_CHAR, AscW(Mid$(text, i, 1)) And &HFFFF&, 1
    Next i
    Pump
    If commit Then PressKey UI_KEY_TAB
End Sub

' ボタンを押す（フォーカスを移してスペースキー。チェックボックス・トグルボタンは切り替わる）
Public Sub Click(ByVal controlName As String)
    Dim h As LongPtr
    Target(controlName).SetFocus
    Pump
    ' ボタンはキーを離したとき、チェックボックスは文字（スペース）で反応するので、両方を送る
    h = GetFocus()
    PostMessageW h, WM_KEYDOWN, UI_KEY_SPACE, &H390001
    PostMessageW h, WM_CHAR, UI_KEY_SPACE, &H390001
    PostMessageW h, WM_KEYUP, UI_KEY_SPACE, &HC0390001
    Pump
End Sub

' コントロールの値
Public Function Value(ByVal controlName As String) As Variant
    Value = Target(controlName).Value
End Function

' フォームを閉じる（省略すると今のフォーム）。レコードの変更は保存、フォームのデザインは保存しない
Public Sub CloseForm(Optional ByVal formName As String = "")
    If Len(formName) = 0 Then formName = m_form
    If Len(formName) = 0 Then Exit Sub
    If CurrentProject.AllForms(formName).IsLoaded Then DoCmd.Close acForm, formName, acSaveNo
    If formName = m_form Then m_form = vbNullString
    Pump
End Sub
