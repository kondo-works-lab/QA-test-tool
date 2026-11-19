Attribute VB_Name = "modForms"
Option Explicit

'==============================================================
' QA試験 フォーム作成
'   「フォーム作成」を1回実行すると、次の2つのユーザーフォームを作る。
'     frmApprover ：押す人を選ぶ（電子印鑑シートの氏名の一覧）。OK で ApplyApproverStamp を呼ぶ
'     frmSaveType ：保存形式を選ぶ（Excel／PDF のチェックボックス）。OK で SetSaveTypes を呼ぶ
'   ・実行には Excel の設定「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」が必要
'     （作った後は元に戻してよい）。
'   ・同じ名前のフォームが既にあれば作らない（足りないフォームだけ作る）。
'==============================================================

Private Const VBEXT_CT_MSFORM As Long = 3   ' vbext_ct_MSForm

Sub フォーム作成()

    Dim vbp As Object
    Dim created As String
    Dim skipped As String
    Dim errMsg As String

    On Error Resume Next
    Set vbp = ThisWorkbook.VBProject
    If Not vbp Is Nothing Then
        If vbp.VBComponents.Count = 0 Then Set vbp = Nothing
    End If
    errMsg = Err.Description
    On Error GoTo 0
    If vbp Is Nothing Then
        ShowMsg "フォームを作れません。" & vbCrLf & _
                "［ファイル］→［オプション］→［トラスト センター］→［トラスト センターの設定］→［マクロの設定］で、" & vbCrLf & _
                "「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」にチェックを入れてから、もう一度実行してください。" & vbCrLf & errMsg, _
                vbExclamation, "フォーム作成"
        Exit Sub
    End If

    On Error GoTo ErrHandler
    If ComponentExists(vbp, "frmApprover") Then
        skipped = skipped & "frmApprover "
    Else
        CreateApproverForm vbp
        created = created & "frmApprover "
    End If
    If ComponentExists(vbp, "frmSaveType") Then
        skipped = skipped & "frmSaveType "
    Else
        CreateSaveTypeForm vbp
        created = created & "frmSaveType "
    End If

    ShowMsg "作成したフォーム：" & IIf(created = "", "なし", created) & vbCrLf & _
            "既にあったフォーム：" & IIf(skipped = "", "なし", skipped), vbInformation, "フォーム作成"
    Exit Sub

ErrHandler:
    ShowMsg "フォームの作成に失敗しました。" & vbCrLf & Err.Description, vbExclamation, "フォーム作成"

End Sub

Private Function ComponentExists(vbp As Object, compName As String) As Boolean
    Dim comp As Object
    For Each comp In vbp.VBComponents
        If comp.Name = compName Then
            ComponentExists = True
            Exit Function
        End If
    Next comp
End Function

'--------------------------------------------------------------
' 押す人を選ぶフォーム
'--------------------------------------------------------------
Private Sub CreateApproverForm(vbp As Object)

    Dim comp As Object
    Dim code As String

    Set comp = vbp.VBComponents.Add(VBEXT_CT_MSFORM)
    comp.Name = "frmApprover"
    comp.Properties("Caption") = "押す人を選択"
    comp.Properties("Width") = 220
    comp.Properties("Height") = 230

    AddControl comp, "Forms.Label.1", "lblTitle", 12, 8, 190, 14, "押す人を選んでください"
    AddControl comp, "Forms.ListBox.1", "lstNames", 12, 26, 190, 130, ""
    AddControl comp, "Forms.CommandButton.1", "cmdOK", 30, 166, 70, 24, "OK"
    AddControl comp, "Forms.CommandButton.1", "cmdCancel", 112, 166, 70, 24, "キャンセル"

    code = _
        "Option Explicit" & vbCrLf & vbCrLf & _
        "' 電子印鑑シートの B3 から下の氏名を一覧に出す" & vbCrLf & _
        "Private Sub UserForm_Initialize()" & vbCrLf & _
        "    Dim wsStamp As Worksheet" & vbCrLf & _
        "    Dim r As Long" & vbCrLf & _
        "    Set wsStamp = ThisWorkbook.Worksheets(STAMP_SHEET)" & vbCrLf & _
        "    Me.lstNames.Clear" & vbCrLf & _
        "    r = 3" & vbCrLf & _
        "    Do While CellText(wsStamp.Cells(r, 2).Value) <> """"" & vbCrLf & _
        "        Me.lstNames.AddItem CellText(wsStamp.Cells(r, 2).Value)" & vbCrLf & _
        "        r = r + 1" & vbCrLf & _
        "    Loop" & vbCrLf & _
        "End Sub" & vbCrLf & vbCrLf & _
        "Private Sub cmdOK_Click()" & vbCrLf & _
        "    If Me.lstNames.ListIndex = -1 Then" & vbCrLf & _
        "        MsgBox ""担当者を選択してください"", vbExclamation" & vbCrLf & _
        "        Exit Sub" & vbCrLf & _
        "    End If" & vbCrLf & _
        "    ApplyApproverStamp Me.lstNames.Value" & vbCrLf & _
        "    Unload Me" & vbCrLf & _
        "End Sub" & vbCrLf & vbCrLf & _
        "Private Sub cmdCancel_Click()" & vbCrLf & _
        "    Unload Me" & vbCrLf & _
        "End Sub" & vbCrLf

    SetFormCode comp, code

End Sub

'--------------------------------------------------------------
' 保存形式を選ぶフォーム
'--------------------------------------------------------------
Private Sub CreateSaveTypeForm(vbp As Object)

    Dim comp As Object
    Dim code As String

    Set comp = vbp.VBComponents.Add(VBEXT_CT_MSFORM)
    comp.Name = "frmSaveType"
    comp.Properties("Caption") = "保存"
    comp.Properties("Width") = 220
    comp.Properties("Height") = 150

    AddControl comp, "Forms.Label.1", "lblTitle", 12, 8, 190, 14, "保存する形式を選んでください（両方も可）"
    AddControl comp, "Forms.CheckBox.1", "chkExcel", 24, 30, 160, 18, "Excel（.xlsx）"
    AddControl comp, "Forms.CheckBox.1", "chkPdf", 24, 52, 160, 18, "PDF"
    AddControl comp, "Forms.CommandButton.1", "cmdOK", 30, 84, 70, 24, "OK"
    AddControl comp, "Forms.CommandButton.1", "cmdCancel", 112, 84, 70, 24, "キャンセル"

    code = _
        "Option Explicit" & vbCrLf & vbCrLf & _
        "Private Sub cmdOK_Click()" & vbCrLf & _
        "    If Not Me.chkExcel.Value And Not Me.chkPdf.Value Then" & vbCrLf & _
        "        MsgBox ""Excel か PDF の少なくとも1つを選んでください"", vbExclamation" & vbCrLf & _
        "        Exit Sub" & vbCrLf & _
        "    End If" & vbCrLf & _
        "    SetSaveTypes CBool(Me.chkExcel.Value), CBool(Me.chkPdf.Value)" & vbCrLf & _
        "    Unload Me" & vbCrLf & _
        "End Sub" & vbCrLf & vbCrLf & _
        "Private Sub cmdCancel_Click()" & vbCrLf & _
        "    Unload Me" & vbCrLf & _
        "End Sub" & vbCrLf

    SetFormCode comp, code

End Sub

'--------------------------------------------------------------
' フォームに部品を置く
'--------------------------------------------------------------
Private Sub AddControl(comp As Object, progId As String, ctlName As String, _
    l As Double, t As Double, w As Double, h As Double, caption As String)

    Dim ctl As Object

    Set ctl = comp.Designer.Controls.Add(progId, ctlName, True)
    ctl.Left = l
    ctl.Top = t
    ctl.Width = w
    ctl.Height = h
    ctl.Font.Name = "メイリオ"
    ctl.Font.Size = 9
    If caption <> "" Then ctl.caption = caption

End Sub

'--------------------------------------------------------------
' フォームのコードを入れる（自動で入る Option Explicit と重ならないよう、先に消す）
'--------------------------------------------------------------
Private Sub SetFormCode(comp As Object, code As String)
    With comp.CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        .AddFromString code
    End With
End Sub
