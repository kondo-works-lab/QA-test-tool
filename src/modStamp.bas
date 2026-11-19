Attribute VB_Name = "modStamp"
Option Explicit

'==============================================================
' QA試験 電子印鑑（既存の QA試験.xlsm の電子印鑑機能を流用）
'   ・電子印鑑シート：B3:B10 に氏名、C列に部署。
'   ・丸印（承認者・担当者）：押すたびに苗字だけの丸印を作って表紙に押す。
'     苗字は氏名の最初の空白（全角・半角）より前。空白が無ければ氏名全体。
'   ・データ印（明細の検査者・表紙の検査ごとの担当者）：押すたびに作る。
'     上から部署／日付（押した日）／名前の3段で、段の間に線を入れる。
'   ・丸印・データ印とも同じ大きさ（STAMP_SIZE）で押す。
'   ・押す人は既存の承認者選択フォーム（frmApprover）で選ぶ。
'     frmApprover は OK で ApplyApproverStamp 氏名 を呼ぶ。
'   ・承認者印は全検査が OK のとき、担当者印は全検査の判定後に押せる。
'==============================================================

Private Const STAMP_NAME_COL As Long = 2    ' B列：氏名
Private Const STAMP_DEPT_COL As Long = 3    ' C列：部署
Private Const STAMP_FIRST_ROW As Long = 3
Private Const STAMP_LAST_ROW As Long = 10
Private Const STAMP_COLOR As Long = &HC0&   ' RGB(192, 0, 0)
Private Const STAMP_FONT_NAME As String = "メイリオ"

' 押印先（ApplyApproverStamp で使う）
Private mTargetArea As Range
Private mTargetShapeName As String
Private mInspectorSheet As Worksheet     ' 検査者印のときの明細シート（丸印のときは Nothing）
Private mPressTime As Date

'--------------------------------------------------------------
' 表紙の承認者欄に押印する（全検査が OK のときだけ）
'--------------------------------------------------------------
Sub 承認者押印()
    If Not AllInspectionsJudged() Then
        ShowMsg "判定が済んでいない検査があるため、承認者印を押せません。" & vbCrLf & _
                "先に各明細シートで「検査者押印」を押してください。", vbExclamation, "承認者押印"
        Exit Sub
    End If
    If CalcOverall() <> RESULT_OK Then
        ShowMsg "NG の検査があるため、承認者印を押せません。" & vbCrLf & _
                "承認者印は全検査が OK のときに押せます。", vbExclamation, "承認者押印"
        Exit Sub
    End If
    SetStampTarget ThisWorkbook.Worksheets(COVER_SHEET).Range(COVER_APPROVER_CELL).MergeArea, APPROVER_STAMP_NAME
    ShowStampForm
End Sub

'--------------------------------------------------------------
' 表紙の担当者欄に押印する
'--------------------------------------------------------------
Sub 担当者押印()
    If Not AllInspectionsJudged() Then
        ShowMsg "判定が済んでいない検査があるため、担当者印を押せません。" & vbCrLf & _
                "先に各明細シートで「検査者押印」を押してください。", vbExclamation, "担当者押印"
        Exit Sub
    End If
    SetStampTarget ThisWorkbook.Worksheets(COVER_SHEET).Range(COVER_CHARGE_CELL).MergeArea, CHARGE_STAMP_NAME
    ShowStampForm
End Sub

'--------------------------------------------------------------
' 丸印の押印先を設定する（テストからも使う）
'--------------------------------------------------------------
Public Sub SetStampTarget(targetArea As Range, shapeName As String)
    Set mTargetArea = targetArea
    mTargetShapeName = shapeName
    Set mInspectorSheet = Nothing
End Sub

'--------------------------------------------------------------
' 検査者印（データ印）を押す明細シートを設定する（テストからも使う）
'--------------------------------------------------------------
Public Sub BeginInspectorStamp(ws As Worksheet, pressTime As Date)
    Set mTargetArea = Nothing
    mTargetShapeName = ""
    Set mInspectorSheet = ws
    mPressTime = pressTime
End Sub

'--------------------------------------------------------------
' 承認者選択フォームを表示する
'   フォームが無くてもブック全体がコンパイルエラーにならないよう、名前で呼び出す
'--------------------------------------------------------------
Public Sub ShowStampForm()
    On Error GoTo ErrHandler
    VBA.UserForms.Add("frmApprover").Show
    Exit Sub
ErrHandler:
    ShowMsg "承認者選択フォーム（frmApprover）を表示できません。" & vbCrLf & _
            "マクロ「フォーム作成」を実行してから、もう一度押してください。" & vbCrLf & Err.Description, _
            vbExclamation, "押印"
End Sub

'--------------------------------------------------------------
' 選ばれた人の印を押印先に押す（frmApprover の OK ボタンから呼ばれる）
'   検査者印のときは明細シートの判定も行う
'--------------------------------------------------------------
Sub ApplyApproverStamp(personName As String)

    Dim ws As Worksheet

    If Not mInspectorSheet Is Nothing Then
        Set ws = mInspectorSheet
        Set mInspectorSheet = Nothing
        ApplyInspectorStamp ws, Trim(personName), mPressTime
        Exit Sub
    End If

    If mTargetArea Is Nothing Then
        ShowMsg "押印先が選ばれていません。押印ボタンからやり直してください。", vbExclamation, "押印"
        Exit Sub
    End If

    ApplyRoundStamp mTargetArea, mTargetShapeName, Trim(personName)

End Sub

'--------------------------------------------------------------
' 丸印（苗字）を押印先の中央に押す
'   既定の文字の大きさで円に収まらない長い苗字は、文字を小さくする
'--------------------------------------------------------------
Private Sub ApplyRoundStamp(targetArea As Range, shapeName As String, personName As String)

    Const BASE_FONT_SIZE As Double = 16
    Const FONT_SIZE_FLOOR As Double = 6
    Dim ws As Worksheet
    Dim surname As String
    Dim fontSize As Double

    Set ws = targetArea.Worksheet
    surname = SurnameOf(personName)
    If surname = "" Then
        ShowMsg "押す人の氏名がありません。電子印鑑シートの氏名を確認してください。", vbExclamation, "押印"
        Exit Sub
    End If

    ' 既存の印があれば削除してから押し直す
    If ShapeExists(ws, shapeName) Then ws.Shapes(shapeName).Delete

    fontSize = BASE_FONT_SIZE
    If Len(surname) > 2 Then fontSize = Int(STAMP_SIZE * 0.75 / Len(surname) * 2) / 2
    If fontSize > BASE_FONT_SIZE Then fontSize = BASE_FONT_SIZE
    If fontSize < FONT_SIZE_FLOOR Then fontSize = FONT_SIZE_FLOOR

    With ws.Shapes.AddShape(msoShapeOval, _
        targetArea.Left + (targetArea.Width - STAMP_SIZE) / 2, _
        targetArea.Top + (targetArea.Height - STAMP_SIZE) / 2, STAMP_SIZE, STAMP_SIZE)
        .Name = shapeName
        .AlternativeText = personName
        .Fill.Visible = msoFalse
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = STAMP_COLOR
        .Line.Weight = 2.25
        With .TextFrame2
            .MarginLeft = 0
            .MarginRight = 0
            .MarginTop = 0
            .MarginBottom = 0
            .WordWrap = msoFalse
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.text = surname
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .TextRange.Font.Name = STAMP_FONT_NAME
            .TextRange.Font.NameFarEast = STAMP_FONT_NAME
            .TextRange.Font.Size = fontSize
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = STAMP_COLOR
        End With
    End With

End Sub

'--------------------------------------------------------------
' データ印を押す（部署／日付／名前の3段。段の間に線）
'   印の図形はグループにして shapeName を付け、氏名を代替テキストに残す
'--------------------------------------------------------------
Public Sub DrawDataStamp(targetArea As Range, shapeName As String, dept As String, _
    stampDate As Date, personName As String)

    Dim ws As Worksheet
    Dim s As Double
    Dim x As Double
    Dim y As Double
    Dim cx As Double
    Dim halfChord As Double
    Dim partNames(0 To 5) As Variant
    Dim grp As Shape

    Set ws = targetArea.Worksheet
    If ShapeExists(ws, shapeName) Then ws.Shapes(shapeName).Delete

    s = STAMP_SIZE
    x = targetArea.Left + (targetArea.Width - s) / 2
    y = targetArea.Top + (targetArea.Height - s) / 2
    cx = x + s / 2

    ' 外枠の円
    With ws.Shapes.AddShape(msoShapeOval, x, y, s, s)
        .Fill.Visible = msoFalse
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = STAMP_COLOR
        .Line.Weight = 1.5
        .Name = shapeName & "_円"
        partNames(0) = .Name
    End With

    ' 段の間の線（円の上から 1/3 と 2/3 の高さ。円の中だけに引く）
    halfChord = Sqr((s / 2) ^ 2 - (s / 6) ^ 2)
    With ws.Shapes.AddLine(cx - halfChord, y + s / 3, cx + halfChord, y + s / 3)
        .Line.ForeColor.RGB = STAMP_COLOR
        .Line.Weight = 1
        .Name = shapeName & "_線1"
        partNames(1) = .Name
    End With
    With ws.Shapes.AddLine(cx - halfChord, y + s * 2 / 3, cx + halfChord, y + s * 2 / 3)
        .Line.ForeColor.RGB = STAMP_COLOR
        .Line.Weight = 1
        .Name = shapeName & "_線2"
        partNames(2) = .Name
    End With

    ' 文字（上：部署、中：日付、下：名前）
    partNames(3) = AddStampText(ws, shapeName & "_部署", dept, x + s * 0.15, y + s * 0.06, s * 0.7, s / 3 - s * 0.06, 8)
    partNames(4) = AddStampText(ws, shapeName & "_日付", Format(stampDate, "yyyy.mm.dd"), x + s * 0.04, y + s / 3, s * 0.92, s / 3, 8)
    partNames(5) = AddStampText(ws, shapeName & "_名前", personName, x + s * 0.15, y + s * 2 / 3, s * 0.7, s / 3 - s * 0.06, 10)

    Set grp = ws.Shapes.Range(partNames).Group
    grp.Name = shapeName
    grp.AlternativeText = personName

End Sub

'--------------------------------------------------------------
' データ印の1段分の文字。幅に収まるよう文字を小さくする。図形名を返す
'--------------------------------------------------------------
Private Function AddStampText(ws As Worksheet, shapeName As String, text As String, _
    l As Double, t As Double, w As Double, h As Double, maxFont As Double) As String

    Const FONT_FLOOR As Double = 4.5
    Dim fontSize As Double
    Dim widthUnits As Double
    Dim i As Long

    ' 1文字あたりの幅（全角1、半角0.6）から、収まるフォントの大きさを決める
    For i = 1 To Len(text)
        If Asc(Mid(text, i, 1)) < 0 Or Asc(Mid(text, i, 1)) > 255 Then
            widthUnits = widthUnits + 1.05
        Else
            widthUnits = widthUnits + 0.62
        End If
    Next i
    fontSize = maxFont
    If widthUnits > 0 Then
        If w / widthUnits < fontSize Then fontSize = Int((w / widthUnits) * 2) / 2
    End If
    If fontSize < FONT_FLOOR Then fontSize = FONT_FLOOR

    With ws.Shapes.AddTextbox(msoTextOrientationHorizontal, l, t, w, h)
        .Name = shapeName
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        With .TextFrame2
            .MarginLeft = 0
            .MarginRight = 0
            .MarginTop = 0
            .MarginBottom = 0
            .WordWrap = msoFalse
            .AutoSize = msoAutoSizeNone
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.text = text
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .TextRange.Font.Name = STAMP_FONT_NAME
            .TextRange.Font.NameFarEast = STAMP_FONT_NAME
            .TextRange.Font.Size = fontSize
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = STAMP_COLOR
        End With
        AddStampText = .Name
    End With

End Function

'--------------------------------------------------------------
' 押印されている氏名（印が無ければ空文字）
'--------------------------------------------------------------
Public Function GetStampName(ws As Worksheet, shapeName As String) As String
    If Not ShapeExists(ws, shapeName) Then Exit Function
    GetStampName = Trim(ws.Shapes(shapeName).AlternativeText)
End Function

'--------------------------------------------------------------
' 電子印鑑シートに登録された部署（見つからなければ空文字）
'--------------------------------------------------------------
Public Function GetDepartment(personName As String) As String
    Dim ws As Worksheet
    Dim r As Long
    Set ws = ThisWorkbook.Worksheets(STAMP_SHEET)
    For r = STAMP_FIRST_ROW To STAMP_LAST_ROW
        If CellText(ws.Cells(r, STAMP_NAME_COL).Value) = personName Then
            GetDepartment = CellText(ws.Cells(r, STAMP_DEPT_COL).Value)
            Exit Function
        End If
    Next r
End Function

'--------------------------------------------------------------
' 苗字（氏名の最初の空白より前。空白が無ければ氏名全体）
'--------------------------------------------------------------
Public Function SurnameOf(personName As String) As String
    Dim s As String
    Dim p As Long
    s = Trim(Replace(personName, "　", " "))
    p = InStr(s, " ")
    If p > 0 Then
        SurnameOf = Left(s, p - 1)
    Else
        SurnameOf = s
    End If
End Function


