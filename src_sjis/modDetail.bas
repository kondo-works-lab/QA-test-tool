Attribute VB_Name = "modDetail"
Option Explicit

'==============================================================
' QA試験 明細シート（テスト1つに1シート）
'   ・表紙の品名とロット番号が入ったら、マスタからその品名の検査分だけ
'     明細シートを作る。形は既存の QA試験.xlsm の検査表1つ分
'     （基準値／No.／検査項目／OK欄。NG欄は無し）。
'   ・OK欄は入力規則で「OK」か「NG」だけ入力できる。
'   ・「検査者押印」ボタンで次を行う。
'     ① OK欄の確認（空欄・OK／NG以外があればメッセージを出して止める）
'        検査値をマスタの基準で判定し、OK欄と食い違う行や判定できない値があれば止める
'     ② 押す人を選び、検査者欄にデータ印（部署・日付・名前）を押す
'     ③ OK欄がすべて OK なら総合判定に「OK」、1つでも NG なら「NG」を入れる
'     ④ 表紙の該当検査の判定結果と担当者印（同じデータ印）を入れ、表紙の総合判定を更新する
'   ・「明細削除」ボタンで、そのシートの入力（検査項目・OK欄）と検査者印・総合判定、
'     表紙のその検査の判定結果・担当者印と、表紙の承認者印・担当者印を消す（消す前に確認する）。
'==============================================================

Public Const DETAIL_STAMP_NAME As String = "DetailStamp_検査者"
Private Const STAMP_BUTTON_NAME As String = "btnInspectorStamp"
Private Const CLEAR_BUTTON_NAME As String = "btnDetailClear"
Private Const STAMP_AREA_ROWS As Long = 4      ' 検査者欄・総合判定欄の行数

'--------------------------------------------------------------
' 明細シートをすべて削除する。削除前に入力・判定があったかを返す
'--------------------------------------------------------------
Public Function DeleteDetailSheets() As Boolean

    Dim item As Variant
    Dim ws As Worksheet
    Dim hadData As Boolean
    Dim alertsBefore As Boolean

    alertsBefore = Application.DisplayAlerts
    Application.DisplayAlerts = False
    For Each item In GetDetailSheets()
        Set ws = item
        If HasAnyInput(ws) Or IsJudged(ws) Then hadData = True
        ws.Delete
    Next item
    Application.DisplayAlerts = alertsBefore

    DeleteDetailSheets = hadData

End Function

'--------------------------------------------------------------
' 品名の検査分だけ明細シートを作る（表紙の後ろに、マスタの順で並べる）
'   マスタはチェック済みであること
'--------------------------------------------------------------
Public Sub CreateDetailSheets(productName As String, lotNumber As String)

    Dim inspName As Variant
    Dim ws As Worksheet
    Dim afterSheet As Worksheet
    Dim no As Long

    Set afterSheet = ThisWorkbook.Worksheets(COVER_SHEET)
    For Each inspName In GetInspectionNames(productName)
        no = no + 1
        Set ws = AddSheetAfter(afterSheet)
        ws.Name = DETAIL_PREFIX & inspName
        BuildDetailSheet ws, productName, lotNumber, CStr(inspName), no
        Set afterSheet = ws
    Next inspName

    ThisWorkbook.Worksheets(COVER_SHEET).Activate

End Sub

'--------------------------------------------------------------
' 明細シート1枚分のレイアウトを作る
'--------------------------------------------------------------
Private Sub BuildDetailSheet(ws As Worksheet, productName As String, lotNumber As String, _
    inspName As String, no As Long)

    Dim ms As Worksheet
    Dim mr As Long
    Dim n As Long
    Dim r As Long
    Dim lastRow As Long
    Dim stampArea As Range
    Dim resultArea As Range

    Set ms = ThisWorkbook.Worksheets(MASTER_SHEET)
    mr = GetMasterRow(productName, inspName)
    n = GetCountAt(ms, mr)
    lastRow = DETAIL_FIRST_ROW + n - 1

    ws.Cells.Font.Name = "メイリオ"
    ws.Cells.Font.Size = 10

    ' 列幅
    ws.Columns(1).ColumnWidth = 2
    ws.Columns(DETAIL_COL_STD).ColumnWidth = 24
    ws.Columns(DETAIL_COL_NO).ColumnWidth = 6
    ws.Columns(DETAIL_COL_VALUE).ColumnWidth = 14
    ws.Columns(DETAIL_COL_OK).ColumnWidth = 10
    ws.Columns(6).ColumnWidth = 3
    ws.Range(DETAIL_STAMP_CELL).EntireColumn.ColumnWidth = 12
    ws.Range(DETAIL_RESULT_CELL).EntireColumn.ColumnWidth = 12

    ' 見出し（例：１．耐熱テスト）・品名・ロット番号
    With ws.Range(DETAIL_TITLE_CELL)
        .Value = StrConv(CStr(no), vbWide) & "．" & inspName
        .Font.Size = 14
        .Font.Bold = True
    End With
    ws.Range(DETAIL_PRODUCT_CELL).Offset(0, -1).Value = "品名："
    ws.Range(DETAIL_LOT_CELL).Offset(0, -1).Value = "ロット番号："
    ws.Range(DETAIL_PRODUCT_CELL).Resize(1, 2).Merge
    ws.Range(DETAIL_LOT_CELL).Resize(1, 2).Merge
    ws.Range(DETAIL_PRODUCT_CELL).NumberFormat = "@"
    ws.Range(DETAIL_LOT_CELL).NumberFormat = "@"
    ws.Range(DETAIL_PRODUCT_CELL).Value = productName
    ws.Range(DETAIL_LOT_CELL).Value = lotNumber
    ws.Range(DETAIL_PRODUCT_CELL).Resize(2, 2).HorizontalAlignment = xlLeft

    ' 表の見出し
    ws.Cells(DETAIL_HEADER_ROW, DETAIL_COL_STD).Value = "基準値"
    ws.Cells(DETAIL_HEADER_ROW, DETAIL_COL_NO).Value = "No."
    ws.Cells(DETAIL_HEADER_ROW, DETAIL_COL_VALUE).Value = "検査項目"
    ws.Cells(DETAIL_HEADER_ROW, DETAIL_COL_OK).Value = "OK欄"
    With ws.Range(ws.Cells(DETAIL_HEADER_ROW, DETAIL_COL_STD), ws.Cells(DETAIL_HEADER_ROW, DETAIL_COL_OK))
        .Interior.Color = COLOR_HEADER
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    ' 項目（No. 1～数）
    For r = DETAIL_FIRST_ROW To lastRow
        ws.Cells(r, DETAIL_COL_NO).Value = r - DETAIL_FIRST_ROW + 1
    Next r
    With ws.Range(ws.Cells(DETAIL_FIRST_ROW, DETAIL_COL_STD), ws.Cells(lastRow, DETAIL_COL_STD))
        .Merge
        .Value = StandardTextAt(ms, mr)
        .WrapText = True
        .VerticalAlignment = xlTop
        .HorizontalAlignment = xlLeft
    End With
    ws.Range(ws.Cells(DETAIL_FIRST_ROW, DETAIL_COL_NO), ws.Cells(lastRow, DETAIL_COL_OK)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(DETAIL_FIRST_ROW, DETAIL_COL_VALUE), ws.Cells(lastRow, DETAIL_COL_OK)).Interior.Color = COLOR_INPUT

    ' OK欄：OK か NG だけ入力できるようにする
    With ws.Range(ws.Cells(DETAIL_FIRST_ROW, DETAIL_COL_OK), ws.Cells(lastRow, DETAIL_COL_OK)).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=RESULT_OK & "," & RESULT_NG
        .IgnoreBlank = True
        .InCellDropdown = True
        .ErrorTitle = "OK欄"
        .ErrorMessage = "「OK」か「NG」を選んでください。"
    End With
    SetBorders ws.Range(ws.Cells(DETAIL_HEADER_ROW, DETAIL_COL_STD), ws.Cells(lastRow, DETAIL_COL_OK))

    ' 検査者欄・総合判定（データ印が収まる大きさにする）
    Set stampArea = ws.Range(DETAIL_STAMP_CELL).Resize(STAMP_AREA_ROWS, 1)
    Set resultArea = ws.Range(DETAIL_RESULT_CELL).Resize(STAMP_AREA_ROWS, 1)
    stampArea.Cells(1, 1).Offset(-1, 0).Value = "検査者"
    resultArea.Cells(1, 1).Offset(-1, 0).Value = "総合判定"
    With Union(stampArea.Cells(1, 1).Offset(-1, 0), resultArea.Cells(1, 1).Offset(-1, 0))
        .Interior.Color = COLOR_HEADER
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    stampArea.Merge
    resultArea.Merge
    With resultArea
        .Font.Size = 16
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    EnsureStampArea stampArea
    SetBorders Union(stampArea.Offset(-1, 0).Resize(STAMP_AREA_ROWS + 1, 1), resultArea.Offset(-1, 0).Resize(STAMP_AREA_ROWS + 1, 1))

    AddButton ws, STAMP_BUTTON_NAME, "検査者押印", "検査者押印", _
        stampArea.Left, ws.Range(DETAIL_TITLE_CELL).Top, stampArea.Width, 28
    AddButton ws, CLEAR_BUTTON_NAME, "明細削除", "明細削除", _
        resultArea.Left, ws.Range(DETAIL_TITLE_CELL).Top, resultArea.Width, 28

End Sub

Private Sub SetBorders(Target As Range)
    With Target.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
End Sub

'--------------------------------------------------------------
' 明細シートの数（No. が入っている行の数）
'--------------------------------------------------------------
Public Function DetailItemCount(ws As Worksheet) As Long
    Dim r As Long
    r = DETAIL_FIRST_ROW
    Do While IsNumericValue(ws.Cells(r, DETAIL_COL_NO).Value)
        r = r + 1
    Loop
    DetailItemCount = r - DETAIL_FIRST_ROW
End Function

'--------------------------------------------------------------
' 入力欄（検査項目・OK欄）。項目が無ければ Nothing
'--------------------------------------------------------------
Public Function DetailInputArea(ws As Worksheet) As Range
    Dim cnt As Long
    cnt = DetailItemCount(ws)
    If cnt = 0 Then
        Set DetailInputArea = Nothing
    Else
        Set DetailInputArea = ws.Range(ws.Cells(DETAIL_FIRST_ROW, DETAIL_COL_VALUE), _
                                       ws.Cells(DETAIL_FIRST_ROW + cnt - 1, DETAIL_COL_OK))
    End If
End Function

'--------------------------------------------------------------
' 入力欄に1つでも入力があるか
'--------------------------------------------------------------
Public Function HasAnyInput(ws As Worksheet) As Boolean
    Dim area As Range
    Dim cell As Range
    Set area = DetailInputArea(ws)
    If area Is Nothing Then Exit Function
    For Each cell In area.Cells
        If Not IsBlankValue(cell.Value) Then
            HasAnyInput = True
            Exit Function
        End If
    Next cell
End Function

'--------------------------------------------------------------
' 判定済みかどうか（総合判定に結果がある）
'--------------------------------------------------------------
Public Function IsJudged(ws As Worksheet) As Boolean
    IsJudged = (GetDetailResult(ws) <> "")
End Function

Public Function GetDetailResult(ws As Worksheet) As String
    GetDetailResult = CellText(ws.Range(DETAIL_RESULT_CELL).Value)
End Function

'--------------------------------------------------------------
' OK欄の NG の数
'--------------------------------------------------------------
Public Function CountNG(ws As Worksheet) As Long
    Dim r As Long
    For r = DETAIL_FIRST_ROW To DETAIL_FIRST_ROW + DetailItemCount(ws) - 1
        If NormalizeOkNg(ws.Cells(r, DETAIL_COL_OK).Value) = RESULT_NG Then CountNG = CountNG + 1
    Next r
End Function

'--------------------------------------------------------------
' OK欄がすべて OK か NG になっているか。問題があれば errText に内容を入れて False
'--------------------------------------------------------------
Public Function CheckOkColumn(ws As Worksheet, ByRef errText As String) As Boolean

    Dim r As Long
    Dim cell As Range
    Dim blankList As String
    Dim wrongList As String

    errText = ""
    If DetailItemCount(ws) = 0 Then
        errText = "・判定する項目がありません。表紙の品名を確認してください。" & vbCrLf
        Exit Function
    End If

    For r = DETAIL_FIRST_ROW To DETAIL_FIRST_ROW + DetailItemCount(ws) - 1
        Set cell = ws.Cells(r, DETAIL_COL_OK)
        If IsBlankValue(cell.Value) Then
            blankList = AppendAddress(blankList, cell)
        ElseIf NormalizeOkNg(cell.Value) = "" Then
            wrongList = AppendAddress(wrongList, cell)
        End If
    Next r

    If blankList <> "" Then errText = errText & "・OK欄が空欄です：" & blankList & vbCrLf
    If wrongList <> "" Then errText = errText & "・OK欄に OK／NG 以外が入っています：" & wrongList & vbCrLf
    CheckOkColumn = (errText = "")

End Function

'--------------------------------------------------------------
' 検査値を基準（マスタのルール・基準値）で判定し、OK欄と合っているか確かめる。
' 判定できない値や、OK欄と食い違う行があれば errText に内容を入れて False
'   ・以上／以下／範囲：検査値は数値（全角数字も可）。空欄・数値でなければ判定できない
'   ・一致：大文字小文字・全角半角・前後の空白を区別しない。基準値1に「C,F」のように
'     複数あるときは、どれかに一致すれば合格（区切りは , ， 、）
'--------------------------------------------------------------
Public Function CheckAgainstStandard(ws As Worksheet, ByRef errText As String) As Boolean

    Dim ms As Worksheet
    Dim mr As Long
    Dim productName As String
    Dim inspName As String
    Dim method As String
    Dim std1 As String
    Dim std2 As String
    Dim r As Long
    Dim result As String
    Dim reason As String
    Dim okText As String
    Dim badList As String
    Dim mismatchList As String

    errText = ""
    productName = CellText(ThisWorkbook.Worksheets(COVER_SHEET).Range(COVER_PRODUCT_CELL).Value)
    inspName = InspectionNameOf(ws)
    mr = GetMasterRow(productName, inspName)
    If mr = 0 Then
        errText = "・マスタに「" & productName & "」の「" & inspName & "」がありません。" & vbCrLf
        Exit Function
    End If

    Set ms = ThisWorkbook.Worksheets(MASTER_SHEET)
    method = CellText(ms.Cells(mr, MCOL_METHOD).Value)
    std1 = CellText(ms.Cells(mr, MCOL_STD1).Value)
    std2 = CellText(ms.Cells(mr, MCOL_STD2).Value)

    For r = DETAIL_FIRST_ROW To DETAIL_FIRST_ROW + DetailItemCount(ws) - 1
        If Not JudgeValue(method, std1, std2, ws.Cells(r, DETAIL_COL_VALUE).Value, result, reason) Then
            badList = badList & "　" & ws.Cells(r, DETAIL_COL_VALUE).Address(False, False) & "：" & reason & vbCrLf
        Else
            okText = NormalizeOkNg(ws.Cells(r, DETAIL_COL_OK).Value)
            If okText <> result Then
                mismatchList = mismatchList & "　" & ws.Cells(r, DETAIL_COL_OK).Address(False, False) & _
                    "：検査値「" & CellText(ws.Cells(r, DETAIL_COL_VALUE).Value) & "」は基準では " & result & _
                    "、OK欄は " & okText & vbCrLf
            End If
        End If
    Next r

    If badList <> "" Then errText = errText & "・基準で判定できない検査項目があります" & vbCrLf & badList
    If mismatchList <> "" Then errText = errText & "・OK欄が基準と合いません（基準：" & _
        FormatStandard(method, std1, std2, "") & "）" & vbCrLf & mismatchList
    CheckAgainstStandard = (errText = "")

End Function

'--------------------------------------------------------------
' 1つの検査値を判定する。判定できたら True（result に OK／NG）、できなければ False（reason に理由）
'--------------------------------------------------------------
Public Function JudgeValue(method As String, std1 As String, std2 As String, cellValue As Variant, _
    ByRef result As String, ByRef reason As String) As Boolean

    Dim s As String
    Dim v As Double
    Dim lo As Double
    Dim hi As Double
    Dim item As Variant
    Dim passed As Boolean

    result = ""
    reason = ""
    If IsError(cellValue) Then
        reason = "エラー値です"
        Exit Function
    End If
    s = Trim(StrConv(CellText(cellValue), vbNarrow))
    If s = "" Then
        reason = "空欄です"
        Exit Function
    End If

    Select Case method
        Case METHOD_MIN, METHOD_MAX, METHOD_RANGE
            If Not IsNumeric(s) Then
                reason = "数値ではありません（" & CellText(cellValue) & "）"
                Exit Function
            End If
            v = CDbl(s)
            If Not ToNumber(std1, lo) Then
                reason = "マスタの基準値1が数値ではありません"
                Exit Function
            End If
            Select Case method
                Case METHOD_MIN: passed = (v >= lo)
                Case METHOD_MAX: passed = (v <= lo)
                Case METHOD_RANGE
                    If Not ToNumber(std2, hi) Then
                        reason = "マスタの基準値2が数値ではありません"
                        Exit Function
                    End If
                    passed = (v >= lo And v <= hi)
            End Select
        Case METHOD_MATCH
            For Each item In Split(Replace(Replace(std1, "，", ","), "、", ","), ",")
                If NormalizeMatchText(CStr(item)) = NormalizeMatchText(s) And Trim(CStr(item)) <> "" Then
                    passed = True
                    Exit For
                End If
            Next item
        Case Else
            reason = "マスタのルール「" & method & "」で判定できません"
            Exit Function
    End Select

    result = IIf(passed, RESULT_OK, RESULT_NG)
    JudgeValue = True

End Function

Private Function ToNumber(text As String, ByRef n As Double) As Boolean
    Dim s As String
    s = Trim(StrConv(text, vbNarrow))
    If s = "" Or Not IsNumeric(s) Then Exit Function
    n = CDbl(s)
    ToNumber = True
End Function

' 一致の比較用：大文字小文字・全角半角・前後の空白をそろえる
Private Function NormalizeMatchText(text As String) As String
    NormalizeMatchText = UCase(Trim(StrConv(text, vbNarrow)))
End Function

'--------------------------------------------------------------
' 判定結果（検査者印・総合判定）を消す。何か消したら True
'--------------------------------------------------------------
Public Function ClearDetailJudgement(ws As Worksheet) As Boolean

    If IsJudged(ws) Then ClearDetailJudgement = True
    ws.Range(DETAIL_RESULT_CELL).MergeArea.ClearContents

    If ShapeExists(ws, DETAIL_STAMP_NAME) Then
        ws.Shapes(DETAIL_STAMP_NAME).Delete
        ClearDetailJudgement = True
    End If

End Function

'--------------------------------------------------------------
' ロット番号の表示を更新する（品名が同じでロットだけ変わったとき）
'--------------------------------------------------------------
Public Sub UpdateDetailLot(lotNumber As String)
    Dim item As Variant
    For Each item In GetDetailSheets()
        item.Range(DETAIL_LOT_CELL).Value = lotNumber
    Next item
End Sub

'--------------------------------------------------------------
' 「検査者押印」ボタン：OK欄と基準との照合を確認してから、押す人を選ぶ
'--------------------------------------------------------------
Sub 検査者押印()

    Dim pressTime As Date
    Dim errText As String

    pressTime = Now   ' ボタン押下時刻（データ印の日付に使う）

    If Not IsDetailSheet(ActiveSheet) Then
        ShowMsg "明細シートで「検査者押印」を押してください。", vbExclamation, "検査者押印"
        Exit Sub
    End If
    If Not CheckOkColumn(ActiveSheet, errText) Then
        ShowMsg errText & vbCrLf & "修正してから、もう一度「検査者押印」を押してください。", vbExclamation, "OK欄の確認"
        Exit Sub
    End If
    If Not CheckAgainstStandard(ActiveSheet, errText) Then
        ShowMsg errText & vbCrLf & "修正してから、もう一度「検査者押印」を押してください。", vbExclamation, "基準との照合"
        Exit Sub
    End If

    BeginInspectorStamp ActiveSheet, pressTime
    ShowStampForm

End Sub

'--------------------------------------------------------------
' 検査者印を押して判定する（押す人を選んだ後に呼ばれる）。できたら True
'--------------------------------------------------------------
Public Function ApplyInspectorStamp(ws As Worksheet, personName As String, pressTime As Date) As Boolean

    Dim cs As Worksheet
    Dim errText As String
    Dim result As String
    Dim dept As String
    Dim inspName As String
    Dim r As Long
    Dim eventsBefore As Boolean
    Dim errMsg As String

    On Error GoTo ErrHandler
    eventsBefore = Application.EnableEvents

    If Not CheckOkColumn(ws, errText) Then
        ShowMsg errText & vbCrLf & "修正してから、もう一度「検査者押印」を押してください。", vbExclamation, "OK欄の確認"
        Exit Function
    End If
    ' 検査値を基準で判定し、人が入れた OK欄と食い違えば押さない
    If Not CheckAgainstStandard(ws, errText) Then
        ShowMsg errText & vbCrLf & "修正してから、もう一度「検査者押印」を押してください。", vbExclamation, "基準との照合"
        Exit Function
    End If

    If CountNG(ws) > 0 Then
        result = RESULT_NG
    Else
        result = RESULT_OK
    End If
    dept = GetDepartment(personName)
    inspName = InspectionNameOf(ws)
    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)

    Application.EnableEvents = False

    ' ② 検査者欄にデータ印　③ 総合判定
    DrawDataStamp ws.Range(DETAIL_STAMP_CELL).MergeArea, DETAIL_STAMP_NAME, dept, pressTime, personName
    ws.Range(DETAIL_RESULT_CELL).Value = result
    SetResultFont ws.Range(DETAIL_RESULT_CELL).MergeArea, result

    ' ④ 表紙の判定結果・担当者印・総合判定
    SetCoverResult inspName, result
    r = FindCoverRow(inspName)
    If r > 0 Then DrawDataStamp cs.Cells(r, COVER_COL_STAMP), INSP_STAMP_PREFIX & inspName, dept, pressTime, personName
    UpdateOverall

    Application.EnableEvents = eventsBefore

    ws.Activate
    ApplyInspectorStamp = True
    Exit Function

ErrHandler:
    errMsg = Err.Description
    Application.EnableEvents = eventsBefore
    ShowMsg "検査者印の押印中にエラーが発生しました。" & vbCrLf & errMsg, vbExclamation, "検査者押印"

End Function

'--------------------------------------------------------------
' NG は赤字にする
'--------------------------------------------------------------
Public Sub SetResultFont(cell As Range, result As String)
    If result = RESULT_NG Then
        cell.Font.Color = RGB(192, 0, 0)
    Else
        cell.Font.Color = RGB(0, 0, 0)
    End If
End Sub

'--------------------------------------------------------------
' 「明細削除」ボタン：このシートの入力と印鑑を消す
'--------------------------------------------------------------
Sub 明細削除()
    If Not IsDetailSheet(ActiveSheet) Then
        ShowMsg "明細シートで「明細削除」を押してください。", vbExclamation, "明細削除"
        Exit Sub
    End If
    If ShowMsg("「" & InspectionNameOf(ActiveSheet) & "」の入力（検査項目・OK欄）と印鑑を削除します。" & vbCrLf & _
               "表紙のこの検査の判定結果・担当者印と、承認者印・担当者印も消えます。よろしいですか？", _
               vbYesNo + vbExclamation + vbDefaultButton2, "明細削除") = vbNo Then Exit Sub
    ClearDetailData ActiveSheet
End Sub

'--------------------------------------------------------------
' 明細シートの入力・検査者印・総合判定と、表紙の関係する結果・印鑑を消す
'--------------------------------------------------------------
Public Sub ClearDetailData(ws As Worksheet)

    Dim cs As Worksheet
    Dim area As Range
    Dim inspName As String
    Dim eventsBefore As Boolean

    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)
    inspName = InspectionNameOf(ws)
    eventsBefore = Application.EnableEvents
    Application.EnableEvents = False

    Set area = DetailInputArea(ws)
    If Not area Is Nothing Then area.ClearContents
    ClearDetailJudgement ws

    SetCoverResult inspName, ""
    If ShapeExists(cs, INSP_STAMP_PREFIX & inspName) Then cs.Shapes(INSP_STAMP_PREFIX & inspName).Delete
    DeleteShapesByPrefix cs, "CoverStamp_"
    UpdateOverall

    Application.EnableEvents = eventsBefore

End Sub
