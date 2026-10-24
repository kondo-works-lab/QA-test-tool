Attribute VB_Name = "modDetail"
Option Explicit

'==============================================================
' QA試験 明細シート（検査ごとに1シート）
'   ・表紙の品名に応じて、マスタからその品名の検査分だけ明細シートを作り直す。
'   ・「判定」ボタンで次を行う。
'     ① 未入力チェック（白い入力欄の未入力・グレーの欄への入力・数値以外の入力）
'        問題があればメッセージを出して判定しない。
'     ② 各入力値を基準と照合し、基準外のセルの背景をオレンジにする。
'     ③ オレンジの個数が不合格個数以上なら「不合格」、未満なら「合格」。
'     ④ 判定結果を表紙の該当検査欄に反映し、総合判定を更新する。
'     ⑤ Access へ判定データを登録する（ボタン押下時刻を渡す）。
'==============================================================

Private Const JUDGE_BUTTON_NAME As String = "btnJudge"

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
Public Sub CreateDetailSheets(productName As String)

    Dim inspName As Variant
    Dim ws As Worksheet
    Dim afterSheet As Worksheet

    Set afterSheet = ThisWorkbook.Worksheets(COVER_SHEET)
    For Each inspName In GetInspectionNames(productName)
        Set ws = AddSheetAfter(afterSheet)
        ws.Name = DETAIL_PREFIX & inspName
        BuildDetailSheet ws, productName, CStr(inspName)
        Set afterSheet = ws
    Next inspName

    ThisWorkbook.Worksheets(COVER_SHEET).Activate

End Sub

'--------------------------------------------------------------
' 明細シート1枚分のレイアウトを作る
'--------------------------------------------------------------
Private Sub BuildDetailSheet(ws As Worksheet, productName As String, inspName As String)

    Dim ms As Worksheet
    Dim itemRows As Collection
    Dim mr As Variant
    Dim r As Long
    Dim c As Long
    Dim n As Long
    Dim maxN As Long
    Dim lastCol As Long
    Dim lastRow As Long
    Dim sumRow As Long
    Dim no As Long

    Set ms = ThisWorkbook.Worksheets(MASTER_SHEET)
    Set itemRows = GetItemRows(productName, inspName)

    For Each mr In itemRows
        If CLng(ms.Cells(mr, MCOL_N).Value) > maxN Then maxN = CLng(ms.Cells(mr, MCOL_N).Value)
    Next mr
    lastCol = DETAIL_COL_INPUT + maxN - 1
    lastRow = DETAIL_FIRST_ROW + itemRows.Count - 1
    sumRow = lastRow + 2

    ws.Cells.Font.Name = "メイリオ"
    ws.Cells.Font.Size = 10

    ws.Range("B2").Value = "検査明細　" & inspName
    ws.Range("B2").Font.Size = 14
    ws.Range("B2").Font.Bold = True

    ws.Range("C3").Value = "品名"
    ws.Range("C4").Value = "不合格基準"
    ws.Range("C3:C4").Interior.Color = COLOR_HEADER
    ws.Range(DETAIL_PRODUCT_CELL).Resize(1, 2).Merge
    ws.Range(DETAIL_FAILCOUNT_CELL).Resize(1, 2).Merge
    ws.Range(DETAIL_PRODUCT_CELL).NumberFormat = "@"
    ws.Range(DETAIL_PRODUCT_CELL).Value = productName
    ws.Range(DETAIL_FAILCOUNT_CELL).NumberFormat = """基準外 ""0"" 個以上で不合格"""
    ws.Range(DETAIL_FAILCOUNT_CELL).Value = GetFailCount(productName, inspName)
    ws.Range("D3:E4").HorizontalAlignment = xlLeft
    SetBorders ws.Range("C3:E4")

    ' 見出し
    ws.Cells(DETAIL_HEADER_ROW, DETAIL_COL_NO).Value = "No."
    ws.Cells(DETAIL_HEADER_ROW, DETAIL_COL_ITEM).Value = "検査項目"
    ws.Cells(DETAIL_HEADER_ROW, DETAIL_COL_METHOD).Value = "判定方式"
    ws.Cells(DETAIL_HEADER_ROW, DETAIL_COL_STD).Value = "基準"
    ws.Cells(DETAIL_HEADER_ROW, DETAIL_COL_STD1).Value = "基準値1"
    ws.Cells(DETAIL_HEADER_ROW, DETAIL_COL_STD2).Value = "基準値2"
    ws.Cells(DETAIL_HEADER_ROW, DETAIL_COL_N).Value = "n数"
    For c = 1 To maxN
        ws.Cells(DETAIL_HEADER_ROW, DETAIL_COL_INPUT + c - 1).Value = c
    Next c
    With ws.Range(ws.Cells(DETAIL_HEADER_ROW, DETAIL_COL_NO), ws.Cells(DETAIL_HEADER_ROW, lastCol))
        .Interior.Color = COLOR_HEADER
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    ' 項目
    r = DETAIL_FIRST_ROW
    For Each mr In itemRows
        no = no + 1
        n = CLng(ms.Cells(mr, MCOL_N).Value)
        ws.Cells(r, DETAIL_COL_NO).Value = no
        ws.Cells(r, DETAIL_COL_ITEM).Value = CellText(ms.Cells(mr, MCOL_ITEM).Value)
        ws.Cells(r, DETAIL_COL_METHOD).Value = CellText(ms.Cells(mr, MCOL_METHOD).Value)
        ws.Cells(r, DETAIL_COL_STD1).NumberFormat = "@"
        ws.Cells(r, DETAIL_COL_STD1).Value = CellText(ms.Cells(mr, MCOL_STD1).Value)
        ws.Cells(r, DETAIL_COL_STD2).NumberFormat = "@"
        ws.Cells(r, DETAIL_COL_STD2).Value = CellText(ms.Cells(mr, MCOL_STD2).Value)
        ws.Cells(r, DETAIL_COL_N).Value = n
        ws.Cells(r, DETAIL_COL_STD).Value = FormatStandard(ws.Cells(r, DETAIL_COL_METHOD).Value, _
            ws.Cells(r, DETAIL_COL_STD1).Value, ws.Cells(r, DETAIL_COL_STD2).Value)
        For c = 1 To maxN
            If c <= n Then
                ws.Cells(r, DETAIL_COL_INPUT + c - 1).Interior.Color = COLOR_INPUT
            Else
                ws.Cells(r, DETAIL_COL_INPUT + c - 1).Interior.Color = COLOR_INACTIVE
            End If
        Next c
        r = r + 1
    Next mr

    ws.Range(ws.Cells(DETAIL_FIRST_ROW, DETAIL_COL_NO), ws.Cells(lastRow, DETAIL_COL_METHOD)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(DETAIL_FIRST_ROW, DETAIL_COL_ITEM), ws.Cells(lastRow, DETAIL_COL_ITEM)).HorizontalAlignment = xlLeft
    ws.Range(ws.Cells(DETAIL_FIRST_ROW, DETAIL_COL_INPUT), ws.Cells(lastRow, lastCol)).HorizontalAlignment = xlCenter
    SetBorders ws.Range(ws.Cells(DETAIL_HEADER_ROW, DETAIL_COL_NO), ws.Cells(lastRow, lastCol))

    ' 集計（基準外の個数・判定）
    ws.Cells(sumRow, DETAIL_COL_ITEM).Value = "基準外の個数"
    ws.Cells(sumRow + 1, DETAIL_COL_ITEM).Value = "判定"
    ws.Range(ws.Cells(sumRow, DETAIL_COL_ITEM), ws.Cells(sumRow + 1, DETAIL_COL_ITEM)).Interior.Color = COLOR_HEADER
    ws.Range(ws.Cells(sumRow, DETAIL_COL_ITEM), ws.Cells(sumRow + 1, DETAIL_COL_ITEM)).Font.Bold = True
    ws.Range(ws.Cells(sumRow, DETAIL_COL_METHOD), ws.Cells(sumRow + 1, DETAIL_COL_METHOD)).HorizontalAlignment = xlCenter
    ws.Cells(sumRow + 1, DETAIL_COL_METHOD).Font.Bold = True
    SetBorders ws.Range(ws.Cells(sumRow, DETAIL_COL_ITEM), ws.Cells(sumRow + 1, DETAIL_COL_METHOD))

    ' 列幅（基準値1・基準値2・n数は判定用のため非表示）
    ws.Columns(1).ColumnWidth = 2
    ws.Columns(DETAIL_COL_NO).ColumnWidth = 5
    ws.Columns(DETAIL_COL_ITEM).ColumnWidth = 18
    ws.Columns(DETAIL_COL_METHOD).ColumnWidth = 9
    ws.Columns(DETAIL_COL_STD).ColumnWidth = 16
    ws.Range(ws.Columns(DETAIL_COL_STD1), ws.Columns(DETAIL_COL_N)).Hidden = True
    ws.Range(ws.Columns(DETAIL_COL_INPUT), ws.Columns(lastCol)).ColumnWidth = 7

    AddButton ws, JUDGE_BUTTON_NAME, "判定", "判定実行", _
        ws.Cells(2, DETAIL_COL_INPUT).Left, ws.Cells(2, DETAIL_COL_INPUT).Top, 80, 28

End Sub

Private Sub SetBorders(target As Range)
    With target.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
End Sub

'--------------------------------------------------------------
' 基準の表示用の文字を作る
'--------------------------------------------------------------
Public Function FormatStandard(method As Variant, std1 As Variant, std2 As Variant) As String
    Select Case CStr(method)
        Case METHOD_MATCH: FormatStandard = Join(SplitMatchValues(CStr(std1)), " / ")
        Case METHOD_MIN:   FormatStandard = CStr(std1) & " 以上"
        Case METHOD_MAX:   FormatStandard = CStr(std1) & " 以下"
        Case METHOD_RANGE: FormatStandard = CStr(std1) & " ～ " & CStr(std2)
        Case Else:         FormatStandard = ""
    End Select
End Function

'--------------------------------------------------------------
' 「一致」の合格の文字を配列にする（, ，、 で区切る。全角は半角にそろえる）
'--------------------------------------------------------------
Public Function SplitMatchValues(std1 As String) As Variant
    Dim text As String
    Dim parts As Variant
    Dim result() As String
    Dim i As Long
    Dim cnt As Long

    text = StrConv(Trim(std1), vbNarrow)
    text = Replace(text, "、", ",")
    text = Replace(text, "，", ",")
    parts = Split(text, ",")

    ReDim result(0 To UBound(parts))
    For i = 0 To UBound(parts)
        If Trim(parts(i)) <> "" Then
            result(cnt) = Trim(parts(i))
            cnt = cnt + 1
        End If
    Next i
    If cnt = 0 Then
        SplitMatchValues = Array()
    Else
        ReDim Preserve result(0 To cnt - 1)
        SplitMatchValues = result
    End If
End Function

'--------------------------------------------------------------
' 入力値が基準を満たすかどうか（入力チェック済みの値を渡すこと）
'   一致：前後の空白を除き、全角／半角・大文字／小文字を区別せずに比べる
'   以上・以下・範囲：境界値ちょうどは合格
'--------------------------------------------------------------
Public Function IsPassValue(inputValue As Variant, method As String, std1 As Variant, std2 As Variant) As Boolean

    Dim v As Double
    Dim s As String
    Dim passValue As Variant

    Select Case method
        Case METHOD_MATCH
            s = StrConv(Trim(CStr(inputValue)), vbNarrow)
            IsPassValue = False
            For Each passValue In SplitMatchValues(CStr(std1))
                If StrComp(s, CStr(passValue), vbTextCompare) = 0 Then
                    IsPassValue = True
                    Exit Function
                End If
            Next passValue
        Case METHOD_MIN
            IsPassValue = (CDbl(inputValue) >= CDbl(std1))
        Case METHOD_MAX
            IsPassValue = (CDbl(inputValue) <= CDbl(std1))
        Case METHOD_RANGE
            v = CDbl(inputValue)
            IsPassValue = (v >= CDbl(std1) And v <= CDbl(std2))
        Case Else
            IsPassValue = False
    End Select

End Function

'--------------------------------------------------------------
' 明細シートの項目数・入力欄の列数・集計行
'--------------------------------------------------------------
Public Function DetailItemCount(ws As Worksheet) As Long
    Dim r As Long
    r = DETAIL_FIRST_ROW
    Do While Not IsBlankValue(ws.Cells(r, DETAIL_COL_NO).Value)
        r = r + 1
    Loop
    DetailItemCount = r - DETAIL_FIRST_ROW
End Function

Public Function DetailInputColumns(ws As Worksheet) As Long
    Dim c As Long
    c = DETAIL_COL_INPUT
    Do While Not IsBlankValue(ws.Cells(DETAIL_HEADER_ROW, c).Value)
        c = c + 1
    Loop
    DetailInputColumns = c - DETAIL_COL_INPUT
End Function

Public Function DetailSumRow(ws As Worksheet) As Long
    DetailSumRow = DETAIL_FIRST_ROW + DetailItemCount(ws) + 1   ' 基準外の個数の行（次の行が判定）
End Function

'--------------------------------------------------------------
' 入力欄全体（グレーの欄を含む）。項目が無ければ Nothing
'--------------------------------------------------------------
Public Function DetailInputArea(ws As Worksheet) As Range
    Dim cnt As Long
    Dim cols As Long
    cnt = DetailItemCount(ws)
    cols = DetailInputColumns(ws)
    If cnt = 0 Or cols = 0 Then
        Set DetailInputArea = Nothing
    Else
        Set DetailInputArea = ws.Range(ws.Cells(DETAIL_FIRST_ROW, DETAIL_COL_INPUT), _
                                       ws.Cells(DETAIL_FIRST_ROW + cnt - 1, DETAIL_COL_INPUT + cols - 1))
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
' 判定済みかどうか（判定欄に結果がある）
'--------------------------------------------------------------
Public Function IsJudged(ws As Worksheet) As Boolean
    IsJudged = (CellText(ws.Cells(DetailSumRow(ws) + 1, DETAIL_COL_METHOD).Value) <> "")
End Function

Public Function GetDetailResult(ws As Worksheet) As String
    GetDetailResult = CellText(ws.Cells(DetailSumRow(ws) + 1, DETAIL_COL_METHOD).Value)
End Function

'--------------------------------------------------------------
' 判定結果（オレンジ・基準外の個数・判定）を消す。何か消したら True
'--------------------------------------------------------------
Public Function ClearDetailJudgement(ws As Worksheet) As Boolean

    Dim sumRow As Long
    Dim cell As Range
    Dim area As Range

    sumRow = DetailSumRow(ws)
    If IsJudged(ws) Then ClearDetailJudgement = True

    Set area = DetailInputArea(ws)
    If Not area Is Nothing Then
        For Each cell In area.Cells
            If cell.Interior.Color = COLOR_NG Then
                cell.Interior.Color = COLOR_INPUT
                ClearDetailJudgement = True
            End If
        Next cell
    End If

    ws.Cells(sumRow, DETAIL_COL_METHOD).ClearContents
    ws.Cells(sumRow + 1, DETAIL_COL_METHOD).ClearContents

End Function

'--------------------------------------------------------------
' 「判定」ボタン：ボタンのある明細シートを判定する
'--------------------------------------------------------------
Sub 判定実行()

    Dim pressTime As Date

    pressTime = Now   ' ボタン押下時刻（Access に登録する）

    If Not IsDetailSheet(ActiveSheet) Then
        ShowMsg "明細シートで「判定」を押してください。", vbExclamation, "判定"
        Exit Sub
    End If
    JudgeDetailSheet ActiveSheet, pressTime

End Sub

'--------------------------------------------------------------
' 明細シートを判定する。判定できたら True
'--------------------------------------------------------------
Public Function JudgeDetailSheet(ws As Worksheet, pressTime As Date) As Boolean

    Dim cnt As Long
    Dim cols As Long
    Dim r As Long
    Dim c As Long
    Dim n As Long
    Dim cell As Range
    Dim method As String
    Dim blankList As String
    Dim grayList As String
    Dim numberErrList As String
    Dim stdErrList As String
    Dim msg As String
    Dim ngCount As Long
    Dim failCount As Long
    Dim result As String
    Dim sumRow As Long
    Dim eventsBefore As Boolean
    Dim errMsg As String

    On Error GoTo ErrHandler
    eventsBefore = Application.EnableEvents

    cnt = DetailItemCount(ws)
    cols = DetailInputColumns(ws)
    If cnt = 0 Then
        ShowMsg "判定する項目がありません。表紙の品名を確認してください。", vbExclamation, "判定"
        Exit Function
    End If

    failCount = 0
    If IsNumericValue(ws.Range(DETAIL_FAILCOUNT_CELL).Value) Then failCount = CLng(ws.Range(DETAIL_FAILCOUNT_CELL).Value)
    If failCount < 1 Then
        ShowMsg "不合格基準（" & DETAIL_FAILCOUNT_CELL & "）が正しくありません。表紙の品名を入れ直してください。", vbExclamation, "判定"
        Exit Function
    End If

    '------------------------------------------------
    ' ① 入力チェック
    '------------------------------------------------
    For r = DETAIL_FIRST_ROW To DETAIL_FIRST_ROW + cnt - 1
        method = CellText(ws.Cells(r, DETAIL_COL_METHOD).Value)
        n = 0
        If IsNumericValue(ws.Cells(r, DETAIL_COL_N).Value) Then n = CLng(ws.Cells(r, DETAIL_COL_N).Value)
        If Not IsValidStandard(method, ws.Cells(r, DETAIL_COL_STD1).Value, ws.Cells(r, DETAIL_COL_STD2).Value) Then
            stdErrList = AppendAddress(stdErrList, ws.Cells(r, DETAIL_COL_STD))
        End If
        For c = 1 To cols
            Set cell = ws.Cells(r, DETAIL_COL_INPUT + c - 1)
            If c > n Then
                If Not IsBlankValue(cell.Value) Then grayList = AppendAddress(grayList, cell)
            ElseIf IsBlankValue(cell.Value) Then
                blankList = AppendAddress(blankList, cell)
            ElseIf method <> METHOD_MATCH And Not IsNumericValue(cell.Value) Then
                numberErrList = AppendAddress(numberErrList, cell)
            End If
        Next c
    Next r

    If blankList <> "" Then msg = msg & "・未入力の欄があります：" & blankList & vbCrLf
    If grayList <> "" Then msg = msg & "・グレーの欄に入力があります：" & grayList & vbCrLf
    If numberErrList <> "" Then msg = msg & "・数値以外が入力されています：" & numberErrList & vbCrLf
    If stdErrList <> "" Then msg = msg & "・基準を読み取れません（表紙の品名を入れ直してください）：" & stdErrList & vbCrLf
    If msg <> "" Then
        ShowMsg msg & vbCrLf & "修正してから、もう一度「判定」を押してください。", vbExclamation, "入力確認"
        Exit Function
    End If

    '------------------------------------------------
    ' ② 基準外のセルをオレンジにする　③ 合否
    '------------------------------------------------
    Application.EnableEvents = False
    For r = DETAIL_FIRST_ROW To DETAIL_FIRST_ROW + cnt - 1
        method = CellText(ws.Cells(r, DETAIL_COL_METHOD).Value)
        n = CLng(ws.Cells(r, DETAIL_COL_N).Value)
        For c = 1 To n
            Set cell = ws.Cells(r, DETAIL_COL_INPUT + c - 1)
            If IsPassValue(cell.Value, method, ws.Cells(r, DETAIL_COL_STD1).Value, ws.Cells(r, DETAIL_COL_STD2).Value) Then
                cell.Interior.Color = COLOR_INPUT
            Else
                cell.Interior.Color = COLOR_NG
                ngCount = ngCount + 1
            End If
        Next c
    Next r

    If ngCount >= failCount Then
        result = RESULT_FAIL
    Else
        result = RESULT_PASS
    End If

    sumRow = DetailSumRow(ws)
    ws.Cells(sumRow, DETAIL_COL_METHOD).Value = ngCount
    ws.Cells(sumRow + 1, DETAIL_COL_METHOD).Value = result

    '------------------------------------------------
    ' ④ 表紙へ反映・総合判定
    '------------------------------------------------
    SetCoverResult InspectionNameOf(ws), result
    UpdateOverall
    Application.EnableEvents = eventsBefore

    '------------------------------------------------
    ' ⑤ Access へ登録（失敗してもシートの判定結果は残す）
    '------------------------------------------------
    RegisterJudgement ws, pressTime

    JudgeDetailSheet = True
    Exit Function

ErrHandler:
    errMsg = Err.Description
    Application.EnableEvents = eventsBefore
    ShowMsg "判定中にエラーが発生しました。" & vbCrLf & errMsg, vbExclamation, "判定"

End Function

'--------------------------------------------------------------
' 明細シートに保存された基準が使えるかどうか
'--------------------------------------------------------------
Private Function IsValidStandard(method As String, std1 As Variant, std2 As Variant) As Boolean
    Select Case method
        Case METHOD_MATCH
            IsValidStandard = (UBound(SplitMatchValues(CStr(std1))) >= 0)
        Case METHOD_MIN, METHOD_MAX
            IsValidStandard = IsNumericValue(std1)
        Case METHOD_RANGE
            If IsNumericValue(std1) And IsNumericValue(std2) Then
                IsValidStandard = (CDbl(std1) <= CDbl(std2))
            End If
        Case Else
            IsValidStandard = False
    End Select
End Function
