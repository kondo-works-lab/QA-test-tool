Attribute VB_Name = "modCover"
Option Explicit

'==============================================================
' QA試験 表紙
'   ・品名（C4）が変更されたら、マスタからその品名の検査分だけ
'     明細シートを作り直し、表紙の検査一覧（判定結果・担当者印欄）を並べ直す。
'     判定結果と印鑑は消去する（既存ツールと同じく、メッセージで再判定・再押印を促す）。
'   ・明細シートの入力欄が判定の後に変更された場合、その検査の判定結果と
'     担当者印、表紙の承認者印・担当者印を消去し、再判定・再押印を促す。
'     入力欄がすべて空になった場合（クリア等）はメッセージを出さない。
'   ・総合判定：1つでも不合格なら「不合格」、全検査が合格なら「合格」、
'     それ以外（未判定の検査がある）は空欄。
'==============================================================

Public Const INSP_STAMP_PREFIX As String = "InspStamp_"           ' 検査担当者印の名前（InspStamp_検査名）
Public Const INSP_BUTTON_PREFIX As String = "btnInspStamp_"       ' 検査担当者の押印ボタン（btnInspStamp_No）
Public Const APPROVER_STAMP_NAME As String = "CoverStamp_承認者"
Public Const CHARGE_STAMP_NAME As String = "CoverStamp_担当者"
Private Const LIST_ROW_HEIGHT As Double = 36

'--------------------------------------------------------------
' 表紙の検査一覧の行数（No. が入っている行の数）
'--------------------------------------------------------------
Public Function CoverListCount() As Long
    Dim cs As Worksheet
    Dim r As Long
    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)
    r = COVER_LIST_FIRST_ROW
    Do While Not IsBlankValue(cs.Cells(r, COVER_COL_NO).Value)
        r = r + 1
    Loop
    CoverListCount = r - COVER_LIST_FIRST_ROW
End Function

'--------------------------------------------------------------
' 総合判定欄（検査一覧の1行下）
'--------------------------------------------------------------
Public Function OverallCell() As Range
    Dim cs As Worksheet
    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)
    Set OverallCell = cs.Cells(COVER_LIST_FIRST_ROW + CoverListCount() + 1, COVER_COL_RESULT)
End Function

'--------------------------------------------------------------
' 検査名の表紙の行。無ければ 0
'--------------------------------------------------------------
Public Function FindCoverRow(inspName As String) As Long
    Dim cs As Worksheet
    Dim r As Long
    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)
    For r = COVER_LIST_FIRST_ROW To COVER_LIST_FIRST_ROW + CoverListCount() - 1
        If CellText(cs.Cells(r, COVER_COL_NAME).Value) = inspName Then
            FindCoverRow = r
            Exit Function
        End If
    Next r
    FindCoverRow = 0
End Function

'--------------------------------------------------------------
' 表紙の検査一覧を作り直す（判定結果・担当者印・総合判定は空にする）
'--------------------------------------------------------------
Public Sub RebuildCoverList(inspNames As Collection)

    Dim cs As Worksheet
    Dim oldCount As Long
    Dim r As Long
    Dim no As Long
    Dim inspName As Variant
    Dim lastRow As Long
    Dim ovRow As Long
    Dim btnCell As Range

    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)
    oldCount = CoverListCount()

    ' 前の一覧（総合判定の行まで）と担当者印・押印ボタンを消す
    DeleteShapesByPrefix cs, INSP_STAMP_PREFIX
    DeleteShapesByPrefix cs, INSP_BUTTON_PREFIX
    With cs.Range(cs.Cells(COVER_LIST_FIRST_ROW, COVER_COL_NO), cs.Cells(COVER_LIST_FIRST_ROW + oldCount + 2, COVER_COL_STAMP + 1))
        .UnMerge
        .Clear
        .EntireRow.RowHeight = cs.StandardHeight
    End With

    r = COVER_LIST_FIRST_ROW
    For Each inspName In inspNames
        no = no + 1
        cs.Cells(r, COVER_COL_NO).Value = no
        cs.Cells(r, COVER_COL_NAME).Value = inspName
        cs.Rows(r).RowHeight = LIST_ROW_HEIGHT
        Set btnCell = cs.Cells(r, COVER_COL_STAMP + 1)
        AddButton cs, INSP_BUTTON_PREFIX & no, "押印", "検査担当者押印", _
            btnCell.Left + 4, btnCell.Top + (btnCell.Height - 22) / 2, 50, 22
        r = r + 1
    Next inspName
    lastRow = r - 1

    If no > 0 Then
        With cs.Range(cs.Cells(COVER_LIST_FIRST_ROW, COVER_COL_NO), cs.Cells(lastRow, COVER_COL_STAMP))
            .Borders.LineStyle = xlContinuous
            .VerticalAlignment = xlCenter
        End With
        cs.Range(cs.Cells(COVER_LIST_FIRST_ROW, COVER_COL_NO), cs.Cells(lastRow, COVER_COL_NO)).HorizontalAlignment = xlCenter
        cs.Range(cs.Cells(COVER_LIST_FIRST_ROW, COVER_COL_RESULT), cs.Cells(lastRow, COVER_COL_RESULT)).HorizontalAlignment = xlCenter
    End If

    ' 総合判定の行
    ovRow = COVER_LIST_FIRST_ROW + no + 1
    cs.Cells(ovRow, COVER_COL_NAME).Value = "総合判定"
    cs.Cells(ovRow, COVER_COL_NAME).Interior.Color = COLOR_HEADER
    cs.Cells(ovRow, COVER_COL_NAME).Font.Bold = True
    cs.Cells(ovRow, COVER_COL_NAME).HorizontalAlignment = xlCenter
    With cs.Cells(ovRow, COVER_COL_RESULT)
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    cs.Rows(ovRow).RowHeight = LIST_ROW_HEIGHT
    cs.Range(cs.Cells(ovRow, COVER_COL_NAME), cs.Cells(ovRow, COVER_COL_RESULT)).Borders.LineStyle = xlContinuous

End Sub

'--------------------------------------------------------------
' 表紙の検査の判定結果を設定する（空文字で消去）
'--------------------------------------------------------------
Public Sub SetCoverResult(inspName As String, result As String)
    Dim cs As Worksheet
    Dim r As Long
    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)
    r = FindCoverRow(inspName)
    If r = 0 Then Exit Sub
    cs.Cells(r, COVER_COL_RESULT).Value = result
    SetResultFont cs.Cells(r, COVER_COL_RESULT), result
End Sub

'--------------------------------------------------------------
' 総合判定を更新する
'   1つでも不合格 →「不合格」、全検査が合格 →「合格」、それ以外 → 空欄
'--------------------------------------------------------------
Public Sub UpdateOverall()
    Dim result As String
    Dim cell As Range
    result = CalcOverall()
    Set cell = OverallCell()
    cell.Value = result
    SetResultFont cell, result
End Sub

Public Function CalcOverall() As String

    Dim cs As Worksheet
    Dim r As Long
    Dim cnt As Long
    Dim allPass As Boolean
    Dim v As String

    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)
    cnt = CoverListCount()
    If cnt = 0 Then
        CalcOverall = ""
        Exit Function
    End If

    allPass = True
    For r = COVER_LIST_FIRST_ROW To COVER_LIST_FIRST_ROW + cnt - 1
        v = CellText(cs.Cells(r, COVER_COL_RESULT).Value)
        If v = RESULT_FAIL Then
            CalcOverall = RESULT_FAIL
            Exit Function
        ElseIf v <> RESULT_PASS Then
            allPass = False
        End If
    Next r

    If allPass Then
        CalcOverall = RESULT_PASS
    Else
        CalcOverall = ""
    End If

End Function

Public Function GetOverall() As String
    GetOverall = CellText(OverallCell().Value)
End Function

'--------------------------------------------------------------
' 全検査が判定済みかどうか
'--------------------------------------------------------------
Public Function AllInspectionsJudged() As Boolean
    Dim cs As Worksheet
    Dim r As Long
    Dim cnt As Long
    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)
    cnt = CoverListCount()
    If cnt = 0 Then Exit Function
    For r = COVER_LIST_FIRST_ROW To COVER_LIST_FIRST_ROW + cnt - 1
        If CellText(cs.Cells(r, COVER_COL_RESULT).Value) = "" Then Exit Function
    Next r
    AllInspectionsJudged = True
End Function

Private Sub SetResultFont(cell As Range, result As String)
    If result = RESULT_FAIL Then
        cell.Font.Color = RGB(192, 0, 0)
    Else
        cell.Font.Color = RGB(0, 0, 0)
    End If
End Sub

'--------------------------------------------------------------
' シートの変更（ThisWorkbook の Workbook_SheetChange から呼ばれる）
'--------------------------------------------------------------
Public Sub HandleSheetChange(ByVal sh As Object, ByVal target As Range)

    Dim ws As Worksheet
    Dim area As Range

    If sh.Name = COVER_SHEET Then
        If Not Intersect(target, sh.Range(COVER_PRODUCT_CELL)) Is Nothing Then OnProductChanged
    ElseIf IsDetailSheet(sh) Then
        Set ws = sh
        Set area = DetailInputArea(ws)
        If area Is Nothing Then Exit Sub
        If Not Intersect(target, area) Is Nothing Then OnDetailInputChanged ws
    End If

End Sub

'--------------------------------------------------------------
' 品名が変更された：明細シートと表紙の検査一覧を作り直す
'--------------------------------------------------------------
Public Sub OnProductChanged()

    Dim cs As Worksheet
    Dim productName As String
    Dim hadData As Boolean
    Dim removedStamp As Boolean
    Dim errText As String
    Dim eventsBefore As Boolean
    Dim errMsg As String
    Dim emptyList As New Collection

    On Error GoTo ErrHandler
    eventsBefore = Application.EnableEvents
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)
    productName = CellText(cs.Range(COVER_PRODUCT_CELL).Value)

    ' 判定結果・印鑑を消去し、明細シートを削除する
    hadData = DeleteDetailSheets()
    If DeleteShapesByPrefix(cs, INSP_STAMP_PREFIX) > 0 Then removedStamp = True
    If DeleteShapesByPrefix(cs, "CoverStamp_") > 0 Then removedStamp = True
    RebuildCoverList emptyList

    If productName = "" Then GoTo Finish

    If Not ProductExists(productName) Then
        Application.ScreenUpdating = True
        ShowMsg "品名「" & productName & "」がマスタシートにありません。" & vbCrLf & _
                "品名を確認してください。", vbExclamation, "品名確認"
        GoTo Finish
    End If

    If Not ValidateProductMaster(productName, errText) Then
        Application.ScreenUpdating = True
        ShowMsg "マスタシートの「" & productName & "」の設定に誤りがあります。" & vbCrLf & errText & vbCrLf & _
                "マスタを修正してから、品名を入れ直してください。", vbExclamation, "品名確認"
        GoTo Finish
    End If

    CreateDetailSheets productName
    RebuildCoverList GetInspectionNames(productName)
    cs.Activate

    If hadData Or removedStamp Then
        Application.ScreenUpdating = True
        ShowMsg "品名が変更されたため、明細シートを作り直しました。" & vbCrLf & _
                "（入力データ・判定結果・印鑑は消去されました）", vbInformation, "明細シートの作り直し"
    End If

Finish:
    Application.ScreenUpdating = True
    Application.EnableEvents = eventsBefore
    Exit Sub

ErrHandler:
    errMsg = Err.Description
    Application.ScreenUpdating = True
    Application.EnableEvents = eventsBefore
    ShowMsg "明細シートの作成に失敗しました。" & vbCrLf & errMsg, vbExclamation, "品名確認"

End Sub

'--------------------------------------------------------------
' 明細シートの入力欄が変更された：その検査の判定結果と印鑑を消去する
'--------------------------------------------------------------
Public Sub OnDetailInputChanged(ws As Worksheet)

    Dim cs As Worksheet
    Dim inspName As String
    Dim removedJudge As Boolean
    Dim removedStamp As Boolean
    Dim eventsBefore As Boolean
    Dim msg As String
    Dim errMsg As String

    On Error GoTo ErrHandler
    eventsBefore = Application.EnableEvents
    Application.EnableEvents = False

    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)
    inspName = InspectionNameOf(ws)

    removedJudge = ClearDetailJudgement(ws)
    If FindCoverRow(inspName) > 0 Then
        If CellText(cs.Cells(FindCoverRow(inspName), COVER_COL_RESULT).Value) <> "" Then removedJudge = True
        SetCoverResult inspName, ""
    End If
    UpdateOverall

    If ShapeExists(cs, INSP_STAMP_PREFIX & inspName) Then
        cs.Shapes(INSP_STAMP_PREFIX & inspName).Delete
        removedStamp = True
    End If
    If DeleteShapesByPrefix(cs, "CoverStamp_") > 0 Then removedStamp = True

    Application.EnableEvents = eventsBefore

    ' 入力欄がすべて空（クリア等）の場合はメッセージを出さない
    If Not HasAnyInput(ws) Then Exit Sub

    If removedJudge And removedStamp Then
        msg = "入力内容が変更されたため、判定結果と印鑑を消去しました。" & vbCrLf & _
              "もう一度「判定」と押印を行ってください。"
    ElseIf removedJudge Then
        msg = "入力内容が変更されたため、判定結果を消去しました。" & vbCrLf & _
              "もう一度「判定」を押してください。"
    ElseIf removedStamp Then
        msg = "入力内容が変更されたため、印鑑を消去しました。" & vbCrLf & _
              "もう一度押印を行ってください。"
    End If
    If msg <> "" Then ShowMsg msg, vbInformation, "判定結果の消去"
    Exit Sub

ErrHandler:
    errMsg = Err.Description
    Application.EnableEvents = eventsBefore
    ShowMsg "判定結果の消去に失敗しました。" & vbCrLf & errMsg, vbExclamation, "判定結果の消去"

End Sub
