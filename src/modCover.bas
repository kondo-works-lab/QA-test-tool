Attribute VB_Name = "modCover"
Option Explicit

'==============================================================
' QA試験 表紙
'   ・品名（C4）が入力されたら、マスタからその品名の検査分だけ
'     明細シートを作り、表紙に No. と検査名を並べる（ロット番号は後から入れてもよい）。
'     品名を変えたら作り直す（判定結果と印鑑は消え、メッセージで再判定・再押印を促す）。
'     ロット番号（C5）を入れたり変えたりした場合は、明細シートのロット番号の表示だけ変える。
'   ・マスタに無い品名を入れたら、似た品名を候補としてメッセージに出す。
'   ・明細シートで検査者印を押すと、その検査の判定結果と担当者印（データ印）が
'     表紙に自動で入る（modDetail.ApplyInspectorStamp）。
'   ・明細シートの検査項目・OK欄が押印の後に変更された場合、その検査の判定結果と
'     印鑑、表紙の承認者印・担当者印を消去し、再押印を促す。
'     入力欄がすべて空になった場合（クリア等）はメッセージを出さない。
'   ・総合判定（F4:F7）：1つでも NG なら「NG」、全検査が OK なら「OK」、
'     それ以外（判定していない検査がある）は空欄。
'   ・「表紙削除」：表紙の入力（品名・ロット番号・状態・日付）と印鑑を消す。
'     品名が消えるので明細シートも消える。
'   ・「一括削除」：全明細シートの入力と印鑑を消してから、表紙削除と同じことをする。
'     終わったら品名の欄を選ぶ。どちらも消す前に確認する。
'==============================================================

Public Const INSP_STAMP_PREFIX As String = "InspStamp_"           ' 表紙の検査担当者印の名前（InspStamp_検査名）
Public Const APPROVER_STAMP_NAME As String = "CoverStamp_承認者"
Public Const CHARGE_STAMP_NAME As String = "CoverStamp_担当者"

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
' 総合判定欄（F4:F7）
'--------------------------------------------------------------
Public Function OverallCell() As Range
    Set OverallCell = ThisWorkbook.Worksheets(COVER_SHEET).Range(COVER_OVERALL_CELL)
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

    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)
    oldCount = CoverListCount()

    ' 前の一覧と担当者印を消す
    DeleteShapesByPrefix cs, INSP_STAMP_PREFIX
    With cs.Range(cs.Cells(COVER_LIST_FIRST_ROW, COVER_COL_NO), cs.Cells(COVER_LIST_FIRST_ROW + oldCount + 2, COVER_COL_STAMP))
        .UnMerge
        .Clear
        .EntireRow.RowHeight = cs.StandardHeight
    End With

    r = COVER_LIST_FIRST_ROW
    For Each inspName In inspNames
        no = no + 1
        cs.Cells(r, COVER_COL_NO).Value = no
        cs.Cells(r, COVER_COL_NAME).Value = inspName
        EnsureStampArea cs.Cells(r, COVER_COL_STAMP)   ' 担当者印（データ印）が収まる高さにする
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
        cs.Range(cs.Cells(COVER_LIST_FIRST_ROW, COVER_COL_RESULT), cs.Cells(lastRow, COVER_COL_RESULT)).Font.Size = 14
        cs.Range(cs.Cells(COVER_LIST_FIRST_ROW, COVER_COL_RESULT), cs.Cells(lastRow, COVER_COL_RESULT)).Font.Bold = True
    End If

    OverallCell().MergeArea.ClearContents

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
'   1つでも NG →「NG」、全検査が OK →「OK」、それ以外 → 空欄
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
    Dim allOk As Boolean
    Dim v As String

    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)
    cnt = CoverListCount()
    If cnt = 0 Then
        CalcOverall = ""
        Exit Function
    End If

    allOk = True
    For r = COVER_LIST_FIRST_ROW To COVER_LIST_FIRST_ROW + cnt - 1
        v = CellText(cs.Cells(r, COVER_COL_RESULT).Value)
        If v = RESULT_NG Then
            CalcOverall = RESULT_NG
            Exit Function
        ElseIf v <> RESULT_OK Then
            allOk = False
        End If
    Next r

    If allOk Then
        CalcOverall = RESULT_OK
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

'--------------------------------------------------------------
' シートの変更（ThisWorkbook の Workbook_SheetChange から呼ばれる）
'--------------------------------------------------------------
Public Sub HandleSheetChange(ByVal Sh As Object, ByVal Target As Range)

    Dim ws As Worksheet
    Dim area As Range

    If Sh.Name = COVER_SHEET Then
        ' 入力が変わったことを記録する（DB 登録の前に保存し直してもらうため）
        If Not Intersect(Target, Sh.Range(COVER_PRODUCT_CELL & "," & COVER_LOT_CELL & "," & _
            COVER_STATE_CELL & "," & COVER_DATE_CELL)) Is Nothing Then SetInputChanged True
        If Not Intersect(Target, Sh.Range(COVER_PRODUCT_CELL)) Is Nothing Then
            OnProductChanged
        ElseIf Not Intersect(Target, Sh.Range(COVER_LOT_CELL)) Is Nothing Then
            OnLotChanged
        End If
    ElseIf IsDetailSheet(Sh) Then
        Set ws = Sh
        Set area = DetailInputArea(ws)
        If area Is Nothing Then Exit Sub
        If Not Intersect(Target, area) Is Nothing Then
            SetInputChanged True
            OnDetailInputChanged ws
        End If
    End If

End Sub

'--------------------------------------------------------------
' 品名が変更された：明細シートと表紙の検査一覧を作り直す
'--------------------------------------------------------------
Public Sub OnProductChanged()

    Dim cs As Worksheet
    Dim productName As String
    Dim lotNumber As String
    Dim hadData As Boolean
    Dim removedStamp As Boolean
    Dim errText As String
    Dim candidates As String
    Dim msg As String
    Dim eventsBefore As Boolean
    Dim errMsg As String
    Dim emptyList As New Collection

    On Error GoTo ErrHandler
    eventsBefore = Application.EnableEvents
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)
    productName = CellText(cs.Range(COVER_PRODUCT_CELL).Value)
    lotNumber = CellText(cs.Range(COVER_LOT_CELL).Value)

    ' 判定結果・印鑑を消去し、明細シートを削除する
    hadData = DeleteDetailSheets()
    If DeleteShapesByPrefix(cs, INSP_STAMP_PREFIX) > 0 Then removedStamp = True
    If DeleteShapesByPrefix(cs, "CoverStamp_") > 0 Then removedStamp = True
    RebuildCoverList emptyList

    If productName = "" Then GoTo Finish

    If Not ProductExists(productName) Then
        Application.ScreenUpdating = True
        msg = "品名「" & productName & "」がマスタシートにありません。" & vbCrLf
        candidates = FindProductCandidates(productName)
        If candidates <> "" Then
            msg = msg & vbCrLf & "似た品名の候補：" & vbCrLf & "　" & candidates & vbCrLf
        End If
        ShowMsg msg & vbCrLf & "品名を確認してください。", vbExclamation, "品名確認"
        GoTo Finish
    End If

    If Not ValidateProductMaster(productName, errText) Then
        Application.ScreenUpdating = True
        ShowMsg "マスタシートの「" & productName & "」の設定に誤りがあります。" & vbCrLf & errText & vbCrLf & _
                "マスタを修正してから、品名を入れ直してください。", vbExclamation, "品名確認"
        GoTo Finish
    End If

    CreateDetailSheets productName, lotNumber
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
' ロット番号が変更された
'   明細シートのロット番号の表示だけ変える（入力・判定・印鑑はそのまま）
'--------------------------------------------------------------
Public Sub OnLotChanged()

    Dim cs As Worksheet
    Dim eventsBefore As Boolean

    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)
    eventsBefore = Application.EnableEvents
    Application.EnableEvents = False
    UpdateDetailLot CellText(cs.Range(COVER_LOT_CELL).Value)
    Application.EnableEvents = eventsBefore

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

    If removedJudge Or removedStamp Then
        msg = "入力内容が変更されたため、判定結果と印鑑を消去しました。" & vbCrLf & _
              "もう一度「検査者押印」を押してください。"
        ShowMsg msg, vbInformation, "判定結果の消去"
    End If
    Exit Sub

ErrHandler:
    errMsg = Err.Description
    Application.EnableEvents = eventsBefore
    ShowMsg "判定結果の消去に失敗しました。" & vbCrLf & errMsg, vbExclamation, "判定結果の消去"

End Sub

'--------------------------------------------------------------
' 「表紙削除」ボタン
'--------------------------------------------------------------
Sub 表紙削除()
    If ShowMsg("表紙の入力（品名・ロット番号・状態・日付）と印鑑を削除します。" & vbCrLf & _
               "品名が消えるため、明細シートも削除されます。よろしいですか？", _
               vbYesNo + vbExclamation + vbDefaultButton2, "表紙削除") = vbNo Then Exit Sub
    ClearCoverData
End Sub

'--------------------------------------------------------------
' 「一括削除」ボタン
'--------------------------------------------------------------
Sub 一括削除()

    Dim item As Variant
    Dim ws As Worksheet

    If ShowMsg("表紙と全ての明細シートの入力と印鑑を削除します。" & vbCrLf & _
               "品名が消えるため、明細シートも削除されます。よろしいですか？", _
               vbYesNo + vbExclamation + vbDefaultButton2, "一括削除") = vbNo Then Exit Sub

    For Each item In GetDetailSheets()
        Set ws = item
        ClearDetailData ws
    Next item
    ClearCoverData

    ThisWorkbook.Worksheets(COVER_SHEET).Activate
    ThisWorkbook.Worksheets(COVER_SHEET).Range(COVER_PRODUCT_CELL).Select

End Sub

'--------------------------------------------------------------
' 表紙の入力と印鑑を消す（品名が消えるので明細シート・一覧も消える）
'--------------------------------------------------------------
Public Sub ClearCoverData()

    Dim cs As Worksheet
    Dim eventsBefore As Boolean

    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)
    eventsBefore = Application.EnableEvents
    Application.EnableEvents = False
    cs.Range(COVER_PRODUCT_CELL).MergeArea.ClearContents
    cs.Range(COVER_LOT_CELL).MergeArea.ClearContents
    cs.Range(COVER_STATE_CELL).MergeArea.ClearContents
    cs.Range(COVER_DATE_CELL).MergeArea.ClearContents
    Application.EnableEvents = eventsBefore

    OnProductChanged   ' 品名が空 → 明細シート・一覧・印鑑・総合判定を消す

End Sub


