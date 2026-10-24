Attribute VB_Name = "modStamp"
Option Explicit

'==============================================================
' QA試験 電子印鑑（既存の QA試験.xlsm の電子印鑑機能を流用）
'   ・電子印鑑シートの B3:B10 の氏名から印影（Stamp_氏名）を作成・削除する。
'   ・表紙の承認者欄・担当者欄、各検査の担当者印欄に押印する。
'     押印する人は既存の承認者選択フォーム（frmApprover）で選ぶ。
'     frmApprover は OK で ApplyApproverStamp 氏名 を呼ぶ。
'   ・承認者・担当者は全検査の判定後、各検査の担当者はその検査の判定後に押印できる。
'==============================================================

Private Const STAMP_SHAPE_PREFIX As String = "Stamp_"

' 押印先（ApplyApproverStamp で使う）
Private mTargetArea As Range
Private mTargetShapeName As String

'--------------------------------------------------------------
' 表紙の承認者欄に押印する
'--------------------------------------------------------------
Sub 承認者押印()
    If Not AllInspectionsJudged() Then
        ShowMsg "判定が済んでいない検査があるため、承認者印を押せません。" & vbCrLf & _
                "先に各明細シートで「判定」を押してください。", vbExclamation, "承認者押印"
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
                "先に各明細シートで「判定」を押してください。", vbExclamation, "担当者押印"
        Exit Sub
    End If
    SetStampTarget ThisWorkbook.Worksheets(COVER_SHEET).Range(COVER_CHARGE_CELL).MergeArea, CHARGE_STAMP_NAME
    ShowStampForm
End Sub

'--------------------------------------------------------------
' 表紙の検査一覧の「押印」ボタン：その検査の担当者印欄に押印する
'   ボタン名（btnInspStamp_No）から検査の行を求める
'--------------------------------------------------------------
Sub 検査担当者押印()

    Dim cs As Worksheet
    Dim callerName As String
    Dim no As Long
    Dim r As Long
    Dim inspName As String

    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)
    callerName = CStr(Application.Caller)
    no = Val(Mid(callerName, Len(INSP_BUTTON_PREFIX) + 1))
    If no < 1 Or no > CoverListCount() Then Exit Sub

    r = COVER_LIST_FIRST_ROW + no - 1
    inspName = CellText(cs.Cells(r, COVER_COL_NAME).Value)
    If CellText(cs.Cells(r, COVER_COL_RESULT).Value) = "" Then
        ShowMsg "「" & inspName & "」の判定が済んでいないため、担当者印を押せません。" & vbCrLf & _
                "先に明細シートで「判定」を押してください。", vbExclamation, "検査担当者押印"
        Exit Sub
    End If

    SetStampTarget cs.Cells(r, COVER_COL_STAMP), INSP_STAMP_PREFIX & inspName
    ShowStampForm

End Sub

'--------------------------------------------------------------
' 押印先を設定する（テストからも使う）
'--------------------------------------------------------------
Public Sub SetStampTarget(targetArea As Range, shapeName As String)
    Set mTargetArea = targetArea
    mTargetShapeName = shapeName
End Sub

'--------------------------------------------------------------
' 承認者選択フォームを表示する
'   フォームが無くてもブック全体がコンパイルエラーにならないよう、名前で呼び出す
'--------------------------------------------------------------
Private Sub ShowStampForm()
    On Error GoTo ErrHandler
    VBA.UserForms.Add("frmApprover").Show
    Exit Sub
ErrHandler:
    ShowMsg "承認者選択フォーム（frmApprover）を表示できません。" & vbCrLf & _
            "README の導入手順に従って frmApprover を取り込んでください。" & vbCrLf & Err.Description, _
            vbExclamation, "押印"
End Sub

'--------------------------------------------------------------
' 電子印鑑シートに登録された担当者の印影を、押印先に反映する
'   （frmApprover の OK ボタンから呼ばれる）
'--------------------------------------------------------------
Sub ApplyApproverStamp(personName As String)

    Dim wsStamp As Worksheet
    Dim wsTarget As Worksheet
    Dim stampShape As Shape
    Dim newShape As Shape
    Dim d As Double
    Dim scaleRatio As Double

    If mTargetArea Is Nothing Then
        ShowMsg "押印先が選ばれていません。押印ボタンからやり直してください。", vbExclamation, "押印"
        Exit Sub
    End If

    Set wsStamp = ThisWorkbook.Worksheets(STAMP_SHEET)
    Set wsTarget = mTargetArea.Worksheet

    On Error Resume Next
    Set stampShape = wsStamp.Shapes(STAMP_SHAPE_PREFIX & personName)
    On Error GoTo 0

    If stampShape Is Nothing Then
        ShowMsg "「" & personName & "」の印影が電子印鑑シートに見つかりません。" & vbCrLf & _
                "電子印鑑シートで「電子印鑑生成」を押してください。", vbExclamation, "押印"
        Exit Sub
    End If

    ' 既存の印があれば削除してから貼り直す
    If ShapeExists(wsTarget, mTargetShapeName) Then wsTarget.Shapes(mTargetShapeName).Delete

    stampShape.Copy
    wsTarget.Paste
    Set newShape = wsTarget.Shapes(wsTarget.Shapes.Count)

    With newShape
        .Name = mTargetShapeName
        d = mTargetArea.Height * 0.9
        If d > mTargetArea.Width * 0.9 Then d = mTargetArea.Width * 0.9
        scaleRatio = d / stampShape.Width
        .LockAspectRatio = msoFalse
        .Width = d
        .Height = d
        .Left = mTargetArea.Left + (mTargetArea.Width - d) / 2
        .Top = mTargetArea.Top + (mTargetArea.Height - d) / 2
        .TextFrame.Characters.Font.Size = stampShape.TextFrame.Characters.Font.Size * scaleRatio
    End With

    mTargetArea.Worksheet.Activate
    mTargetArea.Cells(1, 1).Select

End Sub

'--------------------------------------------------------------
' 押印されている氏名（印が無ければ空文字）
'--------------------------------------------------------------
Public Function GetStampName(ws As Worksheet, shapeName As String) As String
    If Not ShapeExists(ws, shapeName) Then Exit Function
    On Error Resume Next
    GetStampName = Trim(ws.Shapes(shapeName).TextFrame2.TextRange.Text)
    On Error GoTo 0
End Function

'--------------------------------------------------------------
' 電子印鑑シートのB3:B10に入力されている氏名のうち、
' まだ印影（Stamp_氏名）が作成されていない氏名について、
' 決まった書式（下記の印影スタイル定数）で新しく印影を生成する。
' ・生成前に、現在のB3:B10に存在しない氏名の印影（担当者名を
' 　変更する前の古い印影）を自動的に削除する。
' ・文字は中央揃え。既定サイズ（40pt角）に1行で収まらない場合は
' 　フォントを8ptまで段階的に縮小し、それでも収まらなければ
' 　円を大きくする。
' （既存の QA試験.xlsm の電子印鑑生成と同じ）
'--------------------------------------------------------------
Sub 電子印鑑生成()

    Const STAMP_NAME_COL As Long = 2    ' B列：氏名
    Const STAMP_TARGET_COL As Long = 3  ' C列：印影の配置列
    Const STAMP_FIRST_ROW As Long = 3
    Const STAMP_LAST_ROW As Long = 10
    Const BASE_STAMP_SIZE As Double = 40  ' 印影の基準サイズ（円の大きさ）
    Const BASE_FONT_SIZE As Double = 12   ' 基準フォントサイズ
    Const FONT_SIZE_FLOOR As Double = 8   ' フォントを縮小する下限
    Const PER_CHAR_FACTOR As Double = 1.48  ' 1文字・1ptあたりの必要幅（実測から算出、安全マージン込み）
    Const MAX_STAMP_SIZE As Double = 80   ' 円の最大サイズ

    Const STAMP_FONT_NAME As String = "メイリオ"
    Const STAMP_LINE_WEIGHT As Double = 2.25

    Dim ws As Worksheet
    Dim newShape As Shape
    Dim targetCell As Range
    Dim r As Long
    Dim personName As String
    Dim boxSize As Double, fontSize As Double, displayText As String
    Dim createdCount As Long
    Dim skippedCount As Long
    Dim removedCount As Long

    Set ws = ThisWorkbook.Worksheets(STAMP_SHEET)

    ' 現在のB3:B10に存在しない氏名の印影（古い担当者名の印影）を先に削除する
    removedCount = RemoveOrphanStamps(ws, STAMP_NAME_COL, STAMP_FIRST_ROW, STAMP_LAST_ROW)

    For r = STAMP_FIRST_ROW To STAMP_LAST_ROW
        personName = Trim(ws.Cells(r, STAMP_NAME_COL).Value)
        If personName <> "" Then
            If ShapeExists(ws, STAMP_SHAPE_PREFIX & personName) Then
                skippedCount = skippedCount + 1
            Else
                DetermineStampLayout personName, BASE_STAMP_SIZE, BASE_FONT_SIZE, _
                    FONT_SIZE_FLOOR, PER_CHAR_FACTOR, MAX_STAMP_SIZE, _
                    boxSize, fontSize, displayText

                ' 印影が収まるよう、生成する行だけ高さを調整する
                If ws.Rows(r).RowHeight < boxSize Then
                    ws.Rows(r).RowHeight = boxSize
                End If

                Set targetCell = ws.Cells(r, STAMP_TARGET_COL)

                Set newShape = ws.Shapes.AddShape(msoShapeOval, _
                    targetCell.Left + (targetCell.Width - boxSize) / 2, _
                    targetCell.Top + (targetCell.Height - boxSize) / 2, boxSize, boxSize)

                With newShape
                    .Name = STAMP_SHAPE_PREFIX & personName
                    .Rotation = 0
                    .Fill.Visible = msoFalse
                    .Line.Visible = msoTrue
                    .Line.ForeColor.RGB = RGB(192, 0, 0)
                    .Line.Weight = STAMP_LINE_WEIGHT
                    .TextFrame.HorizontalAlignment = xlCenter
                    .TextFrame.VerticalAlignment = xlCenter
                    .TextFrame.MarginLeft = 0
                    .TextFrame.MarginRight = 0
                    .TextFrame.MarginTop = 0
                    .TextFrame.MarginBottom = 0
                    With .TextFrame.Characters
                        .Text = displayText
                        .Font.Name = STAMP_FONT_NAME
                        .Font.Size = fontSize
                        .Font.Bold = True
                        .Font.Italic = False
                        .Font.Color = RGB(192, 0, 0)
                    End With
                End With

                createdCount = createdCount + 1
            End If
        End If
    Next r

    ShowMsg createdCount & "件の印影を新規作成しました。" & vbCrLf & _
            "（作成済みの" & skippedCount & "件はそのまま保持、" & _
            "古い担当者名の印影" & removedCount & "件を削除）", vbInformation, "電子印鑑生成"

End Sub

'--------------------------------------------------------------
' 電子印鑑シート上の全ての印影（Stamp_〇〇）を削除する。
'--------------------------------------------------------------
Sub 電子印鑑削除()

    Dim ws As Worksheet
    Dim shp As Shape
    Dim stampCount As Long

    Set ws = ThisWorkbook.Worksheets(STAMP_SHEET)

    For Each shp In ws.Shapes
        If Left(shp.Name, Len(STAMP_SHAPE_PREFIX)) = STAMP_SHAPE_PREFIX Then
            stampCount = stampCount + 1
        End If
    Next shp

    If stampCount = 0 Then
        ShowMsg "削除できる印影がありません。", vbInformation, "電子印鑑削除"
        Exit Sub
    End If

    If ShowMsg("シート上の印影を全て削除します（" & stampCount & "件）。" & vbCrLf & _
               "よろしいですか？", vbYesNo + vbExclamation, "電子印鑑削除") = vbNo Then Exit Sub

    DeleteShapesByPrefix ws, STAMP_SHAPE_PREFIX

    ShowMsg stampCount & "件の印影を削除しました。", vbInformation, "電子印鑑削除"

End Sub

'--------------------------------------------------------------
' 電子印鑑シート上の印影（Stamp_〇〇）のうち、
' nameCol列のfirstRow～lastRowに存在しない氏名のものを削除する。
' 削除した件数を返す。
'--------------------------------------------------------------
Private Function RemoveOrphanStamps(ws As Worksheet, nameCol As Long, firstRow As Long, lastRow As Long) As Long

    Dim currentNames As String
    Dim personName As String
    Dim r As Long
    Dim i As Long
    Dim shp As Shape
    Dim removed As Long

    currentNames = "|"
    For r = firstRow To lastRow
        personName = Trim(ws.Cells(r, nameCol).Value)
        If personName <> "" Then
            currentNames = currentNames & personName & "|"
        End If
    Next r

    For i = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(i)
        If Left(shp.Name, Len(STAMP_SHAPE_PREFIX)) = STAMP_SHAPE_PREFIX Then
            personName = Mid(shp.Name, Len(STAMP_SHAPE_PREFIX) + 1)
            If InStr(currentNames, "|" & personName & "|") = 0 Then
                shp.Delete
                removed = removed + 1
            End If
        End If
    Next i

    RemoveOrphanStamps = removed

End Function

'--------------------------------------------------------------
' 氏名から、印影のサイズ・フォントサイズ・表示文字列を決定する。
' 1行で基準サイズに収まるフォントを優先し、フォントを下限まで縮小しても
' 収まらない場合は円を大きくする。
'--------------------------------------------------------------
Private Sub DetermineStampLayout(personName As String, baseSize As Double, baseFont As Double, _
    floorFont As Double, perCharFactor As Double, maxSize As Double, _
    ByRef outBoxSize As Double, ByRef outFontSize As Double, ByRef outDisplayText As String)

    Dim n As Long
    Dim size As Double

    n = Len(personName)
    outDisplayText = personName

    If n <= 2 Then
        outBoxSize = baseSize
        outFontSize = baseFont
        Exit Sub
    End If

    size = baseSize / (n * perCharFactor)
    If size > baseFont Then size = baseFont
    size = Int(size * 2) / 2  ' 0.5pt単位で切り捨て（安全側）

    If size >= floorFont Then
        outBoxSize = baseSize
        outFontSize = size
    Else
        outFontSize = floorFont
        outBoxSize = n * perCharFactor * floorFont
        If outBoxSize < baseSize Then outBoxSize = baseSize
        If outBoxSize > maxSize Then outBoxSize = maxSize
    End If

End Sub
