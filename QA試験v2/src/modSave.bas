Attribute VB_Name = "modSave"
Option Explicit

'==============================================================
' QA試験 保存（既存の QA試験.xlsm の保存機能を流用）
'   表紙と明細シートを1ファイルにまとめ、Excel（.xlsx）または PDF で保存する。
'   ・保存先フォルダは保存のたびにダイアログで選ぶ（初期表示はこのブックのフォルダ）。
'   ・ファイル名：品名_ロット_日付（yyyymmdd）
'   ・ボタンは保存ファイルから除く。
'   ・PDF は各シート A4 横・1ページに収めて出力する（元のシートの印刷設定は変更しない）。
'   ・保存前に「全検査の判定」「印鑑」を確認し、不足があれば保存してよいか確認する。
'   ・保存に成功したら「保存履歴」シートに記録し、このブック（xlsm）を上書き保存する。
'==============================================================

Private Const HISTORY_SHEET_NAME As String = "保存履歴"

Public gTestSaveFolder As String   ' テスト用：設定するとフォルダ選択ダイアログを出さずにこのフォルダへ保存する

'--------------------------------------------------------------
' Excel（.xlsx）で保存する
'--------------------------------------------------------------
Sub 検査表Excel保存()

    Const DIALOG_TITLE As String = "Excel保存"
    Dim folderPath As String
    Dim savePath As String

    If Not ConfirmReadyToSave(DIALOG_TITLE) Then Exit Sub
    If Not CheckFileNameParts(DIALOG_TITLE) Then Exit Sub
    If Not SelectSaveFolder(folderPath) Then Exit Sub
    If Not BuildSavePath(folderPath, ".xlsx", DIALOG_TITLE, savePath) Then Exit Sub

    If SaveAsExcel(savePath, DIALOG_TITLE) Then
        WriteSaveHistory "Excel", savePath
        SaveThisWorkbook
        ShowMsg "保存しました。" & vbCrLf & savePath, vbInformation, DIALOG_TITLE
    End If

End Sub

'--------------------------------------------------------------
' PDF で保存する
'--------------------------------------------------------------
Sub 検査表PDF保存()

    Const DIALOG_TITLE As String = "PDF保存"
    Dim folderPath As String
    Dim savePath As String

    If Not ConfirmReadyToSave(DIALOG_TITLE) Then Exit Sub
    If Not CheckFileNameParts(DIALOG_TITLE) Then Exit Sub
    If Not SelectSaveFolder(folderPath) Then Exit Sub
    If Not BuildSavePath(folderPath, ".pdf", DIALOG_TITLE, savePath) Then Exit Sub

    If SaveAsPdf(savePath, DIALOG_TITLE) Then
        WriteSaveHistory "PDF", savePath
        SaveThisWorkbook
        ShowMsg "保存しました。" & vbCrLf & savePath, vbInformation, DIALOG_TITLE
    End If

End Sub

'--------------------------------------------------------------
' Excel（.xlsx）と PDF の両方を、選んだ同じフォルダに保存する
'   上書き確認を両方とも済ませてから保存する。どちらかで中止した場合はどちらも保存しない。
'--------------------------------------------------------------
Sub 検査表一括保存()

    Const DIALOG_TITLE As String = "一括保存"
    Dim folderPath As String
    Dim excelPath As String
    Dim pdfPath As String

    If Not ConfirmReadyToSave(DIALOG_TITLE) Then Exit Sub
    If Not CheckFileNameParts(DIALOG_TITLE) Then Exit Sub
    If Not SelectSaveFolder(folderPath) Then Exit Sub
    If Not BuildSavePath(folderPath, ".xlsx", DIALOG_TITLE, excelPath) Then Exit Sub
    If Not BuildSavePath(folderPath, ".pdf", DIALOG_TITLE, pdfPath) Then Exit Sub

    If Not SaveAsExcel(excelPath, DIALOG_TITLE) Then Exit Sub
    WriteSaveHistory "Excel", excelPath

    If Not SaveAsPdf(pdfPath, DIALOG_TITLE) Then
        SaveThisWorkbook   ' Excel 分の履歴を残す
        ShowMsg "Excel のみ保存しました。" & vbCrLf & excelPath, vbExclamation, DIALOG_TITLE
        Exit Sub
    End If
    WriteSaveHistory "PDF", pdfPath
    SaveThisWorkbook

    ShowMsg "保存しました。" & vbCrLf & excelPath & vbCrLf & pdfPath, vbInformation, DIALOG_TITLE

End Sub

'--------------------------------------------------------------
' 保存先フォルダを選ぶ。キャンセルしたら False
'--------------------------------------------------------------
Private Function SelectSaveFolder(ByRef folderPath As String) As Boolean

    Dim dlg As Object

    If gTestSaveFolder <> "" Then
        folderPath = gTestSaveFolder
        SelectSaveFolder = True
        Exit Function
    End If

    Set dlg = Application.FileDialog(msoFileDialogFolderPicker)
    dlg.Title = "保存先フォルダを選択してください"
    dlg.AllowMultiSelect = False
    If ThisWorkbook.Path <> "" Then dlg.InitialFileName = ThisWorkbook.Path & "\"
    If dlg.Show <> -1 Then
        SelectSaveFolder = False
        Exit Function
    End If

    folderPath = dlg.SelectedItems(1)
    SelectSaveFolder = True

End Function

'--------------------------------------------------------------
' 全検査の判定と印鑑の有無を確認する。
' 不足がある場合は保存してよいか確認し、保存を続行する場合のみ True を返す。
'--------------------------------------------------------------
Private Function ConfirmReadyToSave(title As String) As Boolean

    Dim cs As Worksheet
    Dim missingText As String
    Dim r As Long
    Dim noInspStamp As String

    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)

    If CoverListCount() = 0 Then
        ShowMsg "明細シートがありません。表紙の品名を入力してください。", vbExclamation, title
        Exit Function
    End If

    If Not AllInspectionsJudged() Then missingText = missingText & "・判定が済んでいない検査があります" & vbCrLf
    If Not ShapeExists(cs, APPROVER_STAMP_NAME) Then missingText = missingText & "・承認者印がありません" & vbCrLf
    If Not ShapeExists(cs, CHARGE_STAMP_NAME) Then missingText = missingText & "・担当者印がありません" & vbCrLf
    For r = COVER_LIST_FIRST_ROW To COVER_LIST_FIRST_ROW + CoverListCount() - 1
        If Not ShapeExists(cs, INSP_STAMP_PREFIX & CellText(cs.Cells(r, COVER_COL_NAME).Value)) Then
            If noInspStamp <> "" Then noInspStamp = noInspStamp & "、"
            noInspStamp = noInspStamp & CellText(cs.Cells(r, COVER_COL_NAME).Value)
        End If
    Next r
    If noInspStamp <> "" Then missingText = missingText & "・検査担当者印がありません：" & noInspStamp & vbCrLf

    If missingText = "" Then
        ConfirmReadyToSave = True
    Else
        ConfirmReadyToSave = (ShowMsg(missingText & vbCrLf & "このまま保存しますか？", _
                              vbYesNo + vbExclamation + vbDefaultButton2, title) = vbYes)
    End If

End Function

'--------------------------------------------------------------
' ファイル名（品名_ロット_日付）を作る。問題があれば False
'--------------------------------------------------------------
Public Function BuildFileBaseName(ByRef baseName As String, ByRef errText As String) As Boolean

    Dim cs As Worksheet
    Dim productName As String
    Dim lotNumber As String
    Dim dateValue As Variant

    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)
    productName = CellText(cs.Range(COVER_PRODUCT_CELL).Value)
    lotNumber = CellText(cs.Range(COVER_LOT_CELL).Value)
    dateValue = cs.Range(COVER_DATE_CELL).Value

    If productName = "" Then errText = errText & "・品名（" & COVER_PRODUCT_CELL & "）が入力されていません" & vbCrLf
    If lotNumber = "" Then errText = errText & "・ロット（" & COVER_LOT_CELL & "）が入力されていません" & vbCrLf
    If IsError(dateValue) Then
        errText = errText & "・日付（" & COVER_DATE_CELL & "）がエラー値です" & vbCrLf
    ElseIf IsBlankValue(dateValue) Then
        errText = errText & "・日付（" & COVER_DATE_CELL & "）が入力されていません" & vbCrLf
    ElseIf Not IsDate(dateValue) Then
        errText = errText & "・日付（" & COVER_DATE_CELL & "）を日付として読み取れません" & vbCrLf
    End If
    If HasInvalidFileChar(productName) Then errText = errText & "・品名に使用できない文字（\ / : * ? "" < > |）が含まれています" & vbCrLf
    If HasInvalidFileChar(lotNumber) Then errText = errText & "・ロットに使用できない文字（\ / : * ? "" < > |）が含まれています" & vbCrLf

    If errText <> "" Then
        BuildFileBaseName = False
        Exit Function
    End If

    baseName = productName & "_" & lotNumber & "_" & Format(CDate(dateValue), "yyyymmdd")
    BuildFileBaseName = True

End Function

Private Function CheckFileNameParts(title As String) As Boolean
    Dim baseName As String
    Dim errText As String
    If BuildFileBaseName(baseName, errText) Then
        CheckFileNameParts = True
    Else
        ShowMsg "保存できません。" & vbCrLf & errText, vbExclamation, title
    End If
End Function

'--------------------------------------------------------------
' 保存先のフルパスを作る。同名ファイルがある場合は上書きするか確認する。
'--------------------------------------------------------------
Private Function BuildSavePath(folderPath As String, extension As String, _
    title As String, ByRef outPath As String) As Boolean

    Dim baseName As String
    Dim errText As String

    BuildSavePath = False
    If Not BuildFileBaseName(baseName, errText) Then Exit Function

    If Right(folderPath, 1) = "\" Then
        outPath = folderPath & baseName & extension
    Else
        outPath = folderPath & "\" & baseName & extension
    End If

    If Dir(outPath) <> "" Then
        If ShowMsg("同じ名前のファイルが既にあります。上書きしますか？" & vbCrLf & outPath, _
                   vbYesNo + vbQuestion, title) = vbNo Then
            Exit Function
        End If
    End If

    BuildSavePath = True

End Function

'--------------------------------------------------------------
' 保存履歴シートに1行追加する（保存日時・種類・品名・ロット・承認者・保存先）
' 記録に失敗しても保存自体は完了しているため、メッセージのみ表示する。
'--------------------------------------------------------------
Private Sub WriteSaveHistory(fileType As String, savePath As String)

    Dim cs As Worksheet
    Dim histWs As Worksheet
    Dim nextRow As Long
    Dim approverName As String

    On Error GoTo ErrHandler

    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)
    Set histWs = GetHistorySheet()
    nextRow = histWs.Cells(histWs.Rows.Count, 1).End(xlUp).Row + 1

    approverName = GetStampName(cs, APPROVER_STAMP_NAME)
    If approverName = "" Then approverName = "（なし）"

    histWs.Cells(nextRow, 1).Value = Now
    histWs.Cells(nextRow, 2).Value = fileType
    histWs.Cells(nextRow, 3).Value = CellText(cs.Range(COVER_PRODUCT_CELL).Value)
    histWs.Cells(nextRow, 4).Value = CellText(cs.Range(COVER_LOT_CELL).Value)
    histWs.Cells(nextRow, 5).Value = approverName
    histWs.Cells(nextRow, 6).Value = savePath
    Exit Sub

ErrHandler:
    ShowMsg "保存履歴の記録に失敗しました。（ファイルは保存済みです）" & vbCrLf & Err.Description, _
            vbExclamation, "保存履歴"

End Sub

'--------------------------------------------------------------
' 保存履歴を失わないよう、このブック（xlsm）を上書き保存する。
'--------------------------------------------------------------
Private Sub SaveThisWorkbook()

    If gTestMode Then Exit Sub   ' テスト中はブックを保存しない

    If ThisWorkbook.ReadOnly Then
        ShowMsg "このブックは読み取り専用のため、保存履歴を上書き保存できません。" & vbCrLf & _
                "（ファイル自体は保存済みです）", vbExclamation, "保存履歴"
        Exit Sub
    End If

    On Error GoTo ErrHandler
    ThisWorkbook.Save
    Exit Sub

ErrHandler:
    ShowMsg "ブックの上書き保存に失敗しました。（ファイル自体は保存済みです）" & vbCrLf & Err.Description, _
            vbExclamation, "保存履歴"

End Sub

'--------------------------------------------------------------
' 保存履歴シートを返す。無い場合は見出し付きで最後尾に作成する
'--------------------------------------------------------------
Private Function GetHistorySheet() As Worksheet

    Dim histWs As Worksheet
    Dim activeWs As Object

    Set histWs = FindSheet(HISTORY_SHEET_NAME)
    If histWs Is Nothing Then
        Set activeWs = ActiveSheet
        Set histWs = AddSheetAfter(ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        histWs.Name = HISTORY_SHEET_NAME
        histWs.Range("A1:F1").Value = Array("保存日時", "種類", "品名", "ロット", "承認者", "保存先")
        histWs.Range("A1:F1").Font.Bold = True
        histWs.Columns("A").NumberFormat = "yyyy/mm/dd hh:mm:ss"
        histWs.Columns("C:D").NumberFormat = "@"   ' 品名・ロットは文字列のまま記録する
        histWs.Columns("A").ColumnWidth = 20
        histWs.Columns("B").ColumnWidth = 8
        histWs.Columns("C:E").ColumnWidth = 16
        histWs.Columns("F").ColumnWidth = 70
        activeWs.Activate   ' シート追加で切り替わった表示を戻す
    End If

    Set GetHistorySheet = histWs

End Function

'--------------------------------------------------------------
' .xlsx で保存する。成功したら True
'--------------------------------------------------------------
Private Function SaveAsExcel(savePath As String, title As String) As Boolean

    Dim outWb As Workbook
    Dim errMsg As String

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    Set outWb = CreateOutputWorkbook()

    Application.DisplayAlerts = False   ' 上書き確認は BuildSavePath で実施済み
    outWb.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True

    outWb.Close SaveChanges:=False
    Application.ScreenUpdating = True

    SaveAsExcel = True
    Exit Function

ErrHandler:
    errMsg = Err.Description
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    If Not outWb Is Nothing Then outWb.Close SaveChanges:=False
    ShowMsg "Excel の保存に失敗しました。" & vbCrLf & errMsg, vbExclamation, title

End Function

'--------------------------------------------------------------
' PDF で出力する。成功したら True
'--------------------------------------------------------------
Private Function SaveAsPdf(savePath As String, title As String) As Boolean

    Dim outWb As Workbook
    Dim outWs As Worksheet
    Dim errMsg As String

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    ' 元のシートのボタンに触れないよう、コピーから PDF を出力する
    Set outWb = CreateOutputWorkbook()
    For Each outWs In outWb.Worksheets
        SetPdfPageLayout outWs
    Next outWs
    outWb.ExportAsFixedFormat Type:=xlTypePDF, Filename:=savePath, _
        Quality:=xlQualityStandard, IgnorePrintAreas:=False, OpenAfterPublish:=False

    outWb.Close SaveChanges:=False
    Application.ScreenUpdating = True

    SaveAsPdf = True
    Exit Function

ErrHandler:
    errMsg = Err.Description
    Application.ScreenUpdating = True
    If Not outWb Is Nothing Then outWb.Close SaveChanges:=False
    ShowMsg "PDF の保存に失敗しました。" & vbCrLf & errMsg, vbExclamation, title

End Function

'--------------------------------------------------------------
' PDF 用のコピーを A4 横・1ページに収まる印刷設定にする
' プリンタが無い等で設定できない場合は、そのままの設定で出力する
'--------------------------------------------------------------
Private Sub SetPdfPageLayout(outWs As Worksheet)

    On Error Resume Next
    With outWs.PageSetup
        .PaperSize = xlPaperA4
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterHorizontally = True
    End With
    On Error GoTo 0

End Sub

'--------------------------------------------------------------
' 表紙と明細シートを新しいブックにコピーし、ボタンを削除して返す
'--------------------------------------------------------------
Private Function CreateOutputWorkbook() As Workbook

    Dim names() As Variant
    Dim item As Variant
    Dim cnt As Long
    Dim outWb As Workbook
    Dim outWs As Worksheet
    Dim i As Long

    ReDim names(0 To GetDetailSheets().Count)
    names(0) = COVER_SHEET
    cnt = 1
    For Each item In GetDetailSheets()
        names(cnt) = item.Name
        cnt = cnt + 1
    Next item

    ThisWorkbook.Worksheets(names).Copy
    Set outWb = ActiveWorkbook

    ' ボタン（フォームコントロール）は保存ファイルでは使えないため削除する
    For Each outWs In outWb.Worksheets
        For i = outWs.Shapes.Count To 1 Step -1
            If outWs.Shapes(i).Type = msoFormControl Then
                outWs.Shapes(i).Delete
            End If
        Next i
    Next outWs
    outWb.Worksheets(1).Activate

    Set CreateOutputWorkbook = outWb

End Function

'--------------------------------------------------------------
' ファイル名に使用できない文字が含まれているか判定する
'--------------------------------------------------------------
Private Function HasInvalidFileChar(text As String) As Boolean
    Dim invalidChars As Variant
    Dim i As Long

    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = LBound(invalidChars) To UBound(invalidChars)
        If InStr(text, invalidChars(i)) > 0 Then
            HasInvalidFileChar = True
            Exit Function
        End If
    Next i
    HasInvalidFileChar = False
End Function
