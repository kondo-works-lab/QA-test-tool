Attribute VB_Name = "modSave"
Option Explicit

'==============================================================
' QA試験 保存
'   表紙と明細シートを1ファイルにまとめ、Excel（.xlsx）と PDF のうち選んだ形式で保存する。
'   ・「保存」ボタンで保存形式を選ぶフォーム（frmSaveType）を出す。両方選んでもよい（NG のときは出さない）。
'   ・保存先フォルダは保存のたびにダイアログで選ぶ（初期表示はこのブックのフォルダ）。
'     Excel はその中の「Excel」フォルダ、PDF は「PDF」フォルダに保存する（無ければ作る）。
'   ・ファイル名：品名_ロット_日付（yyyymmdd）
'     表紙の総合判定が NG のときは、先頭に【NG】を付ける（Excel・PDF とも）。
'     同じ品名・ロットの NG の保存が2回目以降なら、末尾に「_2回目」「_3回目」…を付ける
'     （回数は保存履歴の、ファイル名が【NG】で始まる Excel 保存の数で数える）。
'   ・総合判定が NG のときは、形式を選ばずに Excel・PDF の両方を保存し、その後に
'     NG メールの下書き（Gmail の作成画面）を開き（modMail）、保存履歴に記録する。送信はしない。
'   ・ボタンは保存ファイルから除く。
'   ・PDF は各シート A4 横・1ページに収めて出力する（元のシートの印刷設定は変更しない）。
'   ・保存前に「全検査の判定」「印鑑」を確認し、不足があれば保存してよいか確認する。
'   ・保存に成功したら「保存履歴」シートに記録し、このブック（xlsm）を上書き保存する。
'==============================================================

Public Const HISTORY_SHEET_NAME As String = "保存履歴"
Public Const HIST_TYPE_EXCEL As String = "Excel"
Public Const HIST_TYPE_PDF As String = "PDF"
Public Const HIST_TYPE_MAIL As String = "メール下書き"
Public Const HIST_TYPE_DB As String = "DB"
Public Const NG_FILE_PREFIX As String = "【NG】"
Public Const EXCEL_FOLDER_NAME As String = "Excel"   ' 選んだ保存先の中に作る Excel 用フォルダ
Public Const PDF_FOLDER_NAME As String = "PDF"       ' 選んだ保存先の中に作る PDF 用フォルダ

Private Const HIST_COL_TIME As Long = 1      ' A 保存日時
Private Const HIST_COL_TYPE As Long = 2      ' B 種類
Private Const HIST_COL_PRODUCT As Long = 3   ' C 品名
Private Const HIST_COL_LOT As Long = 4       ' D ロット
Private Const HIST_COL_PATH As Long = 6      ' F 保存先

Public gTestSaveFolder As String   ' テスト用：設定するとフォルダ選択ダイアログを出さずにこのフォルダへ保存する
Public gTestSaveTypes As String    ' テスト用：設定するとフォームを出さずにこの形式で保存する（"Excel" / "PDF" / "Excel,PDF"）

' 保存形式を選ぶフォーム（frmSaveType）の結果
Private mSaveChosen As Boolean
Private mSaveExcel As Boolean
Private mSavePdf As Boolean

'--------------------------------------------------------------
' 「保存」ボタン：Excel／PDF を選んで保存する
'   総合判定が NG のときは選ばずに Excel・PDF の両方を保存し、NG メールの下書きを開く
'--------------------------------------------------------------
Sub 検査表保存()

    Const DIALOG_TITLE As String = "保存"
    Dim cs As Worksheet
    Dim wantExcel As Boolean
    Dim wantPdf As Boolean
    Dim folderPath As String
    Dim excelFolder As String
    Dim pdfFolder As String
    Dim prefix As String
    Dim suffix As String
    Dim excelPath As String
    Dim pdfPath As String
    Dim isNg As Boolean
    Dim ngCount As Long
    Dim excelSaved As Boolean
    Dim pdfSaved As Boolean
    Dim mailTo As String
    Dim mailErr As String
    Dim savedText As String
    Dim failedText As String
    Dim mailText As String

    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)
    isNg = (GetOverall() = RESULT_NG)

    If isNg Then
        wantExcel = True    ' NG のときは必ず両方保存する（PDF をメールに添付するため）
        wantPdf = True
    ElseIf Not SelectSaveTypes(wantExcel, wantPdf) Then
        Exit Sub
    End If
    If Not ConfirmReadyToSave(DIALOG_TITLE) Then Exit Sub
    If Not CheckFileNameParts(DIALOG_TITLE) Then Exit Sub
    If Not SelectSaveFolder(folderPath) Then Exit Sub

    ' NG：先頭に【NG】、2回目以降は末尾に「_2回目」「_3回目」…
    If isNg Then
        ngCount = CountNgSaves(CellText(cs.Range(COVER_PRODUCT_CELL).Value), CellText(cs.Range(COVER_LOT_CELL).Value)) + 1
        prefix = NG_FILE_PREFIX
        suffix = NgCountSuffix(ngCount)
    End If

    ' 上書き確認をすべて済ませてから保存する（途中で中止したら何も保存しない）
    ' Excel は「Excel」フォルダ、PDF は「PDF」フォルダへ
    If wantExcel Then
        If Not PrepareSubFolder(folderPath, EXCEL_FOLDER_NAME, excelFolder, DIALOG_TITLE) Then Exit Sub
        If Not BuildSavePath(excelFolder, ".xlsx", DIALOG_TITLE, prefix, suffix, excelPath) Then Exit Sub
    End If
    If wantPdf Then
        If Not PrepareSubFolder(folderPath, PDF_FOLDER_NAME, pdfFolder, DIALOG_TITLE) Then Exit Sub
        If Not BuildSavePath(pdfFolder, ".pdf", DIALOG_TITLE, prefix, suffix, pdfPath) Then Exit Sub
    End If

    If wantExcel Then
        excelSaved = SaveAsExcel(excelPath, DIALOG_TITLE)
        If excelSaved Then
            WriteSaveHistory HIST_TYPE_EXCEL, excelPath
            SetInputChanged False   ' DB 登録はこの保存の内容で行える
            savedText = savedText & excelPath & vbCrLf
        Else
            failedText = failedText & "Excel "
        End If
    End If
    If wantPdf Then
        pdfSaved = SaveAsPdf(pdfPath, DIALOG_TITLE)
        If pdfSaved Then
            WriteSaveHistory HIST_TYPE_PDF, pdfPath
            savedText = savedText & pdfPath & vbCrLf
        Else
            failedText = failedText & "PDF "
        End If
    End If

    ' NG のときは保存の後に、NG メールの下書き（Gmail の作成画面）を開く
    If isNg Then
        If Not excelSaved Then excelPath = ""
        If Not pdfSaved Then pdfPath = ""
        If CreateNgMailDraft(excelPath, pdfPath, ngCount, mailTo, mailErr) Then
            WriteSaveHistory HIST_TYPE_MAIL, mailTo
            mailText = vbCrLf & "NG メール（" & ngCount & "回目）の下書きを Gmail で開きました。（宛先：" & mailTo & "）"
            If pdfSaved Then
                mailText = mailText & vbCrLf & "開いたフォルダの PDF を作成画面にドラッグして添付し、内容を確認して送信してください。"
            Else
                mailText = mailText & vbCrLf & "※ PDF を保存できなかったため、添付するファイルがありません。"
            End If
        Else
            mailText = vbCrLf & "※ NG メールの下書きを作れませんでした。" & vbCrLf & mailErr
        End If
    End If

    SaveThisWorkbook   ' 保存履歴を残す

    If failedText = "" Then
        ShowMsg "保存しました。" & vbCrLf & savedText & mailText, vbInformation, DIALOG_TITLE
    ElseIf savedText <> "" Then
        ShowMsg "一部だけ保存しました。（保存できなかった形式：" & Trim(failedText) & "）" & vbCrLf & _
                savedText & mailText, vbExclamation, DIALOG_TITLE
    ElseIf mailText <> "" Then
        ShowMsg "保存できませんでした。" & mailText, vbExclamation, DIALOG_TITLE
    End If

End Sub

'--------------------------------------------------------------
' NG 回数の付け足し（1回目は無し、2回目以降は「_2回目」）
'--------------------------------------------------------------
Public Function NgCountSuffix(ngCount As Long) As String
    If ngCount >= 2 Then NgCountSuffix = "_" & ngCount & "回目"
End Function

'--------------------------------------------------------------
' ファイル名（パスの最後）が【NG】で始まるか
'--------------------------------------------------------------
Public Function IsNgFileName(savePath As String) As Boolean
    Dim fileName As String
    fileName = Mid(savePath, InStrRev(savePath, "\") + 1)
    IsNgFileName = (Left(fileName, Len(NG_FILE_PREFIX)) = NG_FILE_PREFIX)
End Function

'--------------------------------------------------------------
' 保存形式を選ぶ。キャンセルしたら False
'   フォームが無くてもブック全体がコンパイルエラーにならないよう、名前で呼び出す
'--------------------------------------------------------------
Private Function SelectSaveTypes(ByRef wantExcel As Boolean, ByRef wantPdf As Boolean) As Boolean

    If gTestSaveTypes <> "" Then
        wantExcel = (InStr(gTestSaveTypes, HIST_TYPE_EXCEL) > 0)
        wantPdf = (InStr(gTestSaveTypes, HIST_TYPE_PDF) > 0)
        SelectSaveTypes = (wantExcel Or wantPdf)
        Exit Function
    End If

    mSaveChosen = False
    mSaveExcel = False
    mSavePdf = False

    On Error GoTo ErrHandler
    VBA.UserForms.Add("frmSaveType").Show
    On Error GoTo 0

    wantExcel = mSaveExcel
    wantPdf = mSavePdf
    SelectSaveTypes = mSaveChosen And (wantExcel Or wantPdf)
    Exit Function

ErrHandler:
    ShowMsg "保存形式を選ぶフォーム（frmSaveType）を表示できません。" & vbCrLf & _
            "マクロ「フォーム作成」を実行してから、もう一度押してください。" & vbCrLf & Err.Description, _
            vbExclamation, "保存"

End Function

'--------------------------------------------------------------
' frmSaveType の OK ボタンから呼ばれる
'--------------------------------------------------------------
Public Sub SetSaveTypes(wantExcel As Boolean, wantPdf As Boolean)
    mSaveChosen = True
    mSaveExcel = wantExcel
    mSavePdf = wantPdf
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
    dlg.title = "保存先フォルダを選択してください"
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
' 保存先の中のフォルダ（Excel／PDF）を返す。無ければ作る。作れなければ False
'--------------------------------------------------------------
Private Function PrepareSubFolder(folderPath As String, subName As String, _
    ByRef subPath As String, title As String) As Boolean

    If Right(folderPath, 1) = "\" Then
        subPath = folderPath & subName
    Else
        subPath = folderPath & "\" & subName
    End If
    If Dir(subPath, vbDirectory) <> "" Then
        PrepareSubFolder = True
        Exit Function
    End If

    On Error GoTo ErrHandler
    MkDir subPath
    PrepareSubFolder = True
    Exit Function

ErrHandler:
    ShowMsg "保存用のフォルダを作れませんでした。" & vbCrLf & subPath & vbCrLf & Err.Description, vbExclamation, title

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
' 保存先のフルパスを作る（先頭に prefix、末尾に suffix を付ける）。同名ファイルがある場合は上書きするか確認する。
'--------------------------------------------------------------
Private Function BuildSavePath(folderPath As String, extension As String, _
    title As String, prefix As String, suffix As String, ByRef outPath As String) As Boolean

    Dim baseName As String
    Dim errText As String

    BuildSavePath = False
    If Not BuildFileBaseName(baseName, errText) Then Exit Function

    If Right(folderPath, 1) = "\" Then
        outPath = folderPath & prefix & baseName & suffix & extension
    Else
        outPath = folderPath & "\" & prefix & baseName & suffix & extension
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
Public Sub WriteSaveHistory(fileType As String, savePath As String)

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

    histWs.Cells(nextRow, HIST_COL_TIME).Value = Now
    histWs.Cells(nextRow, HIST_COL_TYPE).Value = fileType
    histWs.Cells(nextRow, HIST_COL_PRODUCT).Value = CellText(cs.Range(COVER_PRODUCT_CELL).Value)
    histWs.Cells(nextRow, HIST_COL_LOT).Value = CellText(cs.Range(COVER_LOT_CELL).Value)
    histWs.Cells(nextRow, 5).Value = approverName
    histWs.Cells(nextRow, HIST_COL_PATH).Value = savePath
    Exit Sub

ErrHandler:
    ShowMsg "保存履歴の記録に失敗しました。（ファイルは保存済みです）" & vbCrLf & Err.Description, _
            vbExclamation, "保存履歴"

End Sub

'--------------------------------------------------------------
' 保存履歴の、品名・ロットが同じ NG の Excel 保存（ファイル名が【NG】で始まる）の数
'--------------------------------------------------------------
Public Function CountNgSaves(productName As String, lotNumber As String) As Long

    Dim histWs As Worksheet
    Dim r As Long

    Set histWs = FindSheet(HISTORY_SHEET_NAME)
    If histWs Is Nothing Then Exit Function

    For r = 2 To histWs.Cells(histWs.Rows.Count, 1).End(xlUp).Row
        If IsExcelSaveOf(histWs, r, productName, lotNumber) Then
            If IsNgFileName(CellText(histWs.Cells(r, HIST_COL_PATH).Value)) Then CountNgSaves = CountNgSaves + 1
        End If
    Next r

End Function

'--------------------------------------------------------------
' 保存履歴から、品名・ロットが同じ最新の Excel 保存を探す。見つかったら True
'   saveTime：保存日時
'   ngCount ：その保存が NG なら何回目の NG か（1～）、OK の保存なら 0
'--------------------------------------------------------------
Public Function GetLatestExcelSave(productName As String, lotNumber As String, _
    ByRef saveTime As Date, ByRef ngCount As Long) As Boolean

    Dim histWs As Worksheet
    Dim r As Long
    Dim latestRow As Long
    Dim v As Variant

    Set histWs = FindSheet(HISTORY_SHEET_NAME)
    If histWs Is Nothing Then Exit Function

    For r = histWs.Cells(histWs.Rows.Count, 1).End(xlUp).Row To 2 Step -1
        If IsExcelSaveOf(histWs, r, productName, lotNumber) Then
            v = histWs.Cells(r, HIST_COL_TIME).Value
            If Not IsError(v) Then
                If IsDate(v) Then
                    latestRow = r
                    saveTime = CDate(v)
                    Exit For
                End If
            End If
        End If
    Next r
    If latestRow = 0 Then Exit Function

    ngCount = 0
    If IsNgFileName(CellText(histWs.Cells(latestRow, HIST_COL_PATH).Value)) Then
        For r = 2 To latestRow
            If IsExcelSaveOf(histWs, r, productName, lotNumber) Then
                If IsNgFileName(CellText(histWs.Cells(r, HIST_COL_PATH).Value)) Then ngCount = ngCount + 1
            End If
        Next r
    End If
    GetLatestExcelSave = True

End Function

Private Function IsExcelSaveOf(histWs As Worksheet, r As Long, productName As String, lotNumber As String) As Boolean
    IsExcelSaveOf = (CellText(histWs.Cells(r, HIST_COL_TYPE).Value) = HIST_TYPE_EXCEL _
        And CellText(histWs.Cells(r, HIST_COL_PRODUCT).Value) = productName _
        And CellText(histWs.Cells(r, HIST_COL_LOT).Value) = lotNumber)
End Function

'--------------------------------------------------------------
' 保存履歴を失わないよう、このブック（xlsm）を上書き保存する。
'--------------------------------------------------------------
Public Sub SaveThisWorkbook()

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
