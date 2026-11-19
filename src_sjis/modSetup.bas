Attribute VB_Name = "modSetup"
Option Explicit

'==============================================================
' QA試験 セットアップ
'   新しいブックに各モジュールと「検査項目一覧」シート（既存の QA試験.xlsm から
'   コピー）を入れた後、1回だけ実行する。
'   表紙・マスタ・電子印鑑・設定シートとボタンを作り、
'   検査項目一覧の内容をマスタに移してから、検査項目一覧シートを削除する。
'   明細シートは表紙の品名を入力したときに作られる。
'   ・既にセットアップ済みのブックでは「表紙ボタンと設定の更新」を実行すると、
'     表紙のボタンと設定シートだけを今の形に作り直す（入力・印鑑・マスタはそのまま）。
'==============================================================

Private Const COVER_BUTTON_PREFIX As String = "btnCover_"

Sub セットアップ()

    Dim ws As Worksheet
    Dim i As Long
    Dim alertsBefore As Boolean
    Dim errMsg As String
    Dim importMsg As String

    If Not FindSheet(COVER_SHEET) Is Nothing Then
        ShowMsg "既にセットアップ済みです（「" & COVER_SHEET & "」シートがあります）。", vbInformation, "セットアップ"
        Exit Sub
    End If

    On Error GoTo ErrHandler
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    BuildCoverSheet AddSheet(COVER_SHEET)
    BuildMasterSheet AddSheet(MASTER_SHEET)
    importMsg = ImportItemList()
    BuildStampSheet AddSheet(STAMP_SHEET)
    BuildSettingSheet AddSheet(SETTING_SHEET)

    ' 新しいブックに最初からある空のシート（Sheet1 等）を削除する
    alertsBefore = Application.DisplayAlerts
    Application.DisplayAlerts = False
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        Set ws = ThisWorkbook.Worksheets(i)
        Select Case ws.Name
            Case COVER_SHEET, MASTER_SHEET, STAMP_SHEET, SETTING_SHEET
            Case Else
                If Application.WorksheetFunction.CountA(ws.UsedRange) = 0 And ws.Shapes.Count = 0 Then ws.Delete
        End Select
    Next i
    Application.DisplayAlerts = alertsBefore

    ThisWorkbook.Worksheets(COVER_SHEET).Activate
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    ShowMsg "セットアップが完了しました。" & vbCrLf & importMsg & vbCrLf & _
            "電子印鑑シートで氏名・部署を確認してから、表紙の品名を入力してください。", _
            vbInformation, "セットアップ"
    Exit Sub

ErrHandler:
    errMsg = Err.Description
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ShowMsg "セットアップに失敗しました。" & vbCrLf & errMsg, vbExclamation, "セットアップ"

End Sub

Private Function AddSheet(sheetName As String) As Worksheet
    Dim ws As Worksheet
    Set ws = AddSheetAfter(ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = sheetName
    ws.Cells.Font.Name = "メイリオ"
    ws.Cells.Font.Size = 10
    Set AddSheet = ws
End Function

'--------------------------------------------------------------
' 表紙
'--------------------------------------------------------------
Private Sub BuildCoverSheet(ws As Worksheet)

    Dim r As Long
    Dim labels As Variant
    Dim i As Long
    Dim emptyList As New Collection

    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 7
    ws.Columns("C").ColumnWidth = 22
    ws.Columns("D").ColumnWidth = 14
    ws.Columns("E").ColumnWidth = 12
    ws.Columns("F").ColumnWidth = 12
    ws.Columns("G").ColumnWidth = 12
    ws.Columns("H").ColumnWidth = 12
    ws.Columns("I").ColumnWidth = 3

    ws.Range("B2").Value = "QA試験　検査表"
    ws.Range("B2").Font.Size = 16
    ws.Range("B2").Font.Bold = True

    ' 入力欄（品名・ロット番号・状態・日付）。状態は自由入力（入力規則なし）
    labels = Array("品名", "ロット番号", "状態", "日付")
    For i = 0 To 3
        r = 4 + i
        ws.Cells(r, 2).Value = labels(i)
        ws.Cells(r, 2).Interior.Color = COLOR_HEADER
        ws.Cells(r, 2).HorizontalAlignment = xlCenter
        ws.Cells(r, 2).ShrinkToFit = True
        ws.Range(ws.Cells(r, 3), ws.Cells(r, 4)).Merge
        ws.Range(ws.Cells(r, 3), ws.Cells(r, 4)).NumberFormat = "@"
        ws.Range(ws.Cells(r, 2), ws.Cells(r, 4)).Borders.LineStyle = xlContinuous
    Next i
    ws.Range(COVER_DATE_CELL).MergeArea.NumberFormat = "yyyy/mm/dd"
    ws.Range(COVER_DATE_CELL).MergeArea.HorizontalAlignment = xlLeft

    ' 総合判定（F4:F7）と電子印鑑欄（承認者 G4:G7・担当者 H4:H7）。印鑑が収まる大きさにする
    ws.Range("F3:H3").Value = Array("総合判定", "承認者", "担当者")
    ws.Range("F3:H3").Interior.Color = COLOR_HEADER
    ws.Range("F3:H3").HorizontalAlignment = xlCenter
    ws.Range(COVER_OVERALL_CELL).Resize(4, 1).Merge
    ws.Range(COVER_APPROVER_CELL).Resize(4, 1).Merge
    ws.Range(COVER_CHARGE_CELL).Resize(4, 1).Merge
    With ws.Range(COVER_OVERALL_CELL).MergeArea
        .Font.Size = 16
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    EnsureStampArea ws.Range(COVER_APPROVER_CELL).MergeArea
    EnsureStampArea ws.Range(COVER_CHARGE_CELL).MergeArea
    ws.Range("F3:H7").Borders.LineStyle = xlContinuous

    ' 検査一覧の見出し
    r = COVER_LIST_FIRST_ROW - 1
    ws.Cells(r, COVER_COL_NO).Value = "No."
    ws.Cells(r, COVER_COL_NAME).Value = "検査名"
    ws.Cells(r, COVER_COL_RESULT).Value = "判定結果"
    ws.Cells(r, COVER_COL_STAMP).Value = "担当者印"
    With ws.Range(ws.Cells(r, COVER_COL_NO), ws.Cells(r, COVER_COL_STAMP))
        .Interior.Color = COLOR_HEADER
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
    EnsureStampArea ws.Cells(COVER_LIST_FIRST_ROW, COVER_COL_STAMP)   ' 列幅だけ合わせる（行は一覧で作り直す）

    RebuildCoverList emptyList

    AddCoverButtons ws

End Sub

'--------------------------------------------------------------
' 表紙のボタン（承認者押印・担当者押印・保存・DB登録・表紙削除・一括削除）
'--------------------------------------------------------------
Private Sub AddCoverButtons(ws As Worksheet)

    Dim btnLeft As Double
    Dim captions As Variant
    Dim macros As Variant
    Dim i As Long

    btnLeft = ws.Columns("J").Left
    captions = Array("承認者 押印", "担当者 押印", "保存", "DB登録", "表紙削除", "一括削除")
    macros = Array("承認者押印", "担当者押印", "検査表保存", "検査表DB登録", "表紙削除", "一括削除")
    For i = 0 To UBound(captions)
        AddButton ws, COVER_BUTTON_PREFIX & i, CStr(captions(i)), CStr(macros(i)), _
            btnLeft, ws.Range("B3").Top + i * 32, 100, 26
    Next i

End Sub

'--------------------------------------------------------------
' セットアップ済みのブックの表紙ボタンと設定シートを、今の形に作り直す
'   表紙・明細の入力、印鑑、マスタ、電子印鑑、保存履歴はそのまま
'--------------------------------------------------------------
Sub 表紙ボタンと設定の更新()

    Dim cs As Worksheet
    Dim ss As Worksheet
    Dim errMsg As String

    Set cs = FindSheet(COVER_SHEET)
    If cs Is Nothing Then
        ShowMsg "表紙がありません。先に「セットアップ」を実行してください。", vbExclamation, "表紙ボタンと設定の更新"
        Exit Sub
    End If
    If ShowMsg("表紙のボタンと設定シートを作り直します。" & vbCrLf & _
               "（設定シートに書いた内容は消えます。入力・印鑑・マスタはそのままです）よろしいですか？", _
               vbYesNo + vbQuestion + vbDefaultButton2, "表紙ボタンと設定の更新") = vbNo Then Exit Sub

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    DeleteShapesByPrefix cs, COVER_BUTTON_PREFIX
    AddCoverButtons cs

    Set ss = FindSheet(SETTING_SHEET)
    If ss Is Nothing Then
        Set ss = AddSheet(SETTING_SHEET)
    Else
        ss.Cells.Clear
    End If
    BuildSettingSheet ss

    Application.ScreenUpdating = True
    ShowMsg "表紙のボタンと設定シートを作り直しました。" & vbCrLf & _
            "設定シートに BaseDBファイルの場所と NG メールの宛先を入力してください。", vbInformation, "表紙ボタンと設定の更新"
    Exit Sub

ErrHandler:
    errMsg = Err.Description
    Application.ScreenUpdating = True
    ShowMsg "作り直しに失敗しました。" & vbCrLf & errMsg, vbExclamation, "表紙ボタンと設定の更新"

End Sub

'--------------------------------------------------------------
' マスタ（見出しと説明。中身は ImportItemList で入れる）
'--------------------------------------------------------------
Private Sub BuildMasterSheet(ws As Worksheet)

    ws.Range("B1").Value = "マスタ（品名・検査名・数・基準値・検査のルール）"
    ws.Range("B1").Font.Size = 14
    ws.Range("B1").Font.Bold = True
    ws.Range("B2").Value = "1行に品名1つ・検査1つ。数が0の検査は行わない。検査のルールは 以上／以下／範囲／一致。" & _
                           "範囲は基準値1=下限・基準値2=上限、一致は基準値1に合格の文字。基準の表示が空欄ならルールと基準値から作る。"

    ws.Range(ws.Cells(MASTER_HEADER_ROW, MCOL_PRODUCT), ws.Cells(MASTER_HEADER_ROW, MCOL_NOTE)).Value = _
        Array("品名", "検査名", "数", "検査のルール", "基準値1", "基準値2", "単位", "基準の表示", "備考")
    With ws.Range(ws.Cells(MASTER_HEADER_ROW, MCOL_PRODUCT), ws.Cells(MASTER_HEADER_ROW, MCOL_NOTE))
        .Interior.Color = COLOR_HEADER
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    ws.Range(ws.Cells(MASTER_FIRST_ROW, MCOL_PRODUCT), ws.Cells(MASTER_FIRST_ROW + 500, MCOL_INSPECTION)).NumberFormat = "@"
    ws.Range(ws.Cells(MASTER_FIRST_ROW, MCOL_STD1), ws.Cells(MASTER_FIRST_ROW + 500, MCOL_DISPLAY)).NumberFormat = "@"

    ' 検査のルールは一覧から選べるようにする
    With ws.Range(ws.Cells(MASTER_FIRST_ROW, MCOL_METHOD), ws.Cells(MASTER_FIRST_ROW + 500, MCOL_METHOD)).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Formula1:=METHOD_MIN & "," & METHOD_MAX & "," & METHOD_RANGE & "," & METHOD_MATCH
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    ws.Columns("A").ColumnWidth = 2
    ws.Columns(MCOL_PRODUCT).ColumnWidth = 14
    ws.Columns(MCOL_INSPECTION).ColumnWidth = 16
    ws.Columns(MCOL_COUNT).ColumnWidth = 6
    ws.Columns(MCOL_METHOD).ColumnWidth = 12
    ws.Columns(MCOL_STD1).ColumnWidth = 10
    ws.Columns(MCOL_STD2).ColumnWidth = 10
    ws.Columns(MCOL_UNIT).ColumnWidth = 7
    ws.Columns(MCOL_DISPLAY).ColumnWidth = 26
    ws.Columns(MCOL_NOTE).ColumnWidth = 20

End Sub

'--------------------------------------------------------------
' 検査項目一覧（品名 × 検査ごとの数）をマスタに移し、検査項目一覧シートを削除する
'   基準値・検査のルールは既存の QA試験.xlsm の検査表に書かれているものを入れる。
'   結果のメッセージを返す
'--------------------------------------------------------------
Private Function ImportItemList() As String

    Dim src As Worksheet
    Dim ms As Worksheet
    Dim headerRow As Long
    Dim r As Long
    Dim c As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim outRow As Long
    Dim productName As String
    Dim inspName As String
    Dim std As Variant
    Dim unknownNames As String
    Dim alertsBefore As Boolean

    Set src = FindSheet(ITEM_LIST_SHEET)
    If src Is Nothing Then
        ImportItemList = "※「" & ITEM_LIST_SHEET & "」シートが無いため、マスタは空です。マスタに品名と検査を入力してください。"
        Exit Function
    End If

    ' 見出し（B列に「品名」がある行）を探す
    For r = 1 To 20
        If CellText(src.Cells(r, 2).Value) = "品名" Then
            headerRow = r
            Exit For
        End If
    Next r
    If headerRow = 0 Then
        ImportItemList = "※「" & ITEM_LIST_SHEET & "」シートの見出し（B列の「品名」）が見つからないため、マスタは空です。"
        Exit Function
    End If

    Set ms = ThisWorkbook.Worksheets(MASTER_SHEET)
    lastRow = src.Cells(src.Rows.Count, 2).End(xlUp).Row
    lastCol = src.Cells(headerRow, src.Columns.Count).End(xlToLeft).Column
    outRow = MASTER_FIRST_ROW

    For r = headerRow + 1 To lastRow
        productName = CellText(src.Cells(r, 2).Value)
        If productName <> "" Then
            For c = 3 To lastCol
                inspName = CellText(src.Cells(headerRow, c).Value)
                If inspName <> "" And IsIntegerIn(src.Cells(r, c).Value, 1, MAX_SAMPLES) Then
                    ms.Cells(outRow, MCOL_PRODUCT).Value = productName
                    ms.Cells(outRow, MCOL_INSPECTION).Value = inspName
                    ms.Cells(outRow, MCOL_COUNT).Value = CLng(src.Cells(r, c).Value)
                    std = KnownStandard(inspName)
                    If IsArray(std) Then
                        ms.Cells(outRow, MCOL_METHOD).Value = std(0)
                        ms.Cells(outRow, MCOL_STD1).Value = std(1)
                        ms.Cells(outRow, MCOL_STD2).Value = std(2)
                        ms.Cells(outRow, MCOL_UNIT).Value = std(3)
                        ms.Cells(outRow, MCOL_DISPLAY).Value = std(4)
                    ElseIf InStr(unknownNames, "「" & inspName & "」") = 0 Then
                        unknownNames = unknownNames & "「" & inspName & "」"
                    End If
                    outRow = outRow + 1
                End If
            Next c
        End If
    Next r

    alertsBefore = Application.DisplayAlerts
    Application.DisplayAlerts = False
    src.Delete
    Application.DisplayAlerts = alertsBefore

    ImportItemList = "検査項目一覧から " & (outRow - MASTER_FIRST_ROW) & " 行をマスタに移し、検査項目一覧シートを削除しました。"
    If unknownNames <> "" Then
        ImportItemList = ImportItemList & vbCrLf & "※ 次の検査は基準値が分からないため、マスタの検査のルール・基準値を入力してください：" & unknownNames
    End If

End Function

'--------------------------------------------------------------
' 既存の QA試験.xlsm の検査表に書かれている基準値
'   Array(検査のルール, 基準値1, 基準値2, 単位, 基準の表示)。知らない検査名なら Empty
'--------------------------------------------------------------
Private Function KnownStandard(inspName As String) As Variant
    Select Case inspName
        Case "耐熱テスト":     KnownStandard = Array(METHOD_MIN, "300", "", "℃", "300 ℃以上で 合格")
        Case "溶融性テスト":   KnownStandard = Array(METHOD_RANGE, "200", "300", "℃", "200℃以上300℃以下で合格")
        Case "抵抗値テスト":   KnownStandard = Array(METHOD_RANGE, "50", "100", "", "50以上100以下で合格")
        Case "抜き取りテスト": KnownStandard = Array(METHOD_MATCH, "C,F", "", "", "A~GのうちC，Fが合格")
        Case "測長テスト":     KnownStandard = Array(METHOD_MIN, "0.7", "", "", "0.7以上で合格")
        Case Else:             KnownStandard = Empty
    End Select
End Function

'--------------------------------------------------------------
' 電子印鑑（B3:B10 に氏名、C列に部署。印鑑は押すたびに作る）
'--------------------------------------------------------------
Private Sub BuildStampSheet(ws As Worksheet)

    ws.Columns("B").ColumnWidth = 18
    ws.Columns("C").ColumnWidth = 16
    ws.Range("B2:C2").Value = Array("氏名", "部署")
    ws.Range("B2:C2").Interior.Color = COLOR_HEADER
    ws.Range("B2:C2").HorizontalAlignment = xlCenter
    ws.Range("B1").Value = "電子印鑑担当者欄（丸印は苗字だけ。苗字と名前の間は空白で区切る）"
    ws.Range("B3").Value = "田中"   ' 架空の担当者
    ws.Range("C3").Value = "品質保証課"
    ws.Range("B4").Value = "佐藤"
    ws.Range("C4").Value = "品質保証課"
    ws.Range("B2:C10").Borders.LineStyle = xlContinuous

End Sub

'--------------------------------------------------------------
' 設定（Base DB ファイルの場所、NG メールの下書きの Gmail アカウント・宛先・CC・件名）
'--------------------------------------------------------------
Private Sub BuildSettingSheet(ws As Worksheet)

    ws.Cells.Font.Name = "メイリオ"
    ws.Cells.Font.Size = 10
    ws.Columns("B").ColumnWidth = 20
    ws.Columns("C").ColumnWidth = 60
    ws.Range("B1").Value = "設定"
    ws.Range("B1").Font.Size = 14
    ws.Range("B1").Font.Bold = True

    ' Base DB ファイル（C3）
    ws.Range("B3").Value = "BaseDBファイル"
    ws.Range("C4").Value = "空欄の場合は、このブックと同じフォルダの QA試験DB.odb を使います（自動では作りません）。"

    ' NG メール（C6:C9）
    ws.Range("B6").Value = "NGメール Gmailアカウント"
    ws.Range("B7").Value = "NGメール 宛先"
    ws.Range("B8").Value = "NGメール CC"
    ws.Range("B9").Value = "NGメール 件名"
    ws.Range("C10").Value = "NG のとき、保存の後に Gmail の作成画面を開きます（送信はしません）。宛先が空欄の場合は開きません。" & _
                            "複数の宛先は「,」で区切ります。件名が空欄の場合は「【NG】QA試験 品名 ロット番号」にします。"
    ws.Range("C11").Value = "Gmail アカウントが空欄の場合は、ブラウザでログイン中のアカウントで開きます。"

    With Union(ws.Range("B3"), ws.Range("B6:B9"))
        .Interior.Color = COLOR_HEADER
    End With
    With Union(ws.Range(SETTING_DB_CELL), ws.Range(SETTING_MAIL_FROM_CELL), ws.Range(SETTING_MAIL_TO_CELL), _
               ws.Range(SETTING_MAIL_CC_CELL), ws.Range(SETTING_MAIL_SUBJECT_CELL))
        .Borders.LineStyle = xlContinuous
        .NumberFormat = "@"
    End With

End Sub
