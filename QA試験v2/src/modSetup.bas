Attribute VB_Name = "modSetup"
Option Explicit

'==============================================================
' QA試験 セットアップ
'   新しいブックに各モジュールを取り込んだ後、1回だけ実行する。
'   表紙・マスタ・電子印鑑・設定シートとボタンを作り、
'   架空のサンプルマスタ（品名2つ）を入れる。
'   明細シートは表紙の品名を入力したときに作られる。
'==============================================================

Sub セットアップ()

    Dim ws As Worksheet
    Dim i As Long
    Dim alertsBefore As Boolean
    Dim errMsg As String

    If Not FindSheet(COVER_SHEET) Is Nothing Then
        ShowMsg "既にセットアップ済みです（「" & COVER_SHEET & "」シートがあります）。", vbInformation, "セットアップ"
        Exit Sub
    End If

    On Error GoTo ErrHandler
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    BuildCoverSheet AddSheet(COVER_SHEET)
    BuildMasterSheet AddSheet(MASTER_SHEET)
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

    ShowMsg "セットアップが完了しました。" & vbCrLf & _
            "電子印鑑シートで「電子印鑑生成」を押してから、表紙の品名を入力してください。", _
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
    Dim btnLeft As Double
    Dim captions As Variant
    Dim macros As Variant
    Dim i As Long
    Dim emptyList As New Collection

    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 7
    ws.Columns("C").ColumnWidth = 22
    ws.Columns("D").ColumnWidth = 14
    ws.Columns("E").ColumnWidth = 12
    ws.Columns("F").ColumnWidth = 12
    ws.Columns("G").ColumnWidth = 12
    ws.Columns("H").ColumnWidth = 3

    ws.Range("B2").Value = "QA試験　検査表"
    ws.Range("B2").Font.Size = 16
    ws.Range("B2").Font.Bold = True

    ' 入力欄（品名・ロット・状態・日付）。状態は自由入力（入力規則なし）
    labels = Array("品名", "ロット", "状態", "日付")
    For i = 0 To 3
        r = 4 + i
        ws.Cells(r, 2).Value = labels(i)
        ws.Cells(r, 2).Interior.Color = COLOR_HEADER
        ws.Cells(r, 2).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(r, 3), ws.Cells(r, 4)).Merge
        ws.Range(ws.Cells(r, 3), ws.Cells(r, 4)).NumberFormat = "@"
        ws.Range(ws.Cells(r, 2), ws.Cells(r, 4)).Borders.LineStyle = xlContinuous
    Next i
    ws.Range(COVER_DATE_CELL).MergeArea.NumberFormat = "yyyy/mm/dd"
    ws.Range(COVER_DATE_CELL).MergeArea.HorizontalAlignment = xlLeft

    ' 電子印鑑欄（承認者・担当者）
    ws.Range("F3").Value = "承認者"
    ws.Range("G3").Value = "担当者"
    ws.Range("F3:G3").Interior.Color = COLOR_HEADER
    ws.Range("F3:G3").HorizontalAlignment = xlCenter
    ws.Range(COVER_APPROVER_CELL).Resize(4, 1).Merge
    ws.Range(COVER_CHARGE_CELL).Resize(4, 1).Merge
    ws.Range("F3:G7").Borders.LineStyle = xlContinuous

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

    RebuildCoverList emptyList

    ' ボタン
    btnLeft = ws.Columns("I").Left
    captions = Array("承認者 押印", "担当者 押印", "Excel保存", "PDF保存", "一括保存")
    macros = Array("承認者押印", "担当者押印", "検査表Excel保存", "検査表PDF保存", "検査表一括保存")
    For i = 0 To UBound(captions)
        AddButton ws, "btnCover_" & i, CStr(captions(i)), CStr(macros(i)), _
            btnLeft, ws.Range("B3").Top + i * 32, 100, 26
    Next i

End Sub

'--------------------------------------------------------------
' マスタ（架空のサンプルデータ入り）
'--------------------------------------------------------------
Private Sub BuildMasterSheet(ws As Worksheet)

    Dim data As Variant
    Dim i As Long
    Dim r As Long

    ws.Range("B1").Value = "マスタ（品名・検査・項目と基準）"
    ws.Range("B1").Font.Size = 14
    ws.Range("B1").Font.Bold = True
    ws.Range("B2").Value = "1行に1項目。判定方式は 一致／以上／以下／範囲。一致は基準値1に合格の文字（カンマ区切りで複数可）、" & _
                           "範囲は基準値1=下限・基準値2=上限。不合格個数は検査ごと（その検査のどれか1行に入力）。"

    ws.Range(ws.Cells(MASTER_HEADER_ROW, MCOL_PRODUCT), ws.Cells(MASTER_HEADER_ROW, MCOL_NOTE)).Value = _
        Array("品名", "検査名", "項目名", "判定方式", "基準値1", "基準値2", "n数", "不合格個数", "備考")
    With ws.Range(ws.Cells(MASTER_HEADER_ROW, MCOL_PRODUCT), ws.Cells(MASTER_HEADER_ROW, MCOL_NOTE))
        .Interior.Color = COLOR_HEADER
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    ' 架空のサンプル（品名, 検査名, 項目名, 判定方式, 基準値1, 基準値2, n数, 不合格個数）
    data = Array( _
        Array("SMP-A100", "外観検査", "外観ランク", "一致", "A,B", "", 5, 2), _
        Array("SMP-A100", "外観検査", "色調", "一致", "良", "", 3, ""), _
        Array("SMP-A100", "寸法検査", "長さ(mm)", "範囲", "9.5", "10.5", 5, 1), _
        Array("SMP-A100", "寸法検査", "厚み(mm)", "以下", "2.0", "", 5, ""), _
        Array("SMP-A100", "強度検査", "引張強度(kN)", "以上", "0.9", "", 3, 1), _
        Array("SMP-B200", "耐熱検査", "耐熱温度(℃)", "以上", "300", "", 5, 1), _
        Array("SMP-B200", "抵抗検査", "抵抗値(Ω)", "範囲", "50", "100", 5, 2), _
        Array("SMP-B200", "抵抗検査", "抜き取り判定", "一致", "C,F", "", 5, ""))

    ws.Range(ws.Cells(MASTER_FIRST_ROW, MCOL_PRODUCT), ws.Cells(MASTER_FIRST_ROW + 200, MCOL_ITEM)).NumberFormat = "@"
    For i = 0 To UBound(data)
        r = MASTER_FIRST_ROW + i
        ws.Cells(r, MCOL_PRODUCT).Value = data(i)(0)
        ws.Cells(r, MCOL_INSPECTION).Value = data(i)(1)
        ws.Cells(r, MCOL_ITEM).Value = data(i)(2)
        ws.Cells(r, MCOL_METHOD).Value = data(i)(3)
        If CStr(data(i)(4)) <> "" Then ws.Cells(r, MCOL_STD1).Value = data(i)(4)
        If CStr(data(i)(5)) <> "" Then ws.Cells(r, MCOL_STD2).Value = data(i)(5)
        ws.Cells(r, MCOL_N).Value = data(i)(6)
        If CStr(data(i)(7)) <> "" Then ws.Cells(r, MCOL_FAILCOUNT).Value = data(i)(7)
    Next i

    ws.Columns("A").ColumnWidth = 2
    ws.Columns(MCOL_PRODUCT).ColumnWidth = 14
    ws.Columns(MCOL_INSPECTION).ColumnWidth = 14
    ws.Columns(MCOL_ITEM).ColumnWidth = 18
    ws.Columns(MCOL_METHOD).ColumnWidth = 10
    ws.Columns(MCOL_STD1).ColumnWidth = 10
    ws.Columns(MCOL_STD2).ColumnWidth = 10
    ws.Columns(MCOL_N).ColumnWidth = 7
    ws.Columns(MCOL_FAILCOUNT).ColumnWidth = 11
    ws.Columns(MCOL_NOTE).ColumnWidth = 24

End Sub

'--------------------------------------------------------------
' 電子印鑑（既存ツールと同じ配置：B3:B10 に氏名、C列に印影）
'--------------------------------------------------------------
Private Sub BuildStampSheet(ws As Worksheet)

    ws.Columns("B").ColumnWidth = 18
    ws.Columns("C").ColumnWidth = 12
    ws.Range("B2").Value = "電子印鑑担当者欄"
    ws.Range("B2").Interior.Color = COLOR_HEADER
    ws.Range("B3").Value = "田中"   ' 架空の担当者
    ws.Range("B4").Value = "佐藤"
    ws.Range("B2:B10").Borders.LineStyle = xlContinuous

    AddButton ws, "btnStampCreate", "電子印鑑生成", "電子印鑑生成", ws.Columns("E").Left, ws.Range("B2").Top, 100, 26
    AddButton ws, "btnStampDelete", "電子印鑑削除", "電子印鑑削除", ws.Columns("E").Left, ws.Range("B2").Top + 32, 100, 26

End Sub

'--------------------------------------------------------------
' 設定（Access DB ファイルの場所）
'--------------------------------------------------------------
Private Sub BuildSettingSheet(ws As Worksheet)

    ws.Columns("B").ColumnWidth = 18
    ws.Columns("C").ColumnWidth = 60
    ws.Range("B1").Value = "設定"
    ws.Range("B1").Font.Size = 14
    ws.Range("B1").Font.Bold = True
    ws.Range("B3").Value = "AccessDBファイル"
    ws.Range("B3").Interior.Color = COLOR_HEADER
    ws.Range("C3").Borders.LineStyle = xlContinuous
    ws.Range("C4").Value = "空欄の場合は、このブックと同じフォルダの QA試験DB.accdb を使います（無ければ自動で作成）。"
    ws.Range("C5").Value = "共有フォルダに置く場合は「\\サーバー名\共有\QA試験DB.accdb」のように書きます。"

End Sub
