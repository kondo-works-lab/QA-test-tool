Attribute VB_Name = "modTest"
Option Explicit

'==============================================================
' QA試験 自動テスト
'   「テスト実行」を実行すると、テスト用の架空品名（ZZTEST-…）をマスタの末尾に
'   一時的に追加して、判定・オレンジ表示・表紙への反映・総合判定・
'   品名変更時の作り直し・押印・保存（Excel/PDF）・Access登録を確認し、
'   結果を「テスト結果」シートに書き出す。
'   ・メッセージは表示せずに記録する（確認ダイアログは「はい」で進める）。
'   ・テスト用の DB・保存ファイルは一時フォルダ（%TEMP%）に作る。
'   ・終了後、テスト用のマスタ行を削除し、表紙の品名・ロット・状態・日付を元に戻す
'     （明細シートは作り直されるため、入力中のデータは消える）。
'==============================================================

Private Const RESULT_SHEET As String = "テスト結果"

Private mRow As Long
Private mOkCount As Long
Private mNgCount As Long
Private mSkipCount As Long
Private mResultWs As Worksheet

'--------------------------------------------------------------
' 実行用（ボタンやマクロ一覧から）
'--------------------------------------------------------------
Sub テスト実行()
    If MsgBox("自動テストを実行します。" & vbCrLf & _
              "表紙と明細シートの入力中のデータは消えます。よろしいですか？", _
              vbYesNo + vbExclamation + vbDefaultButton2, "テスト実行") = vbNo Then Exit Sub
    MsgBox RunAllTests(), vbInformation, "テスト実行"
End Sub

'--------------------------------------------------------------
' すべてのテストを実行し、結果の要約を返す
'--------------------------------------------------------------
Public Function RunAllTests() As String

    Dim cs As Worksheet
    Dim ms As Worksheet
    Dim saved(0 To 3) As Variant
    Dim firstTestRow As Long
    Dim errMsg As String

    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)
    Set ms = ThisWorkbook.Worksheets(MASTER_SHEET)

    PrepareResultSheet
    gTestMode = True
    gTestLog = ""
    gTestAnswer = vbYes

    saved(0) = cs.Range(COVER_PRODUCT_CELL).Value
    saved(1) = cs.Range(COVER_LOT_CELL).Value
    saved(2) = cs.Range(COVER_STATE_CELL).Value
    saved(3) = cs.Range(COVER_DATE_CELL).Value

    firstTestRow = MasterLastRow(ms) + 1
    AddTestMaster ms, firstTestRow

    On Error GoTo ErrHandler
    TestStandardFunctions
    TestBuildSheets
    TestJudgement
    TestStamp
    TestProductChange
    TestAccess
    TestSave
    On Error GoTo 0

Cleanup:
    On Error Resume Next
    ms.Range(ms.Rows(firstTestRow), ms.Rows(firstTestRow + 30)).Delete
    SetValue cs, cs.Range(COVER_LOT_CELL), saved(1)
    SetValue cs, cs.Range(COVER_STATE_CELL), saved(2)
    SetValue cs, cs.Range(COVER_DATE_CELL), saved(3)
    SetValue cs, cs.Range(COVER_PRODUCT_CELL), saved(0)
    gTestMode = False
    gTestDbPath = ""
    gTestSaveFolder = ""
    gTestAnswer = 0
    On Error GoTo 0

    mResultWs.Cells(1, 1).Value = "テスト結果：OK " & mOkCount & " 件 / NG " & mNgCount & " 件 / SKIP " & mSkipCount & " 件"
    RunAllTests = "OK " & mOkCount & " / NG " & mNgCount & " / SKIP " & mSkipCount
    Exit Function

ErrHandler:
    errMsg = Err.Description
    Check "予期しないエラー", False, errMsg
    Resume Cleanup

End Function

'--------------------------------------------------------------
' テスト用のマスタ行（架空の品名）
'--------------------------------------------------------------
Private Sub AddTestMaster(ms As Worksheet, firstRow As Long)

    Dim data As Variant
    Dim i As Long

    ' 品名, 検査名, 項目名, 判定方式, 基準値1, 基準値2, n数, 不合格個数
    data = Array( _
        Array("ZZTEST-1", "検査A", "ランク", "一致", "A,B", "", 3, 2), _
        Array("ZZTEST-1", "検査A", "強度", "以上", "0.9", "", 2, ""), _
        Array("ZZTEST-1", "検査B", "長さ", "範囲", "10", "20", 2, 1), _
        Array("ZZTEST-1", "検査B", "厚み", "以下", "5", "", 1, 1), _
        Array("ZZTEST-2", "検査C", "温度", "以上", "300", "", 1, 1), _
        Array("ZZTEST-2", "検査D", "抵抗", "範囲", "50", "100", 1, 1), _
        Array("ZZTEST-2", "検査E", "外観", "一致", "良", "", 1, 1), _
        Array("ZZTEST-BAD", "検査X", "項目", "大きい", "1", "", 1, 1))

    For i = 0 To UBound(data)
        ms.Cells(firstRow + i, MCOL_PRODUCT).Value = data(i)(0)
        ms.Cells(firstRow + i, MCOL_INSPECTION).Value = data(i)(1)
        ms.Cells(firstRow + i, MCOL_ITEM).Value = data(i)(2)
        ms.Cells(firstRow + i, MCOL_METHOD).Value = data(i)(3)
        If CStr(data(i)(4)) <> "" Then ms.Cells(firstRow + i, MCOL_STD1).Value = data(i)(4)
        If CStr(data(i)(5)) <> "" Then ms.Cells(firstRow + i, MCOL_STD2).Value = data(i)(5)
        ms.Cells(firstRow + i, MCOL_N).Value = data(i)(6)
        If CStr(data(i)(7)) <> "" Then ms.Cells(firstRow + i, MCOL_FAILCOUNT).Value = data(i)(7)
    Next i

End Sub

'--------------------------------------------------------------
' 判定方式ごとの合否
'--------------------------------------------------------------
Private Sub TestStandardFunctions()
    Check "一致：A,B に B は合格", IsPassValue("B", METHOD_MATCH, "A,B", "")
    Check "一致：小文字 a は合格", IsPassValue("a", METHOD_MATCH, "A,B", "")
    Check "一致：全角 Ｂ は合格", IsPassValue("Ｂ", METHOD_MATCH, "A，B", "")
    Check "一致：C は不合格", Not IsPassValue("C", METHOD_MATCH, "A,B", "")
    Check "一致：「良」に 良 は合格", IsPassValue("良", METHOD_MATCH, "良", "")
    Check "以上：0.9 以上に 0.9 は合格（境界）", IsPassValue(0.9, METHOD_MIN, "0.9", "")
    Check "以上：0.9 以上に 0.89 は不合格", Not IsPassValue(0.89, METHOD_MIN, "0.9", "")
    Check "以下：5 以下に 5 は合格（境界）", IsPassValue(5, METHOD_MAX, "5", "")
    Check "以下：5 以下に 5.01 は不合格", Not IsPassValue(5.01, METHOD_MAX, "5", "")
    Check "範囲：10～20 に 10 と 20 は合格（境界）", IsPassValue(10, METHOD_RANGE, "10", "20") And IsPassValue(20, METHOD_RANGE, "10", "20")
    Check "範囲：10～20 に 9.99 と 20.01 は不合格", Not IsPassValue(9.99, METHOD_RANGE, "10", "20") And Not IsPassValue(20.01, METHOD_RANGE, "10", "20")
    Check "基準の表示（一致）", FormatStandard(METHOD_MATCH, "A,B", "") = "A / B", FormatStandard(METHOD_MATCH, "A,B", "")
    Check "基準の表示（範囲）", FormatStandard(METHOD_RANGE, "10", "20") = "10 ～ 20", FormatStandard(METHOD_RANGE, "10", "20")
End Sub

'--------------------------------------------------------------
' 品名入力で明細シートと表紙の一覧が作られる
'--------------------------------------------------------------
Private Sub TestBuildSheets()

    Dim cs As Worksheet
    Dim wa As Worksheet
    Dim wb As Worksheet

    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)
    SetValue cs, cs.Range(COVER_PRODUCT_CELL), "ZZTEST-1"

    Set wa = FindSheet(DETAIL_PREFIX & "検査A")
    Set wb = FindSheet(DETAIL_PREFIX & "検査B")
    Check "品名 ZZTEST-1：明細シートが2枚", GetDetailSheets().Count = 2 And Not wa Is Nothing And Not wb Is Nothing, _
        "明細シート数=" & GetDetailSheets().Count
    If wa Is Nothing Or wb Is Nothing Then Exit Sub

    Check "表紙の検査一覧が2行", CoverListCount() = 2 And FindCoverRow("検査A") = COVER_LIST_FIRST_ROW And FindCoverRow("検査B") = COVER_LIST_FIRST_ROW + 1
    Check "総合判定欄の見出し", cs.Cells(COVER_LIST_FIRST_ROW + 3, COVER_COL_NAME).Value = "総合判定"
    Check "明細の項目と基準（検査A）", wa.Cells(7, DETAIL_COL_ITEM).Value = "ランク" And wa.Cells(7, DETAIL_COL_STD).Value = "A / B" _
        And wa.Cells(8, DETAIL_COL_ITEM).Value = "強度" And wa.Cells(8, DETAIL_COL_STD).Value = "0.9 以上", _
        wa.Cells(7, DETAIL_COL_STD).Value & " | " & wa.Cells(8, DETAIL_COL_STD).Value
    Check "明細の項目と基準（検査B）", wb.Cells(7, DETAIL_COL_STD).Value = "10 ～ 20" And wb.Cells(8, DETAIL_COL_STD).Value = "5 以下", _
        wb.Cells(7, DETAIL_COL_STD).Value & " | " & wb.Cells(8, DETAIL_COL_STD).Value
    Check "n数を超える欄はグレー（検査A 強度の3列目）", wa.Cells(8, DETAIL_COL_INPUT + 2).Interior.Color = COLOR_INACTIVE _
        And wa.Cells(8, DETAIL_COL_INPUT + 1).Interior.Color = COLOR_INPUT
    Check "不合格個数が明細に反映（検査A=2）", wa.Range(DETAIL_FAILCOUNT_CELL).Value = 2

End Sub

'--------------------------------------------------------------
' 判定・オレンジ表示・表紙への反映・総合判定
'--------------------------------------------------------------
Private Sub TestJudgement()

    Dim cs As Worksheet
    Dim wa As Worksheet
    Dim wb As Worksheet
    Dim ok As Boolean
    Dim i0 As Long

    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)
    Set wa = FindSheet(DETAIL_PREFIX & "検査A")
    Set wb = FindSheet(DETAIL_PREFIX & "検査B")
    If wa Is Nothing Or wb Is Nothing Then
        Skip "判定のテスト", "明細シートがありません"
        Exit Sub
    End If
    i0 = DETAIL_COL_INPUT

    ' ① 未入力チェック
    gTestLog = ""
    SetValue wa, wa.Cells(7, i0), "A"
    ok = JudgeDetailSheet(wa, Now)
    Check "未入力があると判定しない", Not ok And InStr(gTestLog, "未入力の欄があります") > 0 And GetDetailResult(wa) = "", gTestLog

    ' グレーの欄への入力・数値以外
    SetValue wa, wa.Cells(7, i0 + 1), "B"
    SetValue wa, wa.Cells(7, i0 + 2), "C"
    SetValue wa, wa.Cells(8, i0), "abc"
    SetValue wa, wa.Cells(8, i0 + 1), 0.89
    SetValue wa, wa.Cells(8, i0 + 2), 1
    gTestLog = ""
    ok = JudgeDetailSheet(wa, Now)
    Check "グレーの欄への入力・数値以外を検出", Not ok And InStr(gTestLog, "グレーの欄に入力があります") > 0 _
        And InStr(gTestLog, "数値以外が入力されています") > 0, gTestLog
    SetValue wa, wa.Cells(8, i0 + 2), ""

    ' ② オレンジ ③ 合否：検査A は C と 0.89 が基準外 → 2個 ≧ 2 で不合格
    SetValue wa, wa.Cells(8, i0), 0.9
    gTestLog = ""
    ok = JudgeDetailSheet(wa, Now)
    Check "検査A：判定できる", ok, gTestLog
    Check "検査A：基準外のセルだけオレンジ", wa.Cells(7, i0 + 2).Interior.Color = COLOR_NG And wa.Cells(8, i0 + 1).Interior.Color = COLOR_NG _
        And wa.Cells(7, i0).Interior.Color = COLOR_INPUT And wa.Cells(7, i0 + 1).Interior.Color = COLOR_INPUT _
        And wa.Cells(8, i0).Interior.Color = COLOR_INPUT
    Check "検査A：基準外2個 ≧ 不合格個数2 → 不合格", wa.Cells(DetailSumRow(wa), DETAIL_COL_METHOD).Value = 2 And GetDetailResult(wa) = RESULT_FAIL, _
        GetDetailResult(wa)
    Check "表紙：検査A が不合格", cs.Cells(FindCoverRow("検査A"), COVER_COL_RESULT).Value = RESULT_FAIL
    Check "総合判定：1つでも不合格なら不合格（検査B 未判定）", GetOverall() = RESULT_FAIL, GetOverall()

    ' 判定後に入力を変更すると判定結果が消える
    gTestLog = ""
    SetValue wa, wa.Cells(8, i0 + 1), 1
    Check "入力変更で判定結果・オレンジ・表紙の結果を消去", GetDetailResult(wa) = "" And wa.Cells(7, i0 + 2).Interior.Color = COLOR_INPUT _
        And cs.Cells(FindCoverRow("検査A"), COVER_COL_RESULT).Value = "" And GetOverall() = "" _
        And InStr(gTestLog, "判定結果を消去しました") > 0, gTestLog

    ' 再判定：基準外1個 ＜ 2 → 合格（オレンジは残る）
    ok = JudgeDetailSheet(wa, Now)
    Check "検査A：基準外1個 ＜ 2 → 合格（オレンジ1個）", ok And GetDetailResult(wa) = RESULT_PASS _
        And wa.Cells(7, i0 + 2).Interior.Color = COLOR_NG And wa.Cells(DetailSumRow(wa), DETAIL_COL_METHOD).Value = 1
    Check "総合判定：未判定の検査があれば空欄", GetOverall() = "", GetOverall()

    ' 検査B：境界値ちょうどは合格 → 全検査合格で総合判定「合格」
    SetValue wb, wb.Cells(7, i0), 10
    SetValue wb, wb.Cells(7, i0 + 1), 20
    SetValue wb, wb.Cells(8, i0), 5
    ok = JudgeDetailSheet(wb, Now)
    Check "検査B：境界値ちょうどは合格", ok And GetDetailResult(wb) = RESULT_PASS And wb.Cells(DetailSumRow(wb), DETAIL_COL_METHOD).Value = 0
    Check "総合判定：全検査合格なら合格", GetOverall() = RESULT_PASS, GetOverall()

    ' 検査B：基準外1個 ≧ 1 → 不合格 → 総合判定「不合格」
    SetValue wb, wb.Cells(7, i0 + 1), 20.1
    ok = JudgeDetailSheet(wb, Now)
    Check "検査B：20.1 は範囲外 → 不合格", ok And GetDetailResult(wb) = RESULT_FAIL And wb.Cells(7, i0 + 1).Interior.Color = COLOR_NG
    Check "総合判定：不合格", GetOverall() = RESULT_FAIL And cs.Cells(FindCoverRow("検査B"), COVER_COL_RESULT).Value = RESULT_FAIL

    ' 最後は全検査合格の状態にしておく（保存のテスト用）
    SetValue wb, wb.Cells(7, i0 + 1), 20
    JudgeDetailSheet wb, Now

End Sub

'--------------------------------------------------------------
' 押印（電子印鑑シートに田中の印影がある場合のみ）
'--------------------------------------------------------------
Private Sub TestStamp()

    Dim cs As Worksheet
    Dim wa As Worksheet

    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)
    If Not ShapeExists(ThisWorkbook.Worksheets(STAMP_SHEET), "Stamp_田中") Then
        Skip "押印", "電子印鑑シートに田中の印影がありません（先に「電子印鑑生成」を実行）"
        Exit Sub
    End If

    SetStampTarget cs.Range(COVER_APPROVER_CELL).MergeArea, APPROVER_STAMP_NAME
    ApplyApproverStamp "田中"
    SetStampTarget cs.Range(COVER_CHARGE_CELL).MergeArea, CHARGE_STAMP_NAME
    ApplyApproverStamp "田中"
    SetStampTarget cs.Cells(FindCoverRow("検査A"), COVER_COL_STAMP), INSP_STAMP_PREFIX & "検査A"
    ApplyApproverStamp "田中"
    SetStampTarget cs.Cells(FindCoverRow("検査B"), COVER_COL_STAMP), INSP_STAMP_PREFIX & "検査B"
    ApplyApproverStamp "田中"
    Check "押印：承認者・担当者・検査担当者", ShapeExists(cs, APPROVER_STAMP_NAME) And ShapeExists(cs, CHARGE_STAMP_NAME) _
        And ShapeExists(cs, INSP_STAMP_PREFIX & "検査A") And ShapeExists(cs, INSP_STAMP_PREFIX & "検査B")
    Check "押印した氏名を読み取れる", GetStampName(cs, APPROVER_STAMP_NAME) = "田中", GetStampName(cs, APPROVER_STAMP_NAME)

End Sub

'--------------------------------------------------------------
' 品名変更：明細シートの作り直し・マスタに無い品名・マスタの誤り
'--------------------------------------------------------------
Private Sub TestProductChange()

    Dim cs As Worksheet
    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)

    ' 保存・Access のテスト用に、ZZTEST-1 の判定済み状態を残すため最後に ZZTEST-1 を作り直す
    gTestLog = ""
    SetValue cs, cs.Range(COVER_PRODUCT_CELL), "ZZTEST-2"
    Check "品名変更：検査3つの品名で明細シートが3枚に", GetDetailSheets().Count = 3 And CoverListCount() = 3 _
        And FindSheet(DETAIL_PREFIX & "検査A") Is Nothing And Not FindSheet(DETAIL_PREFIX & "検査E") Is Nothing, _
        "明細シート数=" & GetDetailSheets().Count
    Check "品名変更：判定結果・印鑑を消去しメッセージ", GetOverall() = "" And Not ShapeExists(cs, APPROVER_STAMP_NAME) _
        And InStr(gTestLog, "明細シートを作り直しました") > 0, gTestLog

    gTestLog = ""
    SetValue cs, cs.Range(COVER_PRODUCT_CELL), "ZZTEST-NONE"
    Check "マスタに無い品名：明細シートなし＋メッセージ", GetDetailSheets().Count = 0 And CoverListCount() = 0 _
        And InStr(gTestLog, "マスタシートにありません") > 0, gTestLog

    gTestLog = ""
    SetValue cs, cs.Range(COVER_PRODUCT_CELL), "ZZTEST-BAD"
    Check "マスタの誤り（判定方式）：明細シートなし＋メッセージ", GetDetailSheets().Count = 0 _
        And InStr(gTestLog, "判定方式は") > 0, gTestLog

    gTestLog = ""
    SetValue cs, cs.Range(COVER_PRODUCT_CELL), ""
    Check "品名を空にすると明細シートなし（メッセージなし）", GetDetailSheets().Count = 0 And gTestLog = "", gTestLog

End Sub

'--------------------------------------------------------------
' Access 登録（一時フォルダにテスト用 DB を作る）
'--------------------------------------------------------------
Private Sub TestAccess()

    Dim cs As Worksheet
    Dim ws As Worksheet
    Dim conn As Object
    Dim rs As Object
    Dim tempDir As String
    Dim pressTime As Date
    Dim ok As Boolean

    tempDir = Environ("TEMP")
    If tempDir = "" Then
        Skip "Access登録", "一時フォルダ（%TEMP%）がありません"
        Exit Sub
    End If

    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)
    gTestDbPath = tempDir & "\QA試験_テストDB_" & Format(Now, "yyyymmdd_hhnnss") & ".accdb"

    PrepareJudgedProduct
    SetValue cs, cs.Range(COVER_LOT_CELL), "LOT-T001"
    SetValue cs, cs.Range(COVER_STATE_CELL), "量産"
    SetValue cs, cs.Range(COVER_DATE_CELL), DateSerial(2026, 1, 2)

    Set ws = FindSheet(DETAIL_PREFIX & "検査A")
    pressTime = DateSerial(2026, 1, 2) + TimeSerial(3, 4, 5)
    gTestLog = ""
    ok = JudgeDetailSheet(ws, pressTime)
    Check "Access：DBファイルを自動作成", ok And Dir(gTestDbPath) <> "", gTestLog

    On Error GoTo DbErr
    Set conn = OpenDb(gTestDbPath)
    Set rs = conn.Execute("SELECT COUNT(*) FROM [T_判定] WHERE [検査名]='検査A' AND [品名]='ZZTEST-1'")
    Check "Access：T_判定 に1行登録", rs(0).Value = 1, "件数=" & rs(0).Value
    Set rs = conn.Execute("SELECT TOP 1 [判定ID],[押下日時],[判定結果],[基準外個数],[ロット],[状態],[日付] FROM [T_判定] ORDER BY [判定ID] DESC")
    Check "Access：押下日時・判定結果・ロット等", rs("押下日時").Value = pressTime And rs("判定結果").Value = GetDetailResult(ws) _
        And rs("ロット").Value = "LOT-T001" And rs("状態").Value = "量産" And rs("日付").Value = DateSerial(2026, 1, 2), _
        rs("押下日時").Value & " / " & rs("判定結果").Value
    Set rs = conn.Execute("SELECT COUNT(*) FROM [T_判定明細] WHERE [判定ID]=" & rs("判定ID").Value)
    Check "Access：T_判定明細 に入力値の数（5行）登録", rs(0).Value = 5, "件数=" & rs(0).Value

    ' 再判定は追記される
    JudgeDetailSheet ws, Now
    Set rs = conn.Execute("SELECT COUNT(*) FROM [T_判定] WHERE [検査名]='検査A' AND [品名]='ZZTEST-1'")
    Check "Access：再判定は追記（2行）", rs(0).Value = 2, "件数=" & rs(0).Value
    conn.Close
    mResultWs.Cells(mRow, 2).Value = "（テスト用DB：" & gTestDbPath & "）"
    mRow = mRow + 1
    Exit Sub

DbErr:
    Check "Access：DBの読み取り", False, Err.Description

End Sub

'--------------------------------------------------------------
' 保存（一時フォルダに Excel / PDF を保存）
'--------------------------------------------------------------
Private Sub TestSave()

    Dim cs As Worksheet
    Dim histWs As Worksheet
    Dim histRows As Long
    Dim tempDir As String
    Dim baseName As String
    Dim errText As String

    tempDir = Environ("TEMP")
    If tempDir = "" Then
        Skip "保存（Excel/PDF）", "一時フォルダ（%TEMP%）がありません"
        Exit Sub
    End If

    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)
    If GetDetailSheets().Count = 0 Then PrepareJudgedProduct
    SetValue cs, cs.Range(COVER_LOT_CELL), "LOT-T001"
    SetValue cs, cs.Range(COVER_DATE_CELL), DateSerial(2026, 1, 2)

    Check "ファイル名：品名_ロット_日付", BuildFileBaseName(baseName, errText) And baseName = "ZZTEST-1_LOT-T001_20260102", baseName & errText

    gTestSaveFolder = tempDir & "\QA試験_テスト保存_" & Format(Now, "yyyymmdd_hhnnss")
    MkDir gTestSaveFolder

    Set histWs = FindSheet("保存履歴")
    If Not histWs Is Nothing Then histRows = histWs.Cells(histWs.Rows.Count, 1).End(xlUp).Row

    gTestLog = ""
    検査表一括保存
    Check "保存：Excel（.xlsx）", Dir(gTestSaveFolder & "\" & baseName & ".xlsx") <> "", gTestLog
    Check "保存：PDF", Dir(gTestSaveFolder & "\" & baseName & ".pdf") <> "", gTestLog

    Set histWs = FindSheet("保存履歴")
    If Not histWs Is Nothing Then
        Check "保存履歴に2行記録", histWs.Cells(histWs.Rows.Count, 1).End(xlUp).Row = IIf(histRows = 0, 1, histRows) + 2
        ' テストで追加した履歴は消す（テストで作られたシートならシートごと消す）
        If histRows = 0 Then
            Application.DisplayAlerts = False
            histWs.Delete
            Application.DisplayAlerts = True
        Else
            histWs.Range(histWs.Rows(histRows + 1), histWs.Rows(histRows + 2)).Delete
        End If
    Else
        Check "保存履歴に2行記録", False, "保存履歴シートがありません"
    End If

    mResultWs.Cells(mRow, 2).Value = "（テスト保存先：" & gTestSaveFolder & "）"
    mRow = mRow + 1

End Sub

'--------------------------------------------------------------
' ZZTEST-1 を全検査判定済み（合格）にする
'--------------------------------------------------------------
Private Sub PrepareJudgedProduct()

    Dim cs As Worksheet
    Dim wa As Worksheet
    Dim wb As Worksheet
    Dim i0 As Long

    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)
    SetValue cs, cs.Range(COVER_PRODUCT_CELL), "ZZTEST-1"
    Set wa = FindSheet(DETAIL_PREFIX & "検査A")
    Set wb = FindSheet(DETAIL_PREFIX & "検査B")
    i0 = DETAIL_COL_INPUT

    SetValue wa, wa.Cells(7, i0), "A"
    SetValue wa, wa.Cells(7, i0 + 1), "B"
    SetValue wa, wa.Cells(7, i0 + 2), "C"
    SetValue wa, wa.Cells(8, i0), 0.9
    SetValue wa, wa.Cells(8, i0 + 1), 1
    SetValue wb, wb.Cells(7, i0), 10
    SetValue wb, wb.Cells(7, i0 + 1), 20
    SetValue wb, wb.Cells(8, i0), 5
    JudgeDetailSheet wa, Now
    JudgeDetailSheet wb, Now

End Sub

'--------------------------------------------------------------
' セルに値を入れ、シート変更イベントと同じ処理を呼ぶ
'--------------------------------------------------------------
Private Sub SetValue(ws As Worksheet, cell As Range, v As Variant)
    Dim eventsBefore As Boolean
    eventsBefore = Application.EnableEvents
    Application.EnableEvents = False
    cell.Value = v
    Application.EnableEvents = eventsBefore
    HandleSheetChange ws, cell
End Sub

'--------------------------------------------------------------
' 結果シート
'--------------------------------------------------------------
Private Sub PrepareResultSheet()

    Dim alertsBefore As Boolean

    Set mResultWs = FindSheet(RESULT_SHEET)
    If Not mResultWs Is Nothing Then
        alertsBefore = Application.DisplayAlerts
        Application.DisplayAlerts = False
        mResultWs.Delete
        Application.DisplayAlerts = alertsBefore
    End If
    Set mResultWs = AddSheetAfter(ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    mResultWs.Name = RESULT_SHEET
    mResultWs.Range("A3:D3").Value = Array("No.", "テスト", "結果", "詳細")
    mResultWs.Range("A3:D3").Font.Bold = True
    mResultWs.Columns("A").ColumnWidth = 5
    mResultWs.Columns("B").ColumnWidth = 55
    mResultWs.Columns("C").ColumnWidth = 7
    mResultWs.Columns("D").ColumnWidth = 80
    mRow = 4
    mOkCount = 0
    mNgCount = 0
    mSkipCount = 0

End Sub

Private Sub Check(testName As String, passed As Boolean, Optional detail As String = "")
    mResultWs.Cells(mRow, 1).Value = mOkCount + mNgCount + mSkipCount + 1
    mResultWs.Cells(mRow, 2).Value = testName
    If passed Then
        mResultWs.Cells(mRow, 3).Value = "OK"
        mOkCount = mOkCount + 1
    Else
        mResultWs.Cells(mRow, 3).Value = "NG"
        mResultWs.Cells(mRow, 3).Interior.Color = COLOR_NG
        mResultWs.Cells(mRow, 4).Value = Replace(detail, vbLf, " ")
        mNgCount = mNgCount + 1
    End If
    mRow = mRow + 1
End Sub

Private Sub Skip(testName As String, reason As String)
    mResultWs.Cells(mRow, 1).Value = mOkCount + mNgCount + mSkipCount + 1
    mResultWs.Cells(mRow, 2).Value = testName
    mResultWs.Cells(mRow, 3).Value = "SKIP"
    mResultWs.Cells(mRow, 4).Value = reason
    mSkipCount = mSkipCount + 1
    mRow = mRow + 1
End Sub
