Attribute VB_Name = "modTest"
Option Explicit

'==============================================================
' QA試験 自動テスト
'   「テスト実行」を実行すると、テスト用の架空品名（ZZTEST-…）をマスタの末尾に
'   一時的に追加して、検査項目一覧からのマスタ作成・明細シートの作成・
'   OK欄の入力規則・検査者押印と判定・表紙への反映・総合判定・
'   入力変更時の消去・承認者印の条件・ロット番号の変更・明細削除／表紙削除／一括削除・
'   品名変更時の作り直し・品名の候補・印鑑の大きさ・保存（Excel/PDF・【NG】・NGメールの下書き）・Base登録を確認し、
'   結果を「テスト結果」シートに書き出す。
'   ・メッセージは表示せずに記録する（確認ダイアログは「はい」で進める）。
'   ・押す人・保存形式を選ぶフォームは出さず、選んだ後の処理を直接呼ぶ。
'   ・NG メールの下書きはブラウザで開かず、内容だけ確認する。
'   ・テスト用の DB（本番の DB ファイルのコピー）・保存ファイルは一時フォルダ（%TEMP%）に作る。
'   ・終了後、テスト用のマスタ行を削除し、表紙の品名・ロット番号・状態・日付を元に戻す
'     （明細シートは作り直されるため、入力中のデータは消える）。
'==============================================================

Private Const RESULT_SHEET As String = "テスト結果"
Private Const TEST_PERSON As String = "田中"
Private Const TEST_LOT As String = "LOT-T001"

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

    TestMasterImport   ' テスト用の行を足す前に確認する

    firstTestRow = MasterLastRow(ms) + 1
    AddTestMaster ms, firstTestRow

    On Error GoTo ErrHandler
    TestFunctions
    TestBuildSheets
    TestStampAreaSize
    TestInspectorStamp
    TestJudge
    TestRoundStamp
    TestLotChange
    TestProductChange
    TestClear
    TestSave
    TestSaveNg
    TestBaseDB
    On Error GoTo 0

Cleanup:
    On Error Resume Next
    ms.Range(ms.Rows(firstTestRow), ms.Rows(firstTestRow + 30)).Delete
    SetValue cs, cs.Range(COVER_LOT_CELL), saved(1)
    SetValue cs, cs.Range(COVER_STATE_CELL), saved(2)
    SetValue cs, cs.Range(COVER_DATE_CELL), saved(3)
    SetValue cs, cs.Range(COVER_PRODUCT_CELL), saved(0)
    gTestMode = False
    gTestBaseDbPath = ""
    gTestSaveFolder = ""
    gTestSaveTypes = ""
    gTestMailBody = ""
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

    ' 品名, 検査名, 数, 検査のルール, 基準値1, 基準値2, 単位, 基準の表示
    data = Array( _
        Array("ZZTEST-1", "検査A", 3, "以上", "300", "", "℃", ""), _
        Array("ZZTEST-1", "検査Z", 0, "以上", "1", "", "", ""), _
        Array("ZZTEST-1", "検査B", 2, "一致", "C,F", "", "", "A~GのうちC，Fが合格"), _
        Array("ZZTEST-2", "検査C", 1, "範囲", "50", "100", "", ""), _
        Array("ZZTEST-BAD", "検査X", 1, "大きい", "1", "", "", ""))

    For i = 0 To UBound(data)
        ms.Cells(firstRow + i, MCOL_PRODUCT).Value = data(i)(0)
        ms.Cells(firstRow + i, MCOL_INSPECTION).Value = data(i)(1)
        ms.Cells(firstRow + i, MCOL_COUNT).Value = data(i)(2)
        ms.Cells(firstRow + i, MCOL_METHOD).Validation.Delete   ' 誤ったルールを書くため
        ms.Cells(firstRow + i, MCOL_METHOD).Value = data(i)(3)
        ms.Cells(firstRow + i, MCOL_STD1).Value = data(i)(4)
        ms.Cells(firstRow + i, MCOL_STD2).Value = data(i)(5)
        ms.Cells(firstRow + i, MCOL_UNIT).Value = data(i)(6)
        ms.Cells(firstRow + i, MCOL_DISPLAY).Value = data(i)(7)
    Next i

End Sub

'--------------------------------------------------------------
' セットアップ：検査項目一覧からマスタが作られ、検査項目一覧が削除されている
'--------------------------------------------------------------
Private Sub TestMasterImport()

    Dim ms As Worksheet
    Dim r As Long

    Set ms = ThisWorkbook.Worksheets(MASTER_SHEET)
    If Not ProductExists("CCC-010-50") Then
        Skip "マスタ：検査項目一覧から作成", "マスタに CCC-010-50 がありません（検査項目一覧を入れずにセットアップした）"
        Exit Sub
    End If

    Check "マスタ：検査項目一覧シートは削除済み", FindSheet(ITEM_LIST_SHEET) Is Nothing
    Check "マスタ：CCC-010-50 は5検査（測長あり）", GetInspectionNames("CCC-010-50").Count = 5, _
        "検査数=" & GetInspectionNames("CCC-010-50").Count
    Check "マスタ：AAA-111 は4検査（数0の測長は無し）", GetInspectionNames("AAA-111").Count = 4, _
        "検査数=" & GetInspectionNames("AAA-111").Count
    r = GetMasterRow("CCC-010-50", "耐熱テスト")
    Check "マスタ：耐熱テストの数・ルール・基準", r > 0 And GetCountAt(ms, r) = 10 _
        And CellText(ms.Cells(r, MCOL_METHOD).Value) = METHOD_MIN And CellText(ms.Cells(r, MCOL_STD1).Value) = "300", _
        "行=" & r
    r = GetMasterRow("DDD-111-49", "抜き取りテスト")
    Check "マスタ：抜き取りテストの数（DDD-111-49=3）と基準の表示", r > 0 And GetCountAt(ms, r) = 3 _
        And StandardTextAt(ms, r) = "A~GのうちC，Fが合格", "行=" & r

End Sub

'--------------------------------------------------------------
' 小さな関数
'--------------------------------------------------------------
Private Sub TestFunctions()
    Check "苗字：「山田　太郎」→ 山田", SurnameOf("山田　太郎") = "山田", SurnameOf("山田　太郎")
    Check "苗字：「田中」→ 田中", SurnameOf("田中") = "田中"
    Check "OK欄：全角・小文字の ｏｋ も OK", NormalizeOkNg("ｏｋ") = RESULT_OK And NormalizeOkNg("ng") = RESULT_NG And NormalizeOkNg("x") = ""
    Check "基準の表示（範囲・単位あり）", FormatStandard(METHOD_RANGE, "200", "300", "℃") = "200℃以上300℃以下で合格", _
        FormatStandard(METHOD_RANGE, "200", "300", "℃")
End Sub

'--------------------------------------------------------------
' 品名とロット番号の入力で、明細シートと表紙の一覧が作られる
'--------------------------------------------------------------
Private Sub TestBuildSheets()

    Dim cs As Worksheet
    Dim wa As Worksheet
    Dim wb As Worksheet
    Dim f As String

    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)

    ' テスト前の入力が残っていても結果が変わらないよう、先に品名を消して明細シートを消す
    SetValue cs, cs.Range(COVER_PRODUCT_CELL), ""
    SetValue cs, cs.Range(COVER_LOT_CELL), ""

    ' 品名だけ（ロット番号は空）で作られる
    gTestLog = ""
    SetValue cs, cs.Range(COVER_PRODUCT_CELL), "ZZTEST-1"
    Set wa = FindSheet(DETAIL_PREFIX & "検査A")
    Set wb = FindSheet(DETAIL_PREFIX & "検査B")
    Check "品名だけで：明細シートが2枚（数0の検査Zは無し）", GetDetailSheets().Count = 2 And Not wa Is Nothing _
        And Not wb Is Nothing And FindSheet(DETAIL_PREFIX & "検査Z") Is Nothing, "明細シート数=" & GetDetailSheets().Count
    If wa Is Nothing Or wb Is Nothing Then Exit Sub
    Check "品名だけで：明細のロット番号は空欄（メッセージなし）", wa.Range(DETAIL_LOT_CELL).Value = "" And gTestLog = "", gTestLog

    ' 後からロット番号を入れると明細に表示される
    SetValue cs, cs.Range(COVER_LOT_CELL), TEST_LOT

    Check "表紙：No. と検査名が自動で並ぶ", CoverListCount() = 2 _
        And cs.Cells(COVER_LIST_FIRST_ROW, COVER_COL_NO).Value = 1 And cs.Cells(COVER_LIST_FIRST_ROW, COVER_COL_NAME).Value = "検査A" _
        And cs.Cells(COVER_LIST_FIRST_ROW + 1, COVER_COL_NO).Value = 2 And cs.Cells(COVER_LIST_FIRST_ROW + 1, COVER_COL_NAME).Value = "検査B"
    Check "表紙：総合判定欄は F4（一覧の下には無い）", cs.Range("F3").Value = "総合判定" _
        And OverallCell().Address = "$F$4" And IsBlankValue(cs.Cells(COVER_LIST_FIRST_ROW + 3, COVER_COL_NAME).Value)

    Check "明細：見出し・品名・ロット番号", wa.Range(DETAIL_TITLE_CELL).Value = "１．検査A" _
        And wa.Range(DETAIL_PRODUCT_CELL).Value = "ZZTEST-1" And wa.Range(DETAIL_LOT_CELL).Value = TEST_LOT _
        And wb.Range(DETAIL_TITLE_CELL).Value = "２．検査B", wa.Range(DETAIL_TITLE_CELL).Value
    Check "明細：数の分だけ No. がある（検査A=3、検査B=2）", DetailItemCount(wa) = 3 And DetailItemCount(wb) = 2, _
        DetailItemCount(wa) & " / " & DetailItemCount(wb)
    Check "明細：基準値の表示", wa.Cells(DETAIL_FIRST_ROW, DETAIL_COL_STD).Value = "300℃以上で合格" _
        And wb.Cells(DETAIL_FIRST_ROW, DETAIL_COL_STD).Value = "A~GのうちC，Fが合格", wa.Cells(DETAIL_FIRST_ROW, DETAIL_COL_STD).Value
    Check "明細：見出しは 基準値／No.／検査項目／OK欄（NG欄なし）", _
        wa.Cells(DETAIL_HEADER_ROW, DETAIL_COL_STD).Value = "基準値" And wa.Cells(DETAIL_HEADER_ROW, DETAIL_COL_NO).Value = "No." _
        And wa.Cells(DETAIL_HEADER_ROW, DETAIL_COL_VALUE).Value = "検査項目" And wa.Cells(DETAIL_HEADER_ROW, DETAIL_COL_OK).Value = "OK欄" _
        And IsBlankValue(wa.Cells(DETAIL_HEADER_ROW, DETAIL_COL_OK + 1).Value)

    On Error Resume Next
    f = ""
    f = wa.Cells(DETAIL_FIRST_ROW + 2, DETAIL_COL_OK).Validation.Formula1
    On Error GoTo 0
    Check "明細：OK欄の入力規則は OK,NG の一覧", f = "OK,NG", f

End Sub

'--------------------------------------------------------------
' 印鑑欄の大きさ（印鑑と余白が収まる）
'--------------------------------------------------------------
Private Sub TestStampAreaSize()

    Dim cs As Worksheet
    Dim wa As Worksheet
    Dim need As Double

    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)
    Set wa = FindSheet(DETAIL_PREFIX & "検査A")
    If wa Is Nothing Then
        Skip "印鑑欄の大きさ", "明細シートがありません"
        Exit Sub
    End If
    need = STAMP_SIZE + STAMP_MARGIN * 2

    Check "印鑑欄：明細の検査者欄", AreaFits(wa.Range(DETAIL_STAMP_CELL).MergeArea, need), AreaText(wa.Range(DETAIL_STAMP_CELL).MergeArea)
    Check "印鑑欄：表紙の担当者印（検査ごと）", AreaFits(cs.Cells(COVER_LIST_FIRST_ROW, COVER_COL_STAMP), need), _
        AreaText(cs.Cells(COVER_LIST_FIRST_ROW, COVER_COL_STAMP))
    Check "印鑑欄：表紙の承認者・担当者", AreaFits(cs.Range(COVER_APPROVER_CELL).MergeArea, need) _
        And AreaFits(cs.Range(COVER_CHARGE_CELL).MergeArea, need), AreaText(cs.Range(COVER_APPROVER_CELL).MergeArea)

End Sub

Private Function AreaFits(area As Range, need As Double) As Boolean
    AreaFits = (area.Width >= need - 0.5 And area.Height >= need - 0.5)
End Function

Private Function AreaText(area As Range) As String
    AreaText = "幅=" & Format(area.Width, "0.0") & " 高さ=" & Format(area.Height, "0.0")
End Function

'--------------------------------------------------------------
' 検査者押印：OK欄の確認・データ印・判定・表紙への反映・入力変更での消去
'--------------------------------------------------------------
Private Sub TestInspectorStamp()

    Dim cs As Worksheet
    Dim wa As Worksheet
    Dim wb As Worksheet
    Dim ok As Boolean
    Dim r As Long
    Dim grp As Shape
    Dim dateText As String

    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)
    Set wa = FindSheet(DETAIL_PREFIX & "検査A")
    Set wb = FindSheet(DETAIL_PREFIX & "検査B")
    If wa Is Nothing Or wb Is Nothing Then
        Skip "検査者押印のテスト", "明細シートがありません"
        Exit Sub
    End If
    r = DETAIL_FIRST_ROW

    ' OK欄が空欄なら押さない
    SetValue wa, wa.Cells(r, DETAIL_COL_VALUE), 310
    SetValue wa, wa.Cells(r, DETAIL_COL_OK), "OK"
    gTestLog = ""
    ok = ApplyInspectorStamp(wa, TEST_PERSON, Now)
    Check "OK欄に空欄があると押印・判定しない", Not ok And InStr(gTestLog, "OK欄が空欄です") > 0 _
        And Not ShapeExists(wa, DETAIL_STAMP_NAME) And GetDetailResult(wa) = "", gTestLog

    ' 検査A：すべて OK → 総合判定 OK
    SetValue wa, wa.Cells(r + 1, DETAIL_COL_VALUE), 300
    SetValue wa, wa.Cells(r + 1, DETAIL_COL_OK), "OK"
    SetValue wa, wa.Cells(r + 2, DETAIL_COL_VALUE), 305
    SetValue wa, wa.Cells(r + 2, DETAIL_COL_OK), "OK"
    gTestLog = ""
    ok = ApplyInspectorStamp(wa, TEST_PERSON, Now)
    Check "検査A：すべて OK → 総合判定に OK", ok And GetDetailResult(wa) = RESULT_OK, gTestLog
    Check "検査A：検査者欄にデータ印", ShapeExists(wa, DETAIL_STAMP_NAME) And GetStampName(wa, DETAIL_STAMP_NAME) = TEST_PERSON
    Check "表紙：検査A の判定結果 OK と担当者印が自動で入る", cs.Cells(FindCoverRow("検査A"), COVER_COL_RESULT).Value = RESULT_OK _
        And ShapeExists(cs, INSP_STAMP_PREFIX & "検査A") And GetStampName(cs, INSP_STAMP_PREFIX & "検査A") = TEST_PERSON
    Check "表紙：総合判定は未判定の検査があれば空欄", GetOverall() = "", GetOverall()

    ' データ印の中身（部署／日付／名前の3段と2本の線）と大きさ
    If ShapeExists(wa, DETAIL_STAMP_NAME) Then
        Set grp = wa.Shapes(DETAIL_STAMP_NAME)
        dateText = StampPartText(grp, DETAIL_STAMP_NAME & "_日付")
        Check "データ印：円・線2本・文字3つ", grp.GroupItems.Count = 6, "部品数=" & grp.GroupItems.Count
        Check "データ印：上から部署・日付（押した日）・名前", _
            StampPartText(grp, DETAIL_STAMP_NAME & "_部署") = GetDepartment(TEST_PERSON) _
            And dateText = Format(Date, "yyyy.mm.dd") _
            And StampPartText(grp, DETAIL_STAMP_NAME & "_名前") = TEST_PERSON _
            And grp.GroupItems(DETAIL_STAMP_NAME & "_部署").Top < grp.GroupItems(DETAIL_STAMP_NAME & "_日付").Top _
            And grp.GroupItems(DETAIL_STAMP_NAME & "_日付").Top < grp.GroupItems(DETAIL_STAMP_NAME & "_名前").Top, _
            StampPartText(grp, DETAIL_STAMP_NAME & "_部署") & " / " & dateText
        Check "データ印：大きさ（" & STAMP_SIZE & "pt）", Abs(grp.Width - STAMP_SIZE) < 1 And Abs(grp.Height - STAMP_SIZE) < 1, _
            Format(grp.Width, "0.0") & " x " & Format(grp.Height, "0.0")
    Else
        Skip "データ印の中身", "検査者欄に印がありません"
    End If

    ' 検査B：NG が1つ → 総合判定 NG（押印はする）
    SetValue wb, wb.Cells(r, DETAIL_COL_VALUE), "C"
    SetValue wb, wb.Cells(r, DETAIL_COL_OK), "OK"
    SetValue wb, wb.Cells(r + 1, DETAIL_COL_VALUE), "B"
    SetValue wb, wb.Cells(r + 1, DETAIL_COL_OK), "NG"
    ok = ApplyInspectorStamp(wb, TEST_PERSON, Now)
    Check "検査B：NG が1つ → 総合判定 NG、押印はする", ok And GetDetailResult(wb) = RESULT_NG And ShapeExists(wb, DETAIL_STAMP_NAME)
    Check "表紙：総合判定は1つでも NG なら NG", GetOverall() = RESULT_NG And cs.Cells(FindCoverRow("検査B"), COVER_COL_RESULT).Value = RESULT_NG, _
        GetOverall()

    ' 承認者印は NG があると押せない
    gTestLog = ""
    承認者押印
    Check "承認者印：NG があると押せない", Not ShapeExists(cs, APPROVER_STAMP_NAME) And InStr(gTestLog, "NG の検査があるため") > 0, gTestLog

    ' 押印の後に OK欄を変えると、その検査の判定結果と印鑑が消える
    gTestLog = ""
    SetValue wb, wb.Cells(r + 1, DETAIL_COL_OK), "OK"
    Check "入力変更：明細の判定結果・検査者印を消去", GetDetailResult(wb) = "" And Not ShapeExists(wb, DETAIL_STAMP_NAME)
    Check "入力変更：表紙の判定結果・担当者印を消去しメッセージ", cs.Cells(FindCoverRow("検査B"), COVER_COL_RESULT).Value = "" _
        And Not ShapeExists(cs, INSP_STAMP_PREFIX & "検査B") And GetOverall() = "" _
        And InStr(gTestLog, "判定結果と印鑑を消去しました") > 0, gTestLog
    Check "入力変更：ほかの検査（検査A）はそのまま", GetDetailResult(wa) = RESULT_OK And ShapeExists(cs, INSP_STAMP_PREFIX & "検査A")

    ' 検査値も基準に合う値にして押し直すと全検査 OK → 総合判定 OK
    SetValue wb, wb.Cells(r + 1, DETAIL_COL_VALUE), "F"
    ok = ApplyInspectorStamp(wb, TEST_PERSON, Now)
    Check "押し直し：全検査 OK → 総合判定 OK", ok And GetDetailResult(wb) = RESULT_OK And GetOverall() = RESULT_OK, GetOverall()

End Sub

'--------------------------------------------------------------
' 基準での判定：ルールごとの判定・判定できない値・OK欄との食い違いで押印しない
'--------------------------------------------------------------
Private Sub TestJudge()

    Dim wa As Worksheet
    Dim wb As Worksheet
    Dim ok As Boolean
    Dim r As Long

    ' ルールごとの判定（境界の値は合格）
    '   VBA の And は両辺をすべて評価するため、1つの式で判定関数を何度も呼ばず、JudgeText で1回ずつ結果を取る
    Check "判定：以上 300 → 300 は OK、299.9 は NG", JudgeText(METHOD_MIN, "300", "", 300) = RESULT_OK _
        And JudgeText(METHOD_MIN, "300", "", 299.9) = RESULT_NG
    Check "判定：以下 0.5 → 0.5 は OK、0.7 は NG", JudgeText(METHOD_MAX, "0.5", "", 0.5) = RESULT_OK _
        And JudgeText(METHOD_MAX, "0.5", "", 0.7) = RESULT_NG
    Check "判定：範囲 50～100 → 50・100 は OK、101 は NG", JudgeText(METHOD_RANGE, "50", "100", 50) = RESULT_OK _
        And JudgeText(METHOD_RANGE, "50", "100", 100) = RESULT_OK _
        And JudgeText(METHOD_RANGE, "50", "100", 101) = RESULT_NG
    Check "判定：全角の数字も数値として判定", JudgeText(METHOD_MIN, "300", "", "３１０") = RESULT_OK
    Check "判定：一致 C,F → f・Ｆ・「 C 」は OK、B は NG", JudgeText(METHOD_MATCH, "C,F", "", "f") = RESULT_OK _
        And JudgeText(METHOD_MATCH, "C,F", "", "Ｆ") = RESULT_OK _
        And JudgeText(METHOD_MATCH, "C,F", "", " C ") = RESULT_OK _
        And JudgeText(METHOD_MATCH, "C,F", "", "B") = RESULT_NG
    Check "判定：一致の区切りは全角カンマ・読点も可", JudgeText(METHOD_MATCH, "C，F、G", "", "G") = RESULT_OK
    Check "判定：数値のルールで文字・空欄は判定できない", InStr(JudgeText(METHOD_MIN, "300", "", "abc"), "数値ではありません") > 0 _
        And InStr(JudgeText(METHOD_MIN, "300", "", ""), "空欄") > 0, _
        JudgeText(METHOD_MIN, "300", "", "abc") & " / " & JudgeText(METHOD_MIN, "300", "", "")
    Check "判定：知らないルールは判定できない", InStr(JudgeText("大きい", "1", "", 2), "判定できません") > 0, JudgeText("大きい", "1", "", 2)

    ' 押印：OK欄が基準と食い違うと押さない
    Set wa = FindSheet(DETAIL_PREFIX & "検査A")
    Set wb = FindSheet(DETAIL_PREFIX & "検査B")
    If wa Is Nothing Or wb Is Nothing Then
        Skip "基準と食い違うときの押印", "明細シートがありません"
        Exit Sub
    End If
    r = DETAIL_FIRST_ROW

    SetValue wa, wa.Cells(r, DETAIL_COL_VALUE), 290        ' 300 以上なので NG のはず
    SetValue wa, wa.Cells(r, DETAIL_COL_OK), RESULT_OK
    gTestLog = ""
    ok = ApplyInspectorStamp(wa, TEST_PERSON, Now)
    Check "押印：数値が NG なのに OK欄が OK なら押さない（セルと理由を表示）", Not ok And GetDetailResult(wa) = "" _
        And Not ShapeExists(wa, DETAIL_STAMP_NAME) And InStr(gTestLog, "OK欄が基準と合いません") > 0 _
        And InStr(gTestLog, wa.Cells(r, DETAIL_COL_OK).Address(False, False) & "：検査値「290」は基準では NG、OK欄は OK") > 0, gTestLog

    SetValue wa, wa.Cells(r, DETAIL_COL_VALUE), "abc"
    gTestLog = ""
    ok = ApplyInspectorStamp(wa, TEST_PERSON, Now)
    Check "押印：数値のルールで検査値が数値でなければ押さない", Not ok And InStr(gTestLog, "基準で判定できない検査項目があります") > 0 _
        And InStr(gTestLog, "数値ではありません") > 0, gTestLog

    SetValue wa, wa.Cells(r, DETAIL_COL_VALUE), 290
    SetValue wa, wa.Cells(r, DETAIL_COL_OK), RESULT_NG
    ok = ApplyInspectorStamp(wa, TEST_PERSON, Now)
    Check "押印：OK欄が基準と合っていれば押す（290 → NG）", ok And GetDetailResult(wa) = RESULT_NG

    SetValue wb, wb.Cells(r, DETAIL_COL_VALUE), "ｃ"           ' 全角小文字でも C と一致
    SetValue wb, wb.Cells(r, DETAIL_COL_OK), RESULT_OK
    ok = ApplyInspectorStamp(wb, TEST_PERSON, Now)
    Check "押印：一致は全角小文字「ｃ」でも OK として押す", ok And GetDetailResult(wb) = RESULT_OK, gTestLog

    ' 元に戻す（以降のテストのため全検査 OK にする）
    SetValue wa, wa.Cells(r, DETAIL_COL_VALUE), 310
    SetValue wa, wa.Cells(r, DETAIL_COL_OK), RESULT_OK
    ApplyInspectorStamp wa, TEST_PERSON, Now

End Sub

'--------------------------------------------------------------
' JudgeValue を1回呼び、判定できたら "OK"／"NG"、できなければ "判定不可：理由" を返す
'--------------------------------------------------------------
Private Function JudgeText(method As String, std1 As String, std2 As String, v As Variant) As String
    Dim res As String
    Dim reason As String
    If JudgeValue(method, std1, std2, v, res, reason) Then
        JudgeText = res
    Else
        JudgeText = "判定不可：" & reason
    End If
End Function

Private Function StampPartText(grp As Shape, partName As String) As String
    On Error Resume Next
    StampPartText = grp.GroupItems(partName).TextFrame2.TextRange.text
    On Error GoTo 0
End Function

'--------------------------------------------------------------
' 丸印（承認者・担当者）：苗字・大きさ・入力変更での消去
'   電子印鑑シートに田中の印影がある場合のみ
'--------------------------------------------------------------
Private Sub TestRoundStamp()

    Dim cs As Worksheet
    Dim wa As Worksheet
    Dim shp As Shape

    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)
    Set wa = FindSheet(DETAIL_PREFIX & "検査A")
    If wa Is Nothing Or CalcOverall() <> RESULT_OK Then
        Skip "丸印（承認者・担当者）", "全検査 OK の状態になっていません"
        Exit Sub
    End If

    ' 承認者押印は条件を満たせばフォームを出すため、フォームの後の処理を直接呼ぶ
    SetStampTarget cs.Range(COVER_APPROVER_CELL).MergeArea, APPROVER_STAMP_NAME
    ApplyApproverStamp TEST_PERSON
    SetStampTarget cs.Range(COVER_CHARGE_CELL).MergeArea, CHARGE_STAMP_NAME
    ApplyApproverStamp TEST_PERSON
    Check "丸印：承認者・担当者に押せる", ShapeExists(cs, APPROVER_STAMP_NAME) And ShapeExists(cs, CHARGE_STAMP_NAME)
    If ShapeExists(cs, APPROVER_STAMP_NAME) Then
        Set shp = cs.Shapes(APPROVER_STAMP_NAME)
        Check "丸印：苗字だけ・押した人を読み取れる", Trim(shp.TextFrame2.TextRange.text) = SurnameOf(TEST_PERSON) _
            And GetStampName(cs, APPROVER_STAMP_NAME) = TEST_PERSON, shp.TextFrame2.TextRange.text
        Check "丸印：データ印と同じ大きさ", Abs(shp.Width - STAMP_SIZE) < 1 And Abs(shp.Height - STAMP_SIZE) < 1, _
            Format(shp.Width, "0.0") & " x " & Format(shp.Height, "0.0")
    End If

    ' 氏名に空白があれば苗字だけ（担当者印を押し直して確かめる）
    SetStampTarget cs.Range(COVER_CHARGE_CELL).MergeArea, CHARGE_STAMP_NAME
    ApplyApproverStamp "山田　太郎"
    Check "丸印：「山田　太郎」は「山田」と押す", ShapeExists(cs, CHARGE_STAMP_NAME) _
        And Trim(cs.Shapes(CHARGE_STAMP_NAME).TextFrame2.TextRange.text) = "山田"

    ' 検査項目を変えると承認者印・担当者印も消える
    gTestLog = ""
    SetValue wa, wa.Cells(DETAIL_FIRST_ROW, DETAIL_COL_VALUE), 311
    Check "入力変更：表紙の承認者印・担当者印も消去", Not ShapeExists(cs, APPROVER_STAMP_NAME) _
        And Not ShapeExists(cs, CHARGE_STAMP_NAME) And GetDetailResult(wa) = "", gTestLog
    ApplyInspectorStamp wa, TEST_PERSON, Now

End Sub

'--------------------------------------------------------------
' ロット番号だけ変えた：明細シートはそのまま、ロット番号の表示だけ変わる
'--------------------------------------------------------------
Private Sub TestLotChange()

    Dim cs As Worksheet
    Dim wa As Worksheet

    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)
    Set wa = FindSheet(DETAIL_PREFIX & "検査A")
    If wa Is Nothing Then
        Skip "ロット番号の変更", "明細シートがありません"
        Exit Sub
    End If

    gTestLog = ""
    SetValue cs, cs.Range(COVER_LOT_CELL), "LOT-T002"
    Set wa = FindSheet(DETAIL_PREFIX & "検査A")
    Check "ロット番号の変更：判定・印鑑はそのまま、表示だけ変わる", Not wa Is Nothing And GetDetailResult(wa) = RESULT_OK _
        And ShapeExists(cs, INSP_STAMP_PREFIX & "検査A") And wa.Range(DETAIL_LOT_CELL).Value = "LOT-T002" And gTestLog = "", gTestLog
    SetValue cs, cs.Range(COVER_LOT_CELL), TEST_LOT

End Sub

'--------------------------------------------------------------
' 明細削除・表紙削除・一括削除
'--------------------------------------------------------------
Private Sub TestClear()

    Dim cs As Worksheet
    Dim wa As Worksheet
    Dim wb As Worksheet

    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)
    PrepareJudgedProduct
    SetStampTarget cs.Range(COVER_APPROVER_CELL).MergeArea, APPROVER_STAMP_NAME
    ApplyApproverStamp TEST_PERSON
    Set wa = FindSheet(DETAIL_PREFIX & "検査A")
    Set wb = FindSheet(DETAIL_PREFIX & "検査B")
    If wa Is Nothing Or wb Is Nothing Then
        Skip "削除のテスト", "明細シートがありません"
        Exit Sub
    End If

    ' 確認で「いいえ」なら何も消さない
    gTestAnswer = vbNo
    wa.Activate
    明細削除
    Check "明細削除：確認で「いいえ」なら消さない", HasAnyInput(wa) And GetDetailResult(wa) = RESULT_OK _
        And ShapeExists(cs, APPROVER_STAMP_NAME)
    gTestAnswer = vbYes

    ' 明細削除：そのシートの入力と印鑑、表紙の関係する結果と印鑑が消える
    wa.Activate
    明細削除
    Check "明細削除：入力・検査者印・総合判定を消す", Not HasAnyInput(wa) And GetDetailResult(wa) = "" _
        And Not ShapeExists(wa, DETAIL_STAMP_NAME)
    Check "明細削除：表紙のその検査の結果・担当者印と承認者印を消す", cs.Cells(FindCoverRow("検査A"), COVER_COL_RESULT).Value = "" _
        And Not ShapeExists(cs, INSP_STAMP_PREFIX & "検査A") And Not ShapeExists(cs, APPROVER_STAMP_NAME) And GetOverall() = ""
    Check "明細削除：ほかの検査（検査B）はそのまま", HasAnyInput(wb) And GetDetailResult(wb) = RESULT_OK _
        And ShapeExists(cs, INSP_STAMP_PREFIX & "検査B")

    ' 表紙削除：表紙の入力と印鑑が消え、品名が消えるので明細シートも消える
    cs.Activate
    表紙削除
    Check "表紙削除：品名・ロット番号・状態・日付を消す", CellText(cs.Range(COVER_PRODUCT_CELL).Value) = "" _
        And CellText(cs.Range(COVER_LOT_CELL).Value) = "" And CellText(cs.Range(COVER_STATE_CELL).Value) = "" _
        And CellText(cs.Range(COVER_DATE_CELL).Value) = ""
    Check "表紙削除：明細シート・一覧・印鑑・総合判定も消える", GetDetailSheets().Count = 0 And CoverListCount() = 0 _
        And Not ShapeExists(cs, INSP_STAMP_PREFIX & "検査B") And GetOverall() = ""

    ' 一括削除：全部消して品名の欄を選ぶ
    PrepareJudgedProduct
    SetValue cs, cs.Range(COVER_STATE_CELL), "量産"
    SetStampTarget cs.Range(COVER_CHARGE_CELL).MergeArea, CHARGE_STAMP_NAME
    ApplyApproverStamp TEST_PERSON
    一括削除
    Check "一括削除：表紙・明細の入力と印鑑を全部消す", GetDetailSheets().Count = 0 And CoverListCount() = 0 _
        And CellText(cs.Range(COVER_PRODUCT_CELL).Value) = "" And CellText(cs.Range(COVER_STATE_CELL).Value) = "" _
        And Not ShapeExists(cs, CHARGE_STAMP_NAME) And GetOverall() = ""
    Check "一括削除：終わったら品名の欄を選ぶ", ActiveSheet.Name = COVER_SHEET And ActiveCell.Address = cs.Range(COVER_PRODUCT_CELL).Address, _
        ActiveSheet.Name & "!" & ActiveCell.Address

End Sub

'--------------------------------------------------------------
' 品名変更：明細シートの作り直し・マスタに無い品名（候補）・マスタの誤り
'--------------------------------------------------------------
Private Sub TestProductChange()

    Dim cs As Worksheet
    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)

    gTestLog = ""
    SetValue cs, cs.Range(COVER_PRODUCT_CELL), "ZZTEST-2"
    Check "品名変更：明細シートを作り直す（1枚）", GetDetailSheets().Count = 1 And CoverListCount() = 1 _
        And FindSheet(DETAIL_PREFIX & "検査A") Is Nothing And Not FindSheet(DETAIL_PREFIX & "検査C") Is Nothing, _
        "明細シート数=" & GetDetailSheets().Count
    Check "品名変更：判定結果・印鑑を消去しメッセージ", GetOverall() = "" And Not ShapeExists(cs, INSP_STAMP_PREFIX & "検査A") _
        And InStr(gTestLog, "明細シートを作り直しました") > 0, gTestLog

    gTestLog = ""
    SetValue cs, cs.Range(COVER_PRODUCT_CELL), "zztest"
    Check "マスタに無い品名：明細シートなし＋候補を表示", GetDetailSheets().Count = 0 And CoverListCount() = 0 _
        And InStr(gTestLog, "マスタシートにありません") > 0 And InStr(gTestLog, "ZZTEST-1") > 0 _
        And InStr(gTestLog, "ZZTEST-2") > 0, gTestLog

    gTestLog = ""
    SetValue cs, cs.Range(COVER_PRODUCT_CELL), "ZZTEST-BAD"
    Check "マスタの誤り（検査のルール）：明細シートなし＋メッセージ", GetDetailSheets().Count = 0 _
        And InStr(gTestLog, "検査のルールは") > 0, gTestLog

    gTestLog = ""
    SetValue cs, cs.Range(COVER_PRODUCT_CELL), ""
    Check "品名を空にすると明細シートなし（メッセージなし）", GetDetailSheets().Count = 0 And gTestLog = "", gTestLog

End Sub

'--------------------------------------------------------------
' 保存（一時フォルダに Excel / PDF を保存。全検査 OK なのでファイル名に【NG】は付かない）
'--------------------------------------------------------------
Private Sub TestSave()

    Dim cs As Worksheet
    Dim histRows As Long
    Dim baseName As String
    Dim errText As String

    If Not PrepareTempSaveFolder("保存（Excel/PDF）") Then Exit Sub

    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)
    PrepareJudgedProduct
    SetValue cs, cs.Range(COVER_DATE_CELL), DateSerial(2026, 1, 2)

    Check "ファイル名：品名_ロット_日付", BuildFileBaseName(baseName, errText) And baseName = "ZZTEST-1_" & TEST_LOT & "_20260102", baseName & errText

    histRows = HistoryLastRow()

    ' Excel と PDF の両方を選ぶ
    gTestSaveTypes = HIST_TYPE_EXCEL & "," & HIST_TYPE_PDF
    gTestLog = ""
    検査表保存
    Check "保存：両方選ぶと Excel（.xlsx）は Excel フォルダ、PDF は PDF フォルダに保存", _
        Dir(gTestSaveFolder & "\" & EXCEL_FOLDER_NAME & "\" & baseName & ".xlsx") <> "" _
        And Dir(gTestSaveFolder & "\" & PDF_FOLDER_NAME & "\" & baseName & ".pdf") <> "" _
        And Dir(gTestSaveFolder & "\" & baseName & ".xlsx") = "", gTestLog
    Check "保存：OK のときはファイル名に【NG】を付けず、メールの下書きも作らない", _
        Dir(gTestSaveFolder & "\" & EXCEL_FOLDER_NAME & "\" & NG_FILE_PREFIX & baseName & ".xlsx") = "" And InStr(gTestLog, "[NGメール下書き]") = 0, gTestLog
    Check "保存履歴に2行（Excel・PDF）記録", HistoryLastRow() = HistoryBase(histRows) + 2 _
        And HistoryType(HistoryBase(histRows) + 1) = HIST_TYPE_EXCEL And HistoryType(HistoryBase(histRows) + 2) = HIST_TYPE_PDF, _
        "最終行=" & HistoryLastRow()

    ' PDF だけを選ぶ（同じ名前があるので上書き確認に「はい」）
    gTestSaveTypes = HIST_TYPE_PDF
    gTestLog = ""
    検査表保存
    Check "保存：PDF だけ選ぶと PDF だけ保存（上書き確認あり）", HistoryLastRow() = HistoryBase(histRows) + 3 _
        And HistoryType(HistoryBase(histRows) + 3) = HIST_TYPE_PDF And InStr(gTestLog, "上書きしますか") > 0, gTestLog

    RestoreHistory histRows
    gTestSaveTypes = ""

End Sub

'--------------------------------------------------------------
' NG のときの保存：Excel・PDF の両方を【NG】付きで保存 → NG メールの下書き（テスト中はブラウザを開かない）
'   2回目以降はファイル名の末尾に「_2回目」「_3回目」
'--------------------------------------------------------------
Private Sub TestSaveNg()

    Dim cs As Worksheet
    Dim ss As Worksheet
    Dim histRows As Long
    Dim baseName As String
    Dim errText As String
    Dim savedTo As Variant
    Dim savedSubject As Variant
    Dim savedFrom As Variant
    Dim ngXls As String
    Dim ngPdf As String
    Dim b As Long

    Set ss = FindSheet(SETTING_SHEET)
    If ss Is Nothing Then
        Skip "NG の保存・メール", "設定シートがありません"
        Exit Sub
    End If
    If Not PrepareTempSaveFolder("NG の保存・メール") Then Exit Sub

    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)
    PrepareJudgedProduct
    SetValue cs, cs.Range(COVER_DATE_CELL), DateSerial(2026, 1, 2)
    BuildFileBaseName baseName, errText
    If Not MakeTestNg() Then Exit Sub
    Check "NG：総合判定が NG", GetOverall() = RESULT_NG, GetOverall()

    savedTo = ss.Range(SETTING_MAIL_TO_CELL).Value
    savedSubject = ss.Range(SETTING_MAIL_SUBJECT_CELL).Value
    savedFrom = ss.Range(SETTING_MAIL_FROM_CELL).Value
    ss.Range(SETTING_MAIL_SUBJECT_CELL).Value = ""   ' 既定の件名で確かめる
    ss.Range(SETTING_MAIL_FROM_CELL).Value = "qa-sender@example.com"   ' 下書きを開く Gmail アカウント
    histRows = HistoryLastRow()
    b = HistoryBase(histRows)
    ngXls = gTestSaveFolder & "\" & EXCEL_FOLDER_NAME & "\" & NG_FILE_PREFIX & baseName
    ngPdf = gTestSaveFolder & "\" & PDF_FOLDER_NAME & "\" & NG_FILE_PREFIX & baseName

    ' 1回目：形式に Excel だけを指定しても両方保存し、その後にメールの下書き
    ss.Range(SETTING_MAIL_TO_CELL).Value = "qa-test@example.com"
    gTestSaveTypes = HIST_TYPE_EXCEL
    gTestLog = ""
    gTestMailBody = ""
    検査表保存
    Check "NG 1回目：Excel・PDF とも先頭に【NG】を付けて保存（形式の指定に関係なく両方）", Dir(ngXls & ".xlsx") <> "" _
        And Dir(ngPdf & ".pdf") <> "", gTestLog
    Check "NG 1回目：保存履歴は Excel・PDF・メール下書きの順", HistoryLastRow() = b + 3 _
        And HistoryType(b + 1) = HIST_TYPE_EXCEL And HistoryType(b + 2) = HIST_TYPE_PDF _
        And HistoryType(b + 3) = HIST_TYPE_MAIL, "最終行=" & HistoryLastRow()
    Check "NG 1回目：NG メールの下書き（宛先・件名）", InStr(gTestLog, "[NGメール下書き] qa-test@example.com / 【NG】QA試験 ZZTEST-1 " & TEST_LOT) > 0, gTestLog
    Check "NG メール下書き：Gmail の作成画面の URL（宛先・アカウントを URL 用に変換）", _
        InStr(gTestMailBody, "URL：https://mail.google.com/mail/?view=cm&fs=1&to=qa-test%40example.com&su=") > 0 _
        And InStr(gTestMailBody, "&authuser=qa-sender%40example.com") > 0, gTestMailBody
    Check "NG メール下書き：添付する PDF・本文のファイル名", InStr(gTestMailBody, "添付するファイル：" & ngPdf & ".pdf") > 0 _
        And InStr(gTestMailBody, NG_FILE_PREFIX & baseName & ".pdf（添付）") > 0, gTestMailBody
    Check "NG メールの本文：NG 回数・NG の検査・No.・検査値・検査員", InStr(gTestMailBody, "NG 回数：1回目") > 0 _
        And InStr(gTestMailBody, "・検査B（検査員：" & TEST_PERSON & "）") > 0 _
        And InStr(gTestMailBody, "No.2：B") > 0, gTestMailBody

    ' 2回目：末尾に「_2回目」
    gTestLog = ""
    gTestMailBody = ""
    検査表保存
    Check "NG 2回目：Excel・PDF とも末尾に「_2回目」", Dir(ngXls & "_2回目.xlsx") <> "" _
        And Dir(ngPdf & "_2回目.pdf") <> "" And InStr(gTestMailBody, "NG 回数：2回目") > 0 _
        And InStr(gTestMailBody, "添付するファイル：" & ngPdf & "_2回目.pdf") > 0, gTestLog
    Check "NG 2回目：保存履歴が3行増える", HistoryLastRow() = b + 6, "最終行=" & HistoryLastRow()

    ' 3回目：宛先なし → 下書きは作らず、保存はする
    ss.Range(SETTING_MAIL_TO_CELL).Value = ""
    gTestLog = ""
    検査表保存
    Check "NG 3回目：宛先が空欄なら下書きを作らずに「_3回目」で保存し、メッセージ", Dir(ngXls & "_3回目.xlsx") <> "" _
        And Dir(ngPdf & "_3回目.pdf") <> "" And InStr(gTestLog, "宛先が設定されていません") > 0 _
        And InStr(gTestLog, "[NGメール下書き]") = 0 And HistoryLastRow() = b + 8, gTestLog

    ss.Range(SETTING_MAIL_TO_CELL).Value = savedTo
    ss.Range(SETTING_MAIL_SUBJECT_CELL).Value = savedSubject
    ss.Range(SETTING_MAIL_FROM_CELL).Value = savedFrom
    RestoreHistory histRows
    gTestSaveTypes = ""

End Sub

'--------------------------------------------------------------
' 検査B の No.2 を NG にして押し直す（総合判定が NG になる）。できなければ SKIP して False
'--------------------------------------------------------------
Private Function MakeTestNg() As Boolean
    Dim wb As Worksheet
    Dim r As Long
    Set wb = FindSheet(DETAIL_PREFIX & "検査B")
    If wb Is Nothing Then
        Skip "NG にする", "明細シートがありません"
        Exit Function
    End If
    r = DETAIL_FIRST_ROW + 1
    SetValue wb, wb.Cells(r, DETAIL_COL_VALUE), "B"
    SetValue wb, wb.Cells(r, DETAIL_COL_OK), RESULT_NG
    ApplyInspectorStamp wb, TEST_PERSON, Now
    MakeTestNg = (GetOverall() = RESULT_NG)
End Function

'--------------------------------------------------------------
' Base 登録（本番の DB ファイルを一時フォルダにコピーして使う）
'   DB ファイルが無い、または LibreOffice が無い場合は SKIP
'--------------------------------------------------------------
Private Sub TestBaseDB()

    Dim cs As Worksheet
    Dim srcPath As String
    Dim histRows As Long
    Dim n As Long

    gTestBaseDbPath = ""
    srcPath = GetBaseDbPath()
    If Dir(srcPath) = "" Then
        Skip "Base 登録", "DB ファイルがありません（" & srcPath & "）"
        Exit Sub
    End If
    If Not PrepareTempSaveFolder("Base 登録") Then Exit Sub

    gTestBaseDbPath = Environ("TEMP") & "\QA試験_テストDB_" & Format(Now, "yyyymmdd_hhnnss") & ".odb"
    FileCopy srcPath, gTestBaseDbPath

    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)
    PrepareJudgedProduct
    SetValue cs, cs.Range(COVER_STATE_CELL), "量産"
    SetValue cs, cs.Range(COVER_DATE_CELL), DateSerial(2026, 1, 2)
    histRows = HistoryLastRow()

    ' Excel 保存がまだなら登録しない
    gTestLog = ""
    検査表DB登録
    Check "DB登録：Excel 保存が無いと登録しない", InStr(gTestLog, "Excel 保存がありません") > 0, gTestLog

    gTestSaveTypes = HIST_TYPE_EXCEL
    検査表保存
    gTestSaveTypes = ""

    gTestLog = ""
    検査表DB登録
    If InStr(gTestLog, "LibreOffice を起動できません") > 0 Then
        Skip "Base 登録", "LibreOffice を起動できません"
        GoTo Cleanup
    End If
    Check "DB登録：明細シートの数（2件）を登録", InStr(gTestLog, "登録しました（2件、判定 OK）") > 0, gTestLog
    n = CountBaseRows(gTestBaseDbPath, "ZZTEST-1", TEST_LOT)
    Check "DB登録：DB を開き直しても2件ある（.odb に保存されている）", n = 2, "件数=" & n
    Check "DB登録：保存履歴に DB の行", HistoryType(HistoryLastRow()) = HIST_TYPE_DB

    ' 登録済み：「いいえ」なら何もしない
    gTestAnswer = vbNo
    gTestLog = ""
    検査表DB登録
    gTestAnswer = vbYes
    n = CountBaseRows(gTestBaseDbPath, "ZZTEST-1", TEST_LOT)
    Check "DB登録：登録済みで「いいえ」なら登録しない", InStr(gTestLog, "既に登録されています") > 0 _
        And InStr(gTestLog, "登録しました") = 0 And n = 2, gTestLog & " 件数=" & n

    ' 登録済み：「はい」なら登録し直す（件数は増えない）
    gTestLog = ""
    検査表DB登録
    n = CountBaseRows(gTestBaseDbPath, "ZZTEST-1", TEST_LOT)
    Check "DB登録：登録済みで「はい」なら登録し直す（2件のまま）", InStr(gTestLog, "登録しました（2件") > 0 And n = 2, _
        gTestLog & " 件数=" & n
    Check "DB登録：OK の保存は NG回数 0 で登録", CountBaseRows(gTestBaseDbPath, "ZZTEST-1", TEST_LOT, 0) = 2

    ' NG にしただけ（保存していない）なら登録しない
    If Not MakeTestNg() Then GoTo Cleanup
    gTestLog = ""
    検査表DB登録
    Check "DB登録：保存の後に判定が変わっていたら登録しない", InStr(gTestLog, "入力が変わっています") > 0, gTestLog

    ' NG 1回目・2回目を保存して登録 → 回数ごとに別の行で残る（上書き確認は出ない）
    検査表保存
    gTestLog = ""
    検査表DB登録
    Check "DB登録：NG 1回目は別の行で登録（上書き確認なし）", InStr(gTestLog, "登録しました（2件、NG 1回目）") > 0 _
        And InStr(gTestLog, "既に登録されています") = 0, gTestLog

    ' NG のまま入力だけ変えた（総合判定は NG のまま）→ 保存し直すまで登録しない
    SetValue cs, cs.Range(COVER_STATE_CELL), "試作"
    gTestLog = ""
    検査表DB登録
    Check "DB登録：総合判定が同じでも、保存の後に入力が変わっていたら登録しない", InStr(gTestLog, "入力が変わっています") > 0 _
        And InStr(gTestLog, "登録しました") = 0, gTestLog
    検査表保存
    gTestLog = ""
    検査表DB登録
    n = CountBaseRows(gTestBaseDbPath, "ZZTEST-1", TEST_LOT)
    Check "DB登録：NG 2回目も別の行（OK・NG1・NG2 で合計6件）", InStr(gTestLog, "登録しました（2件、NG 2回目）") > 0 _
        And n = 6 And CountBaseRows(gTestBaseDbPath, "ZZTEST-1", TEST_LOT, 1) = 2 _
        And CountBaseRows(gTestBaseDbPath, "ZZTEST-1", TEST_LOT, 2) = 2, gTestLog & " 件数=" & n

    mResultWs.Cells(mRow, 2).Value = "（テスト用DB：" & gTestBaseDbPath & "）"
    mRow = mRow + 1

Cleanup:
    gTestAnswer = vbYes
    RestoreHistory histRows

End Sub

'--------------------------------------------------------------
' 一時フォルダにテスト用の保存先を作る。作れなければ SKIP して False
'--------------------------------------------------------------
Private Function PrepareTempSaveFolder(testName As String) As Boolean
    Dim tempDir As String
    tempDir = Environ("TEMP")
    If tempDir = "" Then
        Skip testName, "一時フォルダ（%TEMP%）がありません"
        Exit Function
    End If
    gTestSaveFolder = tempDir & "\QA試験_テスト保存_" & Format(Now, "yyyymmdd_hhnnss") & "_" & mRow
    MkDir gTestSaveFolder
    mResultWs.Cells(mRow, 2).Value = "（テスト保存先：" & gTestSaveFolder & "）"
    mRow = mRow + 1
    PrepareTempSaveFolder = True
End Function

'--------------------------------------------------------------
' 保存履歴：最終行（シートが無ければ 0）・テスト前の最終行（見出しのみなら 1）・種類
'--------------------------------------------------------------
Private Function HistoryLastRow() As Long
    Dim histWs As Worksheet
    Set histWs = FindSheet(HISTORY_SHEET_NAME)
    If histWs Is Nothing Then Exit Function
    HistoryLastRow = histWs.Cells(histWs.Rows.Count, 1).End(xlUp).Row
End Function

Private Function HistoryBase(histRows As Long) As Long
    HistoryBase = IIf(histRows = 0, 1, histRows)
End Function

Private Function HistoryType(r As Long) As String
    Dim histWs As Worksheet
    Set histWs = FindSheet(HISTORY_SHEET_NAME)
    If histWs Is Nothing Or r < 2 Then Exit Function
    HistoryType = CellText(histWs.Cells(r, 2).Value)
End Function

'--------------------------------------------------------------
' テストで追加した保存履歴を消す（テストで作られたシートならシートごと消す）
'--------------------------------------------------------------
Private Sub RestoreHistory(histRows As Long)
    Dim histWs As Worksheet
    Set histWs = FindSheet(HISTORY_SHEET_NAME)
    If histWs Is Nothing Then Exit Sub
    If histRows = 0 Then
        Application.DisplayAlerts = False
        histWs.Delete
        Application.DisplayAlerts = True
    ElseIf HistoryLastRow() > histRows Then
        histWs.Range(histWs.Rows(histRows + 1), histWs.Rows(HistoryLastRow())).Delete
    End If
End Sub

'--------------------------------------------------------------
' ZZTEST-1 を全検査押印済み（OK）にする
'--------------------------------------------------------------
Private Sub PrepareJudgedProduct()

    Dim cs As Worksheet
    Dim item As Variant
    Dim ws As Worksheet
    Dim r As Long

    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)
    SetValue cs, cs.Range(COVER_LOT_CELL), TEST_LOT
    SetValue cs, cs.Range(COVER_PRODUCT_CELL), "ZZTEST-1"

    For Each item In GetDetailSheets()
        Set ws = item
        For r = DETAIL_FIRST_ROW To DETAIL_FIRST_ROW + DetailItemCount(ws) - 1
            ' 基準に合う値（検査A：300以上、検査B：C か F）
            If InspectionNameOf(ws) = "検査A" Then
                SetValue ws, ws.Cells(r, DETAIL_COL_VALUE), 300 + r
            Else
                SetValue ws, ws.Cells(r, DETAIL_COL_VALUE), "C"
            End If
            SetValue ws, ws.Cells(r, DETAIL_COL_OK), RESULT_OK
        Next r
        ApplyInspectorStamp ws, TEST_PERSON, Now
    Next item

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


