Attribute VB_Name = "modBaseDB"
Option Explicit

'==============================================================
' QA試験 LibreOffice Base 登録
'   表紙の「DB登録」ボタンで、明細シートの判定データを Base（埋め込み HSQLDB の .odb）に登録する。
'   ・DBファイル：設定シートの C3 に書いたパス。空欄ならこのブックと同じフォルダの QA試験DB.odb。
'     ファイルとテーブル「検査データ」は事前に作っておく（自動では作らない）。
'   ・登録内容：明細シート1枚（検査1つ）につき1行。
'     検査日時は、保存履歴にある同じ品名・ロットの最新の Excel 保存の日時。
'     NG回数は、その保存が NG なら何回目の NG か（1～）、OK の保存なら 0。
'     検査員は明細シートの検査者印の名前（電子印鑑シートの氏名）。
'   ・NG は回数ごとに別の行で残す。同じ品名＋ロット番号＋NG回数が登録済みのときだけ、
'     上書きするか確認する（はい：その回数の行を削除して登録し直す）。
'   ・最新の Excel 保存の後に表紙・明細の入力を変えた場合や、その保存の OK／NG が今の総合判定と違う場合は
'     登録しない（先に保存してもらう）。
'   ・削除と登録は1つのトランザクションで行い、エラーが起きたら何も登録しない。
'   ・接続方法：LibreOffice を COM（UNO）で操作し、.odb を非表示で開いて専用の接続（getIsolatedConnection）で SQL を実行する。
'     埋め込み DB は .odb を保存（store）しないと変更が残らないため、登録後に保存する。
'   ・非表示で開いた .odb は必ず閉じる。LibreOffice 本体は、ほかに開いている文書が無いときだけ終了する。
'==============================================================

Private Const DB_FILE_NAME As String = "QA試験DB.odb"
Private Const DB_TABLE As String = """検査データ"""
Private Const DB_MAX_ITEMS As Long = 10        ' 検査項目の列の数（検査項目1～10）
Private Const DB_TEXT_MAX As Long = 100        ' 品名・ロット番号・状態・検査名・検査員の長さ
Private Const DB_ITEM_MAX As Long = 20         ' 検査項目の長さ

' com.sun.star.sdbc.DataType
Private Const SQL_VARCHAR As Long = 12

Public gTestBaseDbPath As String   ' テスト用：設定すると設定シートより優先する

' 接続中のオブジェクト（CloseBase で閉じる）
Private mSM As Object
Private mDesktop As Object
Private mDoc As Object
Private mConn As Object

'--------------------------------------------------------------
' DB ファイルのパス
'--------------------------------------------------------------
Public Function GetBaseDbPath() As String

    Dim ss As Worksheet
    Dim p As String

    If gTestBaseDbPath <> "" Then
        GetBaseDbPath = gTestBaseDbPath
        Exit Function
    End If

    Set ss = FindSheet(SETTING_SHEET)
    If Not ss Is Nothing Then p = CellText(ss.Range(SETTING_DB_CELL).Value)
    If p = "" Then p = ThisWorkbook.Path & "\" & DB_FILE_NAME
    GetBaseDbPath = p

End Function

'--------------------------------------------------------------
' 「DB登録」ボタン
'--------------------------------------------------------------
Sub 検査表DB登録()

    Const DIALOG_TITLE As String = "DB登録"
    Dim inspTime As Date
    Dim ngCount As Long
    Dim errText As String
    Dim dbPath As String
    Dim registered As Long

    If Not CheckBeforeRegister(inspTime, ngCount, errText) Then
        ShowMsg "DB に登録できません。" & vbCrLf & errText, vbExclamation, DIALOG_TITLE
        Exit Sub
    End If

    dbPath = GetBaseDbPath()
    If Not CheckDbFile(dbPath, errText) Then
        ShowMsg errText, vbExclamation, DIALOG_TITLE
        Exit Sub
    End If

    registered = RegisterToBase(dbPath, inspTime, ngCount, DIALOG_TITLE)
    If registered > 0 Then
        WriteSaveHistory HIST_TYPE_DB, dbPath
        SaveThisWorkbook
        ShowMsg "登録しました（" & registered & "件、" & NgCountText(ngCount) & "）。" & vbCrLf & dbPath, vbInformation, DIALOG_TITLE
    End If

End Sub

'--------------------------------------------------------------
' 登録前のチェック（シートの内容）。問題があれば errText に内容を入れて False
'--------------------------------------------------------------
Private Function CheckBeforeRegister(ByRef inspTime As Date, ByRef ngCount As Long, ByRef errText As String) As Boolean

    Dim cs As Worksheet
    Dim item As Variant
    Dim ws As Worksheet
    Dim r As Long
    Dim baseName As String
    Dim productName As String
    Dim lotNumber As String

    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)
    errText = ""

    If CoverListCount() = 0 Or GetDetailSheets().Count = 0 Then
        errText = "・明細シートがありません。表紙の品名を入力してください。" & vbCrLf
        Exit Function
    End If

    ' 品名・ロット・日付（ファイル名と同じチェック）
    BuildFileBaseName baseName, errText

    productName = CellText(cs.Range(COVER_PRODUCT_CELL).Value)
    lotNumber = CellText(cs.Range(COVER_LOT_CELL).Value)
    If Len(productName) > DB_TEXT_MAX Then errText = errText & "・品名が " & DB_TEXT_MAX & " 文字を超えています" & vbCrLf
    If Len(lotNumber) > DB_TEXT_MAX Then errText = errText & "・ロット番号が " & DB_TEXT_MAX & " 文字を超えています" & vbCrLf
    If Len(CellText(cs.Range(COVER_STATE_CELL).Value)) > DB_TEXT_MAX Then errText = errText & "・状態が " & DB_TEXT_MAX & " 文字を超えています" & vbCrLf

    If Not AllInspectionsJudged() Then errText = errText & "・判定が済んでいない検査があります" & vbCrLf

    For Each item In GetDetailSheets()
        Set ws = item
        If GetStampName(ws, DETAIL_STAMP_NAME) = "" Then
            errText = errText & "・「" & InspectionNameOf(ws) & "」に検査者印がありません" & vbCrLf
        End If
        If DetailItemCount(ws) > DB_MAX_ITEMS Then
            errText = errText & "・「" & InspectionNameOf(ws) & "」の項目が " & DB_MAX_ITEMS & " を超えています（DB の検査項目は" & DB_MAX_ITEMS & "まで）" & vbCrLf
        End If
        For r = DETAIL_FIRST_ROW To DETAIL_FIRST_ROW + DetailItemCount(ws) - 1
            If Len(CellText(ws.Cells(r, DETAIL_COL_VALUE).Value)) > DB_ITEM_MAX Then
                errText = errText & "・「" & InspectionNameOf(ws) & "」の " & ws.Cells(r, DETAIL_COL_VALUE).Address(False, False) & _
                          " が " & DB_ITEM_MAX & " 文字を超えています" & vbCrLf
            End If
        Next r
    Next item

    If productName <> "" And lotNumber <> "" Then
        If Not GetLatestExcelSave(productName, lotNumber, inspTime, ngCount) Then
            errText = errText & "・この品名・ロットの Excel 保存がありません。先に「保存」で Excel を保存してください" & vbCrLf
        ElseIf IsInputChanged() Or ((ngCount > 0) <> (GetOverall() = RESULT_NG)) Then
            errText = errText & "・最新の Excel 保存の後に入力が変わっています（総合判定：" & GetOverall() & "）。" & _
                      "先に「保存」で保存し直してください" & vbCrLf
        End If
    End If

    CheckBeforeRegister = (errText = "")

End Function

'--------------------------------------------------------------
' DB ファイルのチェック（場所・有無・開かれていないか）
'--------------------------------------------------------------
Private Function CheckDbFile(dbPath As String, ByRef errText As String) As Boolean

    Dim folderPath As String
    Dim fileName As String
    Dim p As Long

    If LCase(Left(dbPath, 8)) = "https://" Or LCase(Left(dbPath, 7)) = "http://" Then
        errText = "DB ファイルの場所が OneDrive などの Web 上になっているため使えません。" & vbCrLf & _
                  "PC のフォルダに置いて、設定シートの " & SETTING_DB_CELL & " にパスを書いてください。" & vbCrLf & dbPath
        Exit Function
    End If

    If Dir(dbPath) = "" Then
        errText = "DB ファイルが見つかりません。" & vbCrLf & dbPath & vbCrLf & _
                  "設定シートの " & SETTING_DB_CELL & "（空欄ならこのブックと同じフォルダの " & DB_FILE_NAME & "）を確認してください。"
        Exit Function
    End If

    ' LibreOffice のロックファイル（.~lock.ファイル名#）があれば、開かれている
    p = InStrRev(dbPath, "\")
    folderPath = Left(dbPath, p)
    fileName = Mid(dbPath, p + 1)
    If Dir(folderPath & ".~lock." & fileName & "#", vbHidden + vbNormal) <> "" Then
        errText = "DB ファイルが開かれています。Base を閉じてから再実行してください。" & vbCrLf & _
                  "（Base を閉じても出る場合は、ロックファイル .~lock." & fileName & "# が残っています。手順書を見て削除してください）"
        Exit Function
    End If

    CheckDbFile = True

End Function

'--------------------------------------------------------------
' Base に登録する。登録した件数を返す（中止・失敗は 0。メッセージはこの中で出す）
'--------------------------------------------------------------
Private Function RegisterToBase(dbPath As String, inspTime As Date, ngCount As Long, title As String) As Long

    Dim cs As Worksheet
    Dim item As Variant
    Dim ws As Worksheet
    Dim productName As String
    Dim lotNumber As String
    Dim inTrans As Boolean
    Dim committed As Boolean
    Dim cnt As Long
    Dim errMsg As String
    Dim stage As String

    On Error GoTo ErrHandler

    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)
    productName = CellText(cs.Range(COVER_PRODUCT_CELL).Value)
    lotNumber = CellText(cs.Range(COVER_LOT_CELL).Value)

    stage = "open"
    If Not OpenBase(dbPath, title) Then GoTo Finish

    stage = "sql"
    mConn.setAutoCommit False
    inTrans = True

    ' 同じ品名＋ロット番号＋NG回数が登録済みなら、上書きするか確認する
    If CountRegistered(productName, lotNumber, ngCount) > 0 Then
        If ShowMsg("この品名・ロット番号（" & NgCountText(ngCount) & "）は既に登録されています。上書きしますか？" & vbCrLf & _
                   "品名：" & productName & vbCrLf & "ロット番号：" & lotNumber, _
                   vbYesNo + vbQuestion + vbDefaultButton2, title) = vbNo Then
            mConn.rollback
            inTrans = False
            GoTo Finish
        End If
        DeleteRegistered productName, lotNumber, ngCount
    End If

    For Each item In GetDetailSheets()
        Set ws = item
        InsertInspection ws, inspTime, ngCount
        cnt = cnt + 1
    Next item

    ' 登録した件数を数え直す
    If CountRegistered(productName, lotNumber, ngCount) <> cnt Then
        mConn.rollback
        inTrans = False
        ShowMsg "登録に失敗しました（何も登録されていません）。" & vbCrLf & "登録した件数が明細シートの数と一致しません。", _
                vbExclamation, title
        GoTo Finish
    End If

    mConn.commit
    inTrans = False
    committed = True

    ' .odb に保存する（埋め込み DB は保存しないと変更が残らない）
    stage = "store"
    On Error Resume Next
    mConn.flush          ' 埋め込み HSQLDB の内容を文書に書き出す（対応していなければ何もしない）
    Err.Clear
    On Error GoTo ErrHandler
    mDoc.store

    RegisterToBase = cnt

Finish:
    CloseBase
    Exit Function

ErrHandler:
    errMsg = Err.Description
    On Error Resume Next
    If inTrans Then mConn.rollback
    On Error GoTo 0
    CloseBase
    Select Case stage
        Case "store"
            ShowMsg "DB ファイルへの保存に失敗したため、登録できませんでした。もう一度登録してください。" & vbCrLf & errMsg, _
                    vbExclamation, title
        Case "sql"
            ShowMsg "登録に失敗しました（何も登録されていません）。" & vbCrLf & errMsg, vbExclamation, title
        Case Else
            ShowMsg "DB に接続できませんでした。" & vbCrLf & errMsg, vbExclamation, title
    End Select
    RegisterToBase = 0

End Function

'--------------------------------------------------------------
' 明細シート1枚分を登録する
'--------------------------------------------------------------
Private Sub InsertInspection(ws As Worksheet, inspTime As Date, ngCount As Long)

    Dim cs As Worksheet
    Dim st As Object
    Dim sql As String
    Dim cols As String
    Dim marks As String
    Dim i As Long
    Dim n As Long
    Dim dateValue As Date

    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)

    cols = """品名"", ""ロット番号"", ""状態"", ""日付"", ""検査日時"", ""NG回数"", ""検査名"", ""判定結果"""
    marks = "?, ?, ?, ?, ?, ?, ?, ?"
    For i = 1 To DB_MAX_ITEMS
        cols = cols & ", ""検査項目" & i & """"
        marks = marks & ", ?"
    Next i
    cols = cols & ", ""検査員"""
    marks = marks & ", ?"
    sql = "INSERT INTO " & DB_TABLE & " (" & cols & ") VALUES (" & marks & ")"

    Set st = mConn.prepareStatement(sql)
    st.setString 1, CellText(cs.Range(COVER_PRODUCT_CELL).Value)
    st.setString 2, CellText(cs.Range(COVER_LOT_CELL).Value)
    SetTextOrNull st, 3, CellText(cs.Range(COVER_STATE_CELL).Value)
    dateValue = CDate(cs.Range(COVER_DATE_CELL).Value)
    st.setDate 4, UnoDate(dateValue)
    st.setTimestamp 5, UnoDateTime(inspTime)
    st.setInt 6, ngCount
    st.setString 7, InspectionNameOf(ws)
    st.setString 8, GetDetailResult(ws)

    n = DetailItemCount(ws)
    For i = 1 To DB_MAX_ITEMS
        If i <= n Then
            SetTextOrNull st, 8 + i, CellText(ws.Cells(DETAIL_FIRST_ROW + i - 1, DETAIL_COL_VALUE).Value)
        Else
            st.setNull 8 + i, SQL_VARCHAR
        End If
    Next i
    st.setString 9 + DB_MAX_ITEMS, GetStampName(ws, DETAIL_STAMP_NAME)

    st.executeUpdate
    st.Close

End Sub

' メッセージ用の NG 回数の表示
Private Function NgCountText(ngCount As Long) As String
    If ngCount = 0 Then
        NgCountText = "判定 OK"
    Else
        NgCountText = "NG " & ngCount & "回目"
    End If
End Function

Private Sub SetTextOrNull(st As Object, index As Long, text As String)
    If text = "" Then
        st.setNull index, SQL_VARCHAR
    Else
        st.setString index, text
    End If
End Sub

'--------------------------------------------------------------
' 品名＋ロット番号（＋NG回数）の登録件数（接続中の DB）。ngCount が -1 なら全回数
'--------------------------------------------------------------
Private Function CountRegistered(productName As String, lotNumber As String, ngCount As Long) As Long
    Dim st As Object
    Dim rs As Object
    Dim sql As String
    sql = "SELECT COUNT(*) FROM " & DB_TABLE & " WHERE ""品名"" = ? AND ""ロット番号"" = ?"
    If ngCount >= 0 Then sql = sql & " AND ""NG回数"" = ?"
    Set st = mConn.prepareStatement(sql)
    st.setString 1, productName
    st.setString 2, lotNumber
    If ngCount >= 0 Then st.setInt 3, ngCount
    Set rs = st.executeQuery()
    ' next は VBA の予約語のため CallByName で呼ぶ。COUNT は getInt で読む
    If CallByName(rs, "next", VbMethod) Then CountRegistered = rs.getInt(1)
    st.Close
End Function

Private Sub DeleteRegistered(productName As String, lotNumber As String, ngCount As Long)
    Dim st As Object
    Set st = mConn.prepareStatement("DELETE FROM " & DB_TABLE & " WHERE ""品名"" = ? AND ""ロット番号"" = ? AND ""NG回数"" = ?")
    st.setString 1, productName
    st.setString 2, lotNumber
    st.setInt 3, ngCount
    st.executeUpdate
    st.Close
End Sub

'--------------------------------------------------------------
' テスト用：品名＋ロット番号（＋NG回数。省略すると全回数）の登録件数（DB を開いて数え、閉じる）。読めなければ -1
'--------------------------------------------------------------
Public Function CountBaseRows(dbPath As String, productName As String, lotNumber As String, _
    Optional ngCount As Long = -1) As Long
    On Error GoTo ErrHandler
    CountBaseRows = -1
    If Not OpenBase(dbPath, "DB登録") Then Exit Function
    CountBaseRows = CountRegistered(productName, lotNumber, ngCount)
    CloseBase
    Exit Function
ErrHandler:
    CloseBase
    CountBaseRows = -1
End Function

'--------------------------------------------------------------
' LibreOffice を起動し、.odb を非表示で開いて接続する。できたら True
'--------------------------------------------------------------
Private Function OpenBase(dbPath As String, title As String) As Boolean

    Dim sm As Object
    Dim url As String
    Dim args(0) As Variant
    Dim errMsg As String

    Set mSM = Nothing
    Set mDesktop = Nothing
    Set mDoc = Nothing
    Set mConn = Nothing

    On Error Resume Next
    Set sm = CreateObject("com.sun.star.ServiceManager")
    If Not sm Is Nothing Then Set mDesktop = sm.createInstance("com.sun.star.frame.Desktop")
    errMsg = Err.Description
    On Error GoTo 0
    If mDesktop Is Nothing Then
        ShowMsg "LibreOffice を起動できません。インストールを確認してください。" & vbCrLf & errMsg, vbExclamation, title
        Exit Function
    End If

    Set mSM = sm
    url = ToFileUrl(sm, dbPath)

    If IsDocumentOpen(url) Then
        ShowMsg "DB ファイルが開かれています。Base を閉じてから再実行してください。", vbExclamation, title
        Exit Function
    End If

    Set args(0) = sm.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
    args(0).Name = "Hidden"
    args(0).Value = True

    Set mDoc = mDesktop.loadComponentFromURL(url, "_blank", 0, args)
    If mDoc Is Nothing Then Err.Raise vbObjectError + 1, "OpenBase", "DB ファイルを開けませんでした。" & dbPath
    ' getConnection の接続は Base が共用するため、トランザクション（setAutoCommit / commit）が使えない。
    ' 登録用に専用の接続を取る
    Set mConn = mDoc.DataSource.getIsolatedConnection("", "")

    OpenBase = True

End Function

'--------------------------------------------------------------
' 接続と .odb を閉じる。ほかに開いている文書が無ければ LibreOffice も終了する
'--------------------------------------------------------------
Private Sub CloseBase()

    On Error Resume Next
    If Not mConn Is Nothing Then mConn.Close
    If Not mDoc Is Nothing Then mDoc.Close True
    If Not mDesktop Is Nothing Then
        If Not mDesktop.getComponents().createEnumeration().hasMoreElements() Then mDesktop.terminate
    End If
    Set mConn = Nothing
    Set mDoc = Nothing
    Set mDesktop = Nothing
    Set mSM = Nothing
    Err.Clear
    On Error GoTo 0

End Sub

'--------------------------------------------------------------
' 同じ .odb が LibreOffice で開かれているか
'--------------------------------------------------------------
Private Function IsDocumentOpen(url As String) As Boolean

    Const MAX_COMPONENTS As Long = 500   ' エラーで抜けられなくなるのを防ぐ
    Dim en As Object
    Dim comp As Object
    Dim compUrl As String
    Dim guard As Long

    On Error Resume Next
    Set en = mDesktop.getComponents().createEnumeration()
    If en Is Nothing Then Exit Function
    Do While en.hasMoreElements() And guard < MAX_COMPONENTS
        guard = guard + 1
        Set comp = Nothing
        Set comp = en.nextElement()
        compUrl = ""
        If Not comp Is Nothing Then compUrl = comp.getURL()
        If compUrl <> "" And LCase(compUrl) = LCase(url) Then
            IsDocumentOpen = True
            Exit Do
        End If
    Loop
    Err.Clear
    On Error GoTo 0

End Function

'--------------------------------------------------------------
' Windows のパスを file:/// の URL にする（日本語・空白も変換する）
'--------------------------------------------------------------
Private Function ToFileUrl(sm As Object, path As String) As String
    Dim fcp As Object
    Set fcp = sm.createInstance("com.sun.star.ucb.FileContentProvider")
    ToFileUrl = fcp.getFileURLFromSystemPath("", path)
End Function

'--------------------------------------------------------------
' 日付・日時を UNO の構造体にする
'--------------------------------------------------------------
Private Function UnoDate(d As Date) As Object
    Dim u As Object
    Set u = mSM.Bridge_GetStruct("com.sun.star.util.Date")
    u.Year = Year(d)
    u.Month = Month(d)
    u.Day = Day(d)
    Set UnoDate = u
End Function

Private Function UnoDateTime(d As Date) As Object
    Dim u As Object
    Set u = mSM.Bridge_GetStruct("com.sun.star.util.DateTime")
    u.Year = Year(d)
    u.Month = Month(d)
    u.Day = Day(d)
    u.Hours = Hour(d)
    u.Minutes = Minute(d)
    u.Seconds = Second(d)
    u.NanoSeconds = 0
    Set UnoDateTime = u
End Function
