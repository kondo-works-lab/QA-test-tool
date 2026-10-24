Attribute VB_Name = "modAccess"
Option Explicit

'==============================================================
' QA試験 Access 連携
'   明細シートの「判定」で、判定データを Access（.accdb）に登録する。
'   ・DBファイル：設定シートの C3 に書いたパス。空欄ならこのブックと同じフォルダの
'     QA試験DB.accdb。ファイルやテーブルが無ければ自動で作成する
'     （Access 本体は不要。Office に付属する ACE OLEDB プロバイダを使う）。
'   ・登録内容
'       T_判定     ：判定1回（検査1つ）につき1行
'       T_判定明細 ：入力値1つにつき1行（判定IDで T_判定 とつながる）
'   ・再判定した場合は上書きせず追記する（履歴として残す）。
'   ・登録に失敗しても判定結果はシートに残し、メッセージのみ表示する。
'==============================================================

Private Const DB_FILE_NAME As String = "QA試験DB.accdb"
Private Const DB_PATH_CELL As String = "C3"            ' 設定シート
Private Const TABLE_JUDGE As String = "T_判定"
Private Const TABLE_DETAIL As String = "T_判定明細"

' ADO の定数（参照設定なしで使うため）
Private Const adVarWChar As Long = 202
Private Const adInteger As Long = 3
Private Const adDate As Long = 7
Private Const adParamInput As Long = 1
Private Const adCmdText As Long = 1
Private Const adSchemaTables As Long = 20

Public gTestDbPath As String   ' テスト用：設定すると設定シートより優先する

'--------------------------------------------------------------
' DB ファイルのパス
'--------------------------------------------------------------
Public Function GetDbPath() As String

    Dim ss As Worksheet
    Dim p As String

    If gTestDbPath <> "" Then
        GetDbPath = gTestDbPath
        Exit Function
    End If

    Set ss = FindSheet(SETTING_SHEET)
    If Not ss Is Nothing Then p = CellText(ss.Range(DB_PATH_CELL).Value)
    If p = "" Then p = ThisWorkbook.Path & "\" & DB_FILE_NAME
    GetDbPath = p

End Function

'--------------------------------------------------------------
' 明細シートの判定データを登録する。成功したら True
'--------------------------------------------------------------
Public Function RegisterJudgement(ws As Worksheet, pressTime As Date) As Boolean

    Dim cs As Worksheet
    Dim conn As Object
    Dim cmd As Object
    Dim inTrans As Boolean
    Dim judgeId As Long
    Dim dbPath As String
    Dim dateValue As Variant
    Dim sumRow As Long
    Dim r As Long
    Dim c As Long
    Dim n As Long
    Dim cell As Range
    Dim errMsg As String

    On Error GoTo ErrHandler

    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)
    dbPath = GetDbPath()
    Set conn = OpenDb(dbPath)
    EnsureTables conn

    conn.BeginTrans
    inTrans = True

    ' T_判定
    dateValue = cs.Range(COVER_DATE_CELL).Value
    If IsError(dateValue) Then
        dateValue = Null
    ElseIf IsDate(dateValue) Then
        dateValue = CDate(dateValue)
    Else
        dateValue = Null
    End If
    sumRow = DetailSumRow(ws)

    Set cmd = NewCommand(conn, "INSERT INTO [" & TABLE_JUDGE & "] " & _
        "([押下日時],[品名],[ロット],[状態],[日付],[検査名],[判定結果],[基準外個数],[不合格個数],[総合判定],[ブック名]) " & _
        "VALUES (?,?,?,?,?,?,?,?,?,?,?)")
    AddParam cmd, adDate, pressTime
    AddParam cmd, adVarWChar, TextOrNull(cs.Range(COVER_PRODUCT_CELL).Value)
    AddParam cmd, adVarWChar, TextOrNull(cs.Range(COVER_LOT_CELL).Value)
    AddParam cmd, adVarWChar, TextOrNull(cs.Range(COVER_STATE_CELL).Value)
    AddParam cmd, adDate, dateValue
    AddParam cmd, adVarWChar, TextOrNull(InspectionNameOf(ws))
    AddParam cmd, adVarWChar, TextOrNull(GetDetailResult(ws))
    AddParam cmd, adInteger, CLng(ws.Cells(sumRow, DETAIL_COL_METHOD).Value)
    AddParam cmd, adInteger, CLng(ws.Range(DETAIL_FAILCOUNT_CELL).Value)
    AddParam cmd, adVarWChar, TextOrNull(GetOverall())
    AddParam cmd, adVarWChar, TextOrNull(ThisWorkbook.Name)
    cmd.Execute

    judgeId = conn.Execute("SELECT @@IDENTITY")(0)

    ' T_判定明細（入力値1つにつき1行）
    For r = DETAIL_FIRST_ROW To DETAIL_FIRST_ROW + DetailItemCount(ws) - 1
        n = CLng(ws.Cells(r, DETAIL_COL_N).Value)
        For c = 1 To n
            Set cell = ws.Cells(r, DETAIL_COL_INPUT + c - 1)
            Set cmd = NewCommand(conn, "INSERT INTO [" & TABLE_DETAIL & "] " & _
                "([判定ID],[項目No],[項目名],[判定方式],[基準],[サンプルNo],[入力値],[合否]) " & _
                "VALUES (?,?,?,?,?,?,?,?)")
            AddParam cmd, adInteger, judgeId
            AddParam cmd, adInteger, CLng(ws.Cells(r, DETAIL_COL_NO).Value)
            AddParam cmd, adVarWChar, TextOrNull(ws.Cells(r, DETAIL_COL_ITEM).Value)
            AddParam cmd, adVarWChar, TextOrNull(ws.Cells(r, DETAIL_COL_METHOD).Value)
            AddParam cmd, adVarWChar, TextOrNull(ws.Cells(r, DETAIL_COL_STD).Value)
            AddParam cmd, adInteger, c
            AddParam cmd, adVarWChar, TextOrNull(cell.Value)
            If cell.Interior.Color = COLOR_NG Then
                AddParam cmd, adVarWChar, "基準外"
            Else
                AddParam cmd, adVarWChar, "合格"
            End If
            cmd.Execute
        Next c
    Next r

    conn.CommitTrans
    inTrans = False
    conn.Close

    RegisterJudgement = True
    Exit Function

ErrHandler:
    errMsg = Err.Description
    On Error Resume Next
    If inTrans Then conn.RollbackTrans
    If Not conn Is Nothing Then conn.Close
    On Error GoTo 0
    ShowMsg "Access への登録に失敗しました。（判定結果はシートに反映済みです）" & vbCrLf & _
            "DBファイル：" & dbPath & vbCrLf & errMsg, vbExclamation, "Access登録"

End Function

'--------------------------------------------------------------
' DB に接続する。ファイルが無ければ作成する
'   ACE 16.0 → 12.0 の順にプロバイダを試す
'--------------------------------------------------------------
Public Function OpenDb(dbPath As String) As Object

    Dim providers As Variant
    Dim i As Long
    Dim conn As Object
    Dim cat As Object
    Dim lastErr As String

    providers = Array("Microsoft.ACE.OLEDB.16.0", "Microsoft.ACE.OLEDB.12.0")

    For i = LBound(providers) To UBound(providers)
        On Error Resume Next
        Err.Clear
        If Dir(dbPath) = "" Then
            Set cat = CreateObject("ADOX.Catalog")
            cat.Create "Provider=" & providers(i) & ";Data Source=" & dbPath & ";"
            If Err.Number = 0 Then
                Set cat.ActiveConnection = Nothing
            End If
            Set cat = Nothing
        End If
        If Err.Number = 0 Then
            Set conn = CreateObject("ADODB.Connection")
            conn.Open "Provider=" & providers(i) & ";Data Source=" & dbPath & ";"
        End If
        If Err.Number = 0 Then
            On Error GoTo 0
            Set OpenDb = conn
            Exit Function
        End If
        lastErr = Err.Description
        Set conn = Nothing
        On Error GoTo 0
    Next i

    Err.Raise vbObjectError + 1, "OpenDb", "DBに接続できません。" & lastErr

End Function

'--------------------------------------------------------------
' テーブルが無ければ作成する
'--------------------------------------------------------------
Public Sub EnsureTables(conn As Object)

    If Not TableExists(conn, TABLE_JUDGE) Then
        conn.Execute "CREATE TABLE [" & TABLE_JUDGE & "] (" & _
            "[判定ID] COUNTER CONSTRAINT [PK_判定] PRIMARY KEY, " & _
            "[押下日時] DATETIME, [品名] TEXT(255), [ロット] TEXT(255), [状態] TEXT(255), " & _
            "[日付] DATETIME, [検査名] TEXT(255), [判定結果] TEXT(10), " & _
            "[基準外個数] LONG, [不合格個数] LONG, [総合判定] TEXT(10), [ブック名] TEXT(255))"
    End If

    If Not TableExists(conn, TABLE_DETAIL) Then
        conn.Execute "CREATE TABLE [" & TABLE_DETAIL & "] (" & _
            "[明細ID] COUNTER CONSTRAINT [PK_判定明細] PRIMARY KEY, " & _
            "[判定ID] LONG, [項目No] LONG, [項目名] TEXT(255), [判定方式] TEXT(10), " & _
            "[基準] TEXT(255), [サンプルNo] LONG, [入力値] TEXT(255), [合否] TEXT(10))"
        conn.Execute "CREATE INDEX [IX_判定明細_判定ID] ON [" & TABLE_DETAIL & "] ([判定ID])"
    End If

End Sub

Private Function TableExists(conn As Object, tableName As String) As Boolean
    Dim rs As Object
    Set rs = conn.OpenSchema(adSchemaTables, Array(Empty, Empty, tableName, "TABLE"))
    TableExists = Not rs.EOF
    rs.Close
End Function

Private Function NewCommand(conn As Object, sql As String) As Object
    Dim cmd As Object
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    cmd.CommandText = sql
    Set NewCommand = cmd
End Function

Private Sub AddParam(cmd As Object, paramType As Long, paramValue As Variant)
    If paramType = adVarWChar Then
        cmd.Parameters.Append cmd.CreateParameter(, paramType, adParamInput, 255, paramValue)
    Else
        cmd.Parameters.Append cmd.CreateParameter(, paramType, adParamInput, , paramValue)
    End If
End Sub

'--------------------------------------------------------------
' セルの値を255文字以内の文字列にする。空なら Null
'--------------------------------------------------------------
Private Function TextOrNull(v As Variant) As Variant
    Dim s As String
    s = CellText(v)
    If s = "" Then
        TextOrNull = Null
    Else
        TextOrNull = Left(s, 255)
    End If
End Function
