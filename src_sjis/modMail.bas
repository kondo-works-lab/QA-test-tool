Attribute VB_Name = "modMail"
Option Explicit

'==============================================================
' QA試験 NG メールの下書き（Gmail）
'   表紙の総合判定が NG のとき、「保存」ボタンで Excel・PDF を保存した後に呼ばれ、
'   ブラウザで Gmail の新規メール作成画面を開く（宛先・CC・件名・本文を入れた状態）。送信はしない。
'   ・PDF はこの方法では自動で添付できないため、PDF を選んだ状態で保存先フォルダを開く。
'     作成画面へドラッグして添付し、内容を確認して送信してもらう。
'   ・使う Gmail アカウント・宛先・CC・件名は設定シートに書く（宛先が空欄なら下書きを作らない）。
'     Gmail アカウントが空欄なら、ブラウザでログイン中のアカウントで開く。
'   ・件名が空欄なら「【NG】QA試験 品名 ロット番号」にする。
'   ・本文：品名・ロット番号・状態・日付・NG 回数、NG の検査（検査名・検査員・NG の No. と検査値）、保存したファイル名。
'   ・パスワードは使わない。
'==============================================================

Private Const GMAIL_COMPOSE_URL As String = "https://mail.google.com/mail/?view=cm&fs=1"

Public gTestMailBody As String   ' テスト用：テスト中はブラウザを開かず、内容をここに残す

'--------------------------------------------------------------
' NG メールの下書き（Gmail の作成画面）を開く。開けたら True。
'   mailTo に宛先、失敗したら errMsg に理由を入れる
'   excelPath / pdfPath：保存したファイル（保存できなかった形式は空文字）
'   ngCount：何回目の NG の保存か
'--------------------------------------------------------------
Public Function CreateNgMailDraft(excelPath As String, pdfPath As String, ngCount As Long, _
    ByRef mailTo As String, ByRef errMsg As String) As Boolean

    Dim ss As Worksheet
    Dim cs As Worksheet
    Dim account As String
    Dim mailCc As String
    Dim subject As String
    Dim body As String
    Dim url As String
    Dim sh As Object

    errMsg = ""
    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)
    Set ss = FindSheet(SETTING_SHEET)
    If ss Is Nothing Then
        errMsg = "設定シートがありません。"
        Exit Function
    End If

    account = CellText(ss.Range(SETTING_MAIL_FROM_CELL).Value)
    mailTo = CellText(ss.Range(SETTING_MAIL_TO_CELL).Value)
    mailCc = CellText(ss.Range(SETTING_MAIL_CC_CELL).Value)
    subject = CellText(ss.Range(SETTING_MAIL_SUBJECT_CELL).Value)
    If mailTo = "" Then
        errMsg = "宛先が設定されていません（設定シートの " & SETTING_MAIL_TO_CELL & "）。"
        Exit Function
    End If
    If subject = "" Then
        subject = "【NG】QA試験 " & CellText(cs.Range(COVER_PRODUCT_CELL).Value) & " " & _
                  CellText(cs.Range(COVER_LOT_CELL).Value)
    End If
    body = BuildNgMailBody(excelPath, pdfPath, ngCount)

    On Error GoTo ErrHandler
    url = GMAIL_COMPOSE_URL & "&to=" & UrlEncode(mailTo)
    If mailCc <> "" Then url = url & "&cc=" & UrlEncode(mailCc)
    url = url & "&su=" & UrlEncode(subject) & "&body=" & UrlEncode(body)
    If account <> "" Then url = url & "&authuser=" & UrlEncode(account)

    If gTestMode Then
        gTestMailBody = "URL：" & url & vbLf & "添付するファイル：" & pdfPath & vbLf & body
        gTestLog = gTestLog & "[NGメール下書き] " & mailTo & " / " & subject & vbLf
        CreateNgMailDraft = True
        Exit Function
    End If

    ' ShellExecute で既定のブラウザに渡す（FollowHyperlink は長い URL で失敗することがあるため使わない）
    Set sh = CreateObject("Shell.Application")
    sh.ShellExecute url
    ' 添付用に、PDF を選んだ状態で保存先フォルダを開く
    If pdfPath <> "" Then sh.ShellExecute "explorer.exe", "/select,""" & pdfPath & """"

    CreateNgMailDraft = True
    Exit Function

ErrHandler:
    errMsg = "Gmail の作成画面を開けませんでした。" & vbCrLf & Err.Description

End Function

'--------------------------------------------------------------
' URL 用に UTF-8 でエンコードする（Excel の ENCODEURL 関数を使う）
'--------------------------------------------------------------
Private Function UrlEncode(text As String) As String
    UrlEncode = Application.WorksheetFunction.EncodeURL(text)
End Function

'--------------------------------------------------------------
' NG メールの本文
'--------------------------------------------------------------
Public Function BuildNgMailBody(excelPath As String, pdfPath As String, ngCount As Long) As String

    Dim cs As Worksheet
    Dim item As Variant
    Dim ws As Worksheet
    Dim r As Long
    Dim s As String
    Dim ngItems As String
    Dim dateValue As Variant

    Set cs = ThisWorkbook.Worksheets(COVER_SHEET)

    s = "QA試験で NG の検査がありました。" & vbCrLf & vbCrLf
    s = s & "品名：" & CellText(cs.Range(COVER_PRODUCT_CELL).Value) & vbCrLf
    s = s & "ロット番号：" & CellText(cs.Range(COVER_LOT_CELL).Value) & vbCrLf
    s = s & "状態：" & CellText(cs.Range(COVER_STATE_CELL).Value) & vbCrLf
    dateValue = cs.Range(COVER_DATE_CELL).Value
    If Not IsError(dateValue) Then
        If IsDate(dateValue) Then
            s = s & "日付：" & Format(CDate(dateValue), "yyyy/mm/dd") & vbCrLf
        Else
            s = s & "日付：" & CellText(dateValue) & vbCrLf
        End If
    End If

    s = s & "NG 回数：" & ngCount & "回目" & vbCrLf

    s = s & vbCrLf & "■ NG の検査" & vbCrLf
    For Each item In GetDetailSheets()
        Set ws = item
        If GetDetailResult(ws) = RESULT_NG Then
            ngItems = ""
            For r = DETAIL_FIRST_ROW To DETAIL_FIRST_ROW + DetailItemCount(ws) - 1
                If NormalizeOkNg(ws.Cells(r, DETAIL_COL_OK).Value) = RESULT_NG Then
                    If ngItems <> "" Then ngItems = ngItems & "、"
                    ngItems = ngItems & "No." & CellText(ws.Cells(r, DETAIL_COL_NO).Value) & "：" & _
                              CellText(ws.Cells(r, DETAIL_COL_VALUE).Value)
                End If
            Next r
            s = s & "・" & InspectionNameOf(ws) & "（検査員：" & GetStampName(ws, DETAIL_STAMP_NAME) & "）" & vbCrLf
            s = s & "　" & ngItems & vbCrLf
        End If
    Next item

    ' URL が長くなりすぎないよう、ファイルはフォルダを除いた名前だけ書く
    s = s & vbCrLf & "■ 保存したファイル" & vbCrLf
    If excelPath <> "" Then s = s & FileNameOf(excelPath) & vbCrLf
    If pdfPath <> "" Then s = s & FileNameOf(pdfPath) & "（添付）" & vbCrLf

    BuildNgMailBody = s

End Function

Private Function FileNameOf(path As String) As String
    FileNameOf = Mid(path, InStrRev(path, "\") + 1)
End Function
