Attribute VB_Name = "modCommon"
Option Explicit

'==============================================================
' QA試験 共通定義
'   シート名・セル位置・色・検査のルールなどの定数と、
'   各モジュールで使う小さな関数をまとめる。
'==============================================================

' シート名
Public Const COVER_SHEET As String = "表紙"
Public Const MASTER_SHEET As String = "マスタ"
Public Const STAMP_SHEET As String = "電子印鑑"
Public Const SETTING_SHEET As String = "設定"
Public Const ITEM_LIST_SHEET As String = "検査項目一覧"  ' セットアップでマスタに移したら削除する
Public Const DETAIL_PREFIX As String = "明細_"          ' 明細シート名 = 明細_ & 検査名

' 設定シート
Public Const SETTING_DB_CELL As String = "C3"           ' BaseDBファイル（空欄ならこのブックと同じフォルダの既定ファイル）
Public Const SETTING_MAIL_FROM_CELL As String = "C6"    ' NGメールの下書きを開く Gmail アカウント（空欄ならログイン中のアカウント）
Public Const SETTING_MAIL_TO_CELL As String = "C7"      ' NGメールの宛先（複数は , 区切り）
Public Const SETTING_MAIL_CC_CELL As String = "C8"      ' NGメールの CC
Public Const SETTING_MAIL_SUBJECT_CELL As String = "C9" ' NGメールの件名（空欄なら既定の件名）
Private Const NAME_INPUT_CHANGED As String = "QA_InputChanged"   ' 保存の後に入力が変わったかの記録（ブックの非表示の名前）

' 表紙
Public Const COVER_PRODUCT_CELL As String = "C4"        ' 品名
Public Const COVER_LOT_CELL As String = "C5"            ' ロット番号
Public Const COVER_STATE_CELL As String = "C6"          ' 状態（自由入力）
Public Const COVER_DATE_CELL As String = "C7"           ' 日付
Public Const COVER_OVERALL_CELL As String = "F4"        ' 総合判定（F4:F7 結合）
Public Const COVER_APPROVER_CELL As String = "G4"       ' 承認者印欄（G4:G7 結合）
Public Const COVER_CHARGE_CELL As String = "H4"         ' 担当者印欄（H4:H7 結合）
Public Const COVER_LIST_FIRST_ROW As Long = 10          ' 検査一覧の開始行
Public Const COVER_COL_NO As Long = 2                   ' B列 No.
Public Const COVER_COL_NAME As Long = 3                 ' C列 検査名
Public Const COVER_COL_RESULT As Long = 4               ' D列 判定結果
Public Const COVER_COL_STAMP As Long = 5                ' E列 担当者印（データ印）

' 明細シート（既存の QA試験.xlsm の検査表1つ分の形）
Public Const DETAIL_TITLE_CELL As String = "B2"         ' 例：１．耐熱テスト
Public Const DETAIL_PRODUCT_CELL As String = "C3"       ' 品名（C3:D3 結合）
Public Const DETAIL_LOT_CELL As String = "C4"           ' ロット番号（C4:D4 結合）
Public Const DETAIL_HEADER_ROW As Long = 6
Public Const DETAIL_FIRST_ROW As Long = 7               ' 項目の開始行
Public Const DETAIL_COL_STD As Long = 2                 ' B列 基準値（項目の行を結合）
Public Const DETAIL_COL_NO As Long = 3                  ' C列 No.
Public Const DETAIL_COL_VALUE As Long = 4               ' D列 検査項目（測定値など）
Public Const DETAIL_COL_OK As Long = 5                  ' E列 OK欄（OK／NG）
Public Const DETAIL_STAMP_CELL As String = "G7"         ' 検査者欄（G7:G10 結合）
Public Const DETAIL_RESULT_CELL As String = "H7"        ' 総合判定（H7:H10 結合）
Public Const MAX_SAMPLES As Long = 50                   ' 1検査あたりの数の上限

' 検査のルール（マスタの「検査のルール」列に書く文字）
Public Const METHOD_MATCH As String = "一致"
Public Const METHOD_MIN As String = "以上"
Public Const METHOD_MAX As String = "以下"
Public Const METHOD_RANGE As String = "範囲"

' 判定結果（OK欄・総合判定・表紙の判定結果）
Public Const RESULT_OK As String = "OK"
Public Const RESULT_NG As String = "NG"

' 印鑑の大きさ（データ印・丸印とも同じ大きさで押す）
Public Const STAMP_SIZE As Double = 56
Public Const STAMP_MARGIN As Double = 4                 ' 印鑑欄の余白（上下左右）

' 色（BGR）
Public Const COLOR_INPUT As Long = &HFFFFFF             ' 白：入力欄
Public Const COLOR_NG As Long = &HC0FF&                 ' オレンジ（RGB(255,192,0)）：テスト結果の NG
Public Const COLOR_HEADER As Long = &HD9D9D9            ' 見出し

' テスト用（modTest から設定する。通常は使わない）
Public gTestMode As Boolean            ' True：メッセージを表示せず gTestLog に記録する
Public gTestLog As String
Public gTestAnswer As VbMsgBoxResult   ' テスト時の「はい／いいえ」の答え（0 なら「はい」）

'--------------------------------------------------------------
' メッセージを表示する（テスト中は表示せずに記録する）
'--------------------------------------------------------------
Public Function ShowMsg(prompt As String, Optional buttons As Long = 0, _
    Optional title As String = "") As Long

    If gTestMode Then
        gTestLog = gTestLog & "[" & title & "] " & prompt & vbLf
        If (buttons And 7) = vbYesNo Then
            If gTestAnswer = 0 Then
                ShowMsg = vbYes
            Else
                ShowMsg = gTestAnswer
            End If
        Else
            ShowMsg = vbOK
        End If
    Else
        ShowMsg = MsgBox(prompt, buttons, title)
    End If

End Function

'--------------------------------------------------------------
' 空白かどうか（スペースだけの場合も空白とみなす）
'--------------------------------------------------------------
Public Function IsBlankValue(cellValue As Variant) As Boolean
    If IsError(cellValue) Then
        IsBlankValue = False
    Else
        IsBlankValue = (Trim(CStr(cellValue)) = "")
    End If
End Function

'--------------------------------------------------------------
' 数値かどうか（エラー値・TRUE/FALSE は数値以外とする）
'--------------------------------------------------------------
Public Function IsNumericValue(cellValue As Variant) As Boolean
    If IsError(cellValue) Then
        IsNumericValue = False
    ElseIf VarType(cellValue) = vbBoolean Then
        IsNumericValue = False
    ElseIf IsBlankValue(cellValue) Then
        IsNumericValue = False
    Else
        IsNumericValue = IsNumeric(cellValue)
    End If
End Function

'--------------------------------------------------------------
' セルの値を文字列にする（エラー値は空文字）
'--------------------------------------------------------------
Public Function CellText(cellValue As Variant) As String
    If IsError(cellValue) Then
        CellText = ""
    Else
        CellText = Trim(CStr(cellValue))
    End If
End Function

'--------------------------------------------------------------
' OK欄の値を「OK」「NG」にそろえる（全角・小文字も受け付ける）。それ以外は空文字
'--------------------------------------------------------------
Public Function NormalizeOkNg(cellValue As Variant) As String
    Dim s As String
    s = UCase(StrConv(CellText(cellValue), vbNarrow))
    If s = RESULT_OK Or s = RESULT_NG Then NormalizeOkNg = s
End Function

'--------------------------------------------------------------
' メッセージ用にセル番地を「, 」区切りで追加する
'--------------------------------------------------------------
Public Function AppendAddress(listText As String, cell As Range) As String
    If listText = "" Then
        AppendAddress = cell.Address(False, False)
    Else
        AppendAddress = listText & ", " & cell.Address(False, False)
    End If
End Function

'--------------------------------------------------------------
' シートがあれば返す。無ければ Nothing
'--------------------------------------------------------------
Public Function FindSheet(sheetName As String) As Worksheet
    Dim i As Long
    For i = 1 To ThisWorkbook.Worksheets.Count
        If ThisWorkbook.Worksheets(i).Name = sheetName Then
            Set FindSheet = ThisWorkbook.Worksheets(i)
            Exit Function
        End If
    Next i
    Set FindSheet = Nothing
End Function

'--------------------------------------------------------------
' 指定したシートの後ろにシートを追加して返す
'--------------------------------------------------------------
Public Function AddSheetAfter(afterSheet As Worksheet) As Worksheet
    Dim afterIndex As Long
    afterIndex = afterSheet.Index
    ThisWorkbook.Worksheets.Add After:=afterSheet
    Set AddSheetAfter = ThisWorkbook.Worksheets(afterIndex + 1)
End Function

'--------------------------------------------------------------
' 明細シートかどうか
'--------------------------------------------------------------
Public Function IsDetailSheet(ws As Object) As Boolean
    IsDetailSheet = (Left(ws.Name, Len(DETAIL_PREFIX)) = DETAIL_PREFIX)
End Function

'--------------------------------------------------------------
' 明細シートの検査名（シート名から「明細_」を除いたもの）
'--------------------------------------------------------------
Public Function InspectionNameOf(ws As Worksheet) As String
    InspectionNameOf = Mid(ws.Name, Len(DETAIL_PREFIX) + 1)
End Function

'--------------------------------------------------------------
' 明細シートを並び順のまま Collection で返す
'--------------------------------------------------------------
Public Function GetDetailSheets() As Collection
    Dim result As New Collection
    Dim i As Long
    For i = 1 To ThisWorkbook.Worksheets.Count
        If IsDetailSheet(ThisWorkbook.Worksheets(i)) Then result.Add ThisWorkbook.Worksheets(i)
    Next i
    Set GetDetailSheets = result
End Function

'--------------------------------------------------------------
' 印鑑欄の大きさを、印鑑（STAMP_SIZE）＋余白が収まるまで広げる
'   高さは欄の最後の行、幅は欄の最後の列を広げる
'--------------------------------------------------------------
Public Sub EnsureStampArea(area As Range)

    Dim need As Double
    Dim lastRow As Range
    Dim lastCol As Range
    Dim guard As Long

    need = STAMP_SIZE + STAMP_MARGIN * 2
    Set lastRow = area.Rows(area.Rows.Count).EntireRow
    Set lastCol = area.Columns(area.Columns.Count).EntireColumn

    If area.Height < need Then
        lastRow.RowHeight = lastRow.RowHeight + (need - area.Height)
    End If
    Do While area.Width < need And guard < 100
        lastCol.ColumnWidth = lastCol.ColumnWidth + 0.5
        guard = guard + 1
    Loop

End Sub

'--------------------------------------------------------------
' フォームのボタンを追加する（OnAction にマクロ名を設定）
'--------------------------------------------------------------
Public Sub AddButton(ws As Worksheet, buttonName As String, caption As String, macroName As String, _
    leftPos As Double, topPos As Double, btnWidth As Double, btnHeight As Double)

    Dim btn As Object

    ' テスト（LibreOffice 等ボタンを作れない環境）ではボタンを省略する
    If gTestMode Then On Error Resume Next

    Set btn = ws.buttons.Add(leftPos, topPos, btnWidth, btnHeight)
    btn.Name = buttonName
    btn.caption = caption
    btn.OnAction = macroName

End Sub

'--------------------------------------------------------------
' 指定した名前で始まる図形を削除する。削除した数を返す
'--------------------------------------------------------------
Public Function DeleteShapesByPrefix(ws As Worksheet, prefix As String) As Long
    Dim i As Long
    Dim removed As Long
    For i = ws.Shapes.Count To 1 Step -1
        If Left(ws.Shapes(i).Name, Len(prefix)) = prefix Then
            ws.Shapes(i).Delete
            removed = removed + 1
        End If
    Next i
    DeleteShapesByPrefix = removed
End Function

'--------------------------------------------------------------
' 指定した名前の図形があるかどうか
'--------------------------------------------------------------
Public Function ShapeExists(ws As Worksheet, shapeName As String) As Boolean
    Dim i As Long
    For i = 1 To ws.Shapes.Count
        If ws.Shapes(i).Name = shapeName Then
            ShapeExists = True
            Exit Function
        End If
    Next i
    ShapeExists = False
End Function

'--------------------------------------------------------------
' 保存の後に入力が変わったかの記録（ブックの非表示の名前 QA_InputChanged に TRUE/FALSE で持つ）
'   表紙（品名・ロット・状態・日付）か明細の入力欄を変えたら TRUE、Excel 保存に成功したら FALSE にする。
'   DB 登録は、TRUE のとき「保存し直してください」で止める。
'--------------------------------------------------------------
Public Sub SetInputChanged(changed As Boolean)
    On Error Resume Next
    ThisWorkbook.Names.Add Name:=NAME_INPUT_CHANGED, RefersTo:="=" & IIf(changed, "TRUE", "FALSE"), Visible:=False
    On Error GoTo 0
End Sub

Public Function IsInputChanged() As Boolean
    Dim nm As Object
    On Error Resume Next
    Set nm = ThisWorkbook.Names(NAME_INPUT_CHANGED)
    On Error GoTo 0
    If nm Is Nothing Then Exit Function   ' 記録が無い（導入直後）は変更なしとみなす
    IsInputChanged = (UCase(nm.RefersTo) = "=TRUE")
End Function
