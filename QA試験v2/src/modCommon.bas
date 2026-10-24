Attribute VB_Name = "modCommon"
Option Explicit

'==============================================================
' QA試験 共通定義
'   シート名・セル位置・色・判定方式などの定数と、
'   各モジュールで使う小さな関数をまとめる。
'==============================================================

' シート名
Public Const COVER_SHEET As String = "表紙"
Public Const MASTER_SHEET As String = "マスタ"
Public Const STAMP_SHEET As String = "電子印鑑"
Public Const SETTING_SHEET As String = "設定"
Public Const DETAIL_PREFIX As String = "明細_"          ' 明細シート名 = 明細_ & 検査名

' 表紙
Public Const COVER_PRODUCT_CELL As String = "C4"        ' 品名
Public Const COVER_LOT_CELL As String = "C5"            ' ロット
Public Const COVER_STATE_CELL As String = "C6"          ' 状態（自由入力）
Public Const COVER_DATE_CELL As String = "C7"           ' 日付
Public Const COVER_APPROVER_CELL As String = "F4"       ' 承認者印欄（F4:F7 結合）
Public Const COVER_CHARGE_CELL As String = "G4"         ' 担当者印欄（G4:G7 結合）
Public Const COVER_LIST_FIRST_ROW As Long = 10          ' 検査一覧の開始行
Public Const COVER_COL_NO As Long = 2                   ' B列 No.
Public Const COVER_COL_NAME As Long = 3                 ' C列 検査名
Public Const COVER_COL_RESULT As Long = 4               ' D列 判定結果
Public Const COVER_COL_STAMP As Long = 5                ' E列 担当者印

' 明細シート
Public Const DETAIL_PRODUCT_CELL As String = "D3"       ' 品名（D3:E3 結合）
Public Const DETAIL_FAILCOUNT_CELL As String = "D4"     ' 不合格とする基準外の個数（D4:E4 結合）
Public Const DETAIL_HEADER_ROW As Long = 6
Public Const DETAIL_FIRST_ROW As Long = 7               ' 項目の開始行
Public Const DETAIL_COL_NO As Long = 2                  ' B列 No.
Public Const DETAIL_COL_ITEM As Long = 3                ' C列 検査項目
Public Const DETAIL_COL_METHOD As Long = 4              ' D列 判定方式
Public Const DETAIL_COL_STD As Long = 5                 ' E列 基準（表示用）
Public Const DETAIL_COL_STD1 As Long = 6                ' F列 基準値1（非表示）
Public Const DETAIL_COL_STD2 As Long = 7                ' G列 基準値2（非表示）
Public Const DETAIL_COL_N As Long = 8                   ' H列 n数（非表示）
Public Const DETAIL_COL_INPUT As Long = 9               ' I列～ 入力欄（1～n）
Public Const MAX_SAMPLES As Long = 30                   ' 1項目あたりの入力欄の上限

' 判定方式（マスタの「判定方式」列に書く文字）
Public Const METHOD_MATCH As String = "一致"
Public Const METHOD_MIN As String = "以上"
Public Const METHOD_MAX As String = "以下"
Public Const METHOD_RANGE As String = "範囲"

' 判定結果
Public Const RESULT_PASS As String = "合格"
Public Const RESULT_FAIL As String = "不合格"

' 色（BGR）
Public Const COLOR_INPUT As Long = &HFFFFFF             ' 白：入力欄
Public Const COLOR_INACTIVE As Long = &HA6A6A6          ' グレー：使わない欄
Public Const COLOR_NG As Long = &HC0FF&                 ' オレンジ（RGB(255,192,0)）：基準外
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
' フォームのボタンを追加する（OnAction にマクロ名を設定）
'--------------------------------------------------------------
Public Sub AddButton(ws As Worksheet, buttonName As String, caption As String, macroName As String, _
    leftPos As Double, topPos As Double, btnWidth As Double, btnHeight As Double)

    Dim btn As Object

    ' テスト（LibreOffice 等ボタンを作れない環境）ではボタンを省略する
    If gTestMode Then On Error Resume Next

    Set btn = ws.Buttons.Add(leftPos, topPos, btnWidth, btnHeight)
    btn.Name = buttonName
    btn.Caption = caption
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
