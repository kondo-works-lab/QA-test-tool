Attribute VB_Name = "modMaster"
Option Explicit

'==============================================================
' QA試験 マスタ読み込み
'   マスタシートは1行に1項目を書く表形式。
'     B列 品名 / C列 検査名 / D列 項目名 / E列 判定方式 /
'     F列 基準値1 / G列 基準値2 / H列 n数 / I列 不合格個数 / J列 備考
'   ・検査の並び順・項目の並び順は、マスタに書いた行の順。
'   ・不合格個数は検査ごとの値。同じ品名・検査の行のどれか1行以上に書く
'     （複数行に書く場合は同じ値にする）。
'   ・判定方式と基準値
'       一致：基準値1 に合格の文字（「A,B」のようにカンマ区切りで複数可）
'       以上：基準値1 に下限
'       以下：基準値1 に上限
'       範囲：基準値1 に下限、基準値2 に上限（境界値ちょうどは合格）
'==============================================================

Public Const MASTER_HEADER_ROW As Long = 3
Public Const MASTER_FIRST_ROW As Long = 4
Public Const MCOL_PRODUCT As Long = 2      ' B
Public Const MCOL_INSPECTION As Long = 3   ' C
Public Const MCOL_ITEM As Long = 4         ' D
Public Const MCOL_METHOD As Long = 5       ' E
Public Const MCOL_STD1 As Long = 6         ' F
Public Const MCOL_STD2 As Long = 7         ' G
Public Const MCOL_N As Long = 8            ' H
Public Const MCOL_FAILCOUNT As Long = 9    ' I
Public Const MCOL_NOTE As Long = 10        ' J

'--------------------------------------------------------------
' マスタの最終行
'--------------------------------------------------------------
Public Function MasterLastRow(ms As Worksheet) As Long
    MasterLastRow = ms.Cells(ms.Rows.Count, MCOL_PRODUCT).End(xlUp).Row
End Function

'--------------------------------------------------------------
' マスタの行が指定の品名かどうか（前後の空白を除き、大文字・小文字は区別しない）
'--------------------------------------------------------------
Public Function IsProductRow(ms As Worksheet, r As Long, productName As String) As Boolean
    IsProductRow = (StrComp(CellText(ms.Cells(r, MCOL_PRODUCT).Value), productName, vbTextCompare) = 0)
End Function

'--------------------------------------------------------------
' 品名の検査名を、マスタに出てくる順に Collection で返す（重複なし）
'--------------------------------------------------------------
Public Function GetInspectionNames(productName As String) As Collection

    Dim ms As Worksheet
    Dim result As New Collection
    Dim r As Long
    Dim inspName As String

    Set ms = ThisWorkbook.Worksheets(MASTER_SHEET)
    For r = MASTER_FIRST_ROW To MasterLastRow(ms)
        If IsProductRow(ms, r, productName) Then
            inspName = CellText(ms.Cells(r, MCOL_INSPECTION).Value)
            If inspName <> "" And Not ContainsText(result, inspName) Then result.Add inspName
        End If
    Next r

    Set GetInspectionNames = result

End Function

'--------------------------------------------------------------
' 品名・検査の行番号を、マスタに書いた順に Collection で返す
'--------------------------------------------------------------
Public Function GetItemRows(productName As String, inspName As String) As Collection

    Dim ms As Worksheet
    Dim result As New Collection
    Dim r As Long

    Set ms = ThisWorkbook.Worksheets(MASTER_SHEET)
    For r = MASTER_FIRST_ROW To MasterLastRow(ms)
        If IsProductRow(ms, r, productName) Then
            If CellText(ms.Cells(r, MCOL_INSPECTION).Value) = inspName Then result.Add r
        End If
    Next r

    Set GetItemRows = result

End Function

'--------------------------------------------------------------
' 品名・検査の不合格個数（検査内で入力されている値）。無ければ 0
'--------------------------------------------------------------
Public Function GetFailCount(productName As String, inspName As String) As Long

    Dim ms As Worksheet
    Dim r As Variant
    Dim v As Variant

    Set ms = ThisWorkbook.Worksheets(MASTER_SHEET)
    For Each r In GetItemRows(productName, inspName)
        v = ms.Cells(r, MCOL_FAILCOUNT).Value
        If Not IsBlankValue(v) Then
            GetFailCount = CLng(v)
            Exit Function
        End If
    Next r
    GetFailCount = 0

End Function

'--------------------------------------------------------------
' 品名の行をすべてチェックする。問題があれば errText に内容を入れて False
'--------------------------------------------------------------
Public Function ValidateProductMaster(productName As String, ByRef errText As String) As Boolean

    Dim ms As Worksheet
    Dim r As Long
    Dim inspName As String
    Dim method As String
    Dim v As Variant
    Dim failText As String
    Dim inspNames As Collection
    Dim n As Variant
    Dim firstFail As Variant
    Dim rr As Variant
    Dim found As Boolean

    Set ms = ThisWorkbook.Worksheets(MASTER_SHEET)
    errText = ""

    For r = MASTER_FIRST_ROW To MasterLastRow(ms)
        If IsProductRow(ms, r, productName) Then

            inspName = CellText(ms.Cells(r, MCOL_INSPECTION).Value)
            If inspName = "" Then
                AddError errText, ms.Cells(r, MCOL_INSPECTION), "検査名が空です"
            ElseIf Not IsValidSheetPart(inspName) Then
                AddError errText, ms.Cells(r, MCOL_INSPECTION), _
                    "検査名に使えない文字（\ / : * ? [ ]）があるか、" & (31 - Len(DETAIL_PREFIX)) & "文字を超えています"
            End If

            If CellText(ms.Cells(r, MCOL_ITEM).Value) = "" Then
                AddError errText, ms.Cells(r, MCOL_ITEM), "項目名が空です"
            End If

            method = CellText(ms.Cells(r, MCOL_METHOD).Value)
            Select Case method
                Case METHOD_MATCH
                    If IsBlankValue(ms.Cells(r, MCOL_STD1).Value) Then
                        AddError errText, ms.Cells(r, MCOL_STD1), "基準値1（合格の文字）が空です"
                    End If
                Case METHOD_MIN, METHOD_MAX
                    If Not IsNumericValue(ms.Cells(r, MCOL_STD1).Value) Then
                        AddError errText, ms.Cells(r, MCOL_STD1), "基準値1が数値ではありません"
                    End If
                Case METHOD_RANGE
                    If Not IsNumericValue(ms.Cells(r, MCOL_STD1).Value) Then
                        AddError errText, ms.Cells(r, MCOL_STD1), "基準値1（下限）が数値ではありません"
                    ElseIf Not IsNumericValue(ms.Cells(r, MCOL_STD2).Value) Then
                        AddError errText, ms.Cells(r, MCOL_STD2), "基準値2（上限）が数値ではありません"
                    ElseIf CDbl(ms.Cells(r, MCOL_STD1).Value) > CDbl(ms.Cells(r, MCOL_STD2).Value) Then
                        AddError errText, ms.Cells(r, MCOL_STD1), "下限が上限より大きくなっています"
                    End If
                Case Else
                    AddError errText, ms.Cells(r, MCOL_METHOD), _
                        "判定方式は「一致」「以上」「以下」「範囲」のどれかにしてください"
            End Select

            If Not IsIntegerIn(ms.Cells(r, MCOL_N).Value, 1, MAX_SAMPLES) Then
                AddError errText, ms.Cells(r, MCOL_N), "n数が 1～" & MAX_SAMPLES & " の整数ではありません"
            End If

            v = ms.Cells(r, MCOL_FAILCOUNT).Value
            If Not IsBlankValue(v) Then
                If Not IsIntegerIn(v, 1, 9999) Then
                    AddError errText, ms.Cells(r, MCOL_FAILCOUNT), "不合格個数が 1 以上の整数ではありません"
                End If
            End If
        End If
    Next r

    ' 不合格個数：検査ごとに1つ以上入力があり、値がそろっているか
    Set inspNames = GetInspectionNames(productName)
    For Each n In inspNames
        found = False
        firstFail = Empty
        For Each rr In GetItemRows(productName, CStr(n))
            v = ms.Cells(rr, MCOL_FAILCOUNT).Value
            If Not IsBlankValue(v) And IsIntegerIn(v, 1, 9999) Then
                If Not found Then
                    found = True
                    firstFail = CLng(v)
                ElseIf CLng(v) <> firstFail Then
                    AddError errText, ms.Cells(rr, MCOL_FAILCOUNT), _
                        "検査「" & n & "」の不合格個数が行によって違います"
                End If
            End If
        Next rr
        If Not found Then
            failText = failText & "　検査「" & n & "」の不合格個数（" & ms.Cells(MASTER_HEADER_ROW, MCOL_FAILCOUNT).Address(False, False) & _
                       "列）が入力されていません" & vbCrLf
        End If
    Next n
    errText = errText & failText

    If inspNames.Count = 0 And errText = "" Then
        errText = "　品名「" & productName & "」の検査がマスタにありません" & vbCrLf
    End If

    ValidateProductMaster = (errText = "")

End Function

'--------------------------------------------------------------
' 品名がマスタにあるかどうか
'--------------------------------------------------------------
Public Function ProductExists(productName As String) As Boolean
    Dim ms As Worksheet
    Dim r As Long
    Set ms = ThisWorkbook.Worksheets(MASTER_SHEET)
    For r = MASTER_FIRST_ROW To MasterLastRow(ms)
        If IsProductRow(ms, r, productName) Then
            ProductExists = True
            Exit Function
        End If
    Next r
    ProductExists = False
End Function

Private Sub AddError(ByRef errText As String, cell As Range, message As String)
    errText = errText & "　" & cell.Address(False, False) & "：" & message & vbCrLf
End Sub

'--------------------------------------------------------------
' 値が min～max の整数かどうか
'--------------------------------------------------------------
Public Function IsIntegerIn(v As Variant, minValue As Long, maxValue As Long) As Boolean
    IsIntegerIn = False
    If Not IsNumericValue(v) Then Exit Function
    If CDbl(v) <> Int(CDbl(v)) Then Exit Function
    If CDbl(v) < minValue Or CDbl(v) > maxValue Then Exit Function
    IsIntegerIn = True
End Function

'--------------------------------------------------------------
' シート名に使えるかどうか（「明細_」を付けて31文字以内、禁止文字なし）
'--------------------------------------------------------------
Private Function IsValidSheetPart(text As String) As Boolean
    Dim ch As Variant
    IsValidSheetPart = False
    If Len(DETAIL_PREFIX & text) > 31 Then Exit Function
    For Each ch In Array("\", "/", ":", "*", "?", "[", "]")
        If InStr(text, ch) > 0 Then Exit Function
    Next ch
    IsValidSheetPart = True
End Function

Private Function ContainsText(col As Collection, text As String) As Boolean
    Dim v As Variant
    For Each v In col
        If v = text Then
            ContainsText = True
            Exit Function
        End If
    Next v
    ContainsText = False
End Function
