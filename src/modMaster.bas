Attribute VB_Name = "modMaster"
Option Explicit

'==============================================================
' QA試験 マスタ読み込み
'   マスタシートは1行に「品名1つ・検査1つ」を書く表形式。
'     B列 品名 / C列 検査名 / D列 数 / E列 検査のルール /
'     F列 基準値1 / G列 基準値2 / H列 単位 / I列 基準の表示 / J列 備考
'   ・検査の並び順は、マスタに書いた行の順。数が 0 の検査は行わない。
'   ・検査のルールと基準値
'       以上：基準値1 に下限
'       以下：基準値1 に上限
'       範囲：基準値1 に下限、基準値2 に上限
'       一致：基準値1 に合格の文字（「C,F」のようにカンマ区切りで複数可）
'   ・明細シートの基準値欄には「基準の表示」を出す。空欄ならルールと基準値から作る。
'   ・合否は検査者が明細シートの OK欄 で付ける（ルールは基準値欄の表示に使う）。
'==============================================================

Public Const MASTER_HEADER_ROW As Long = 3
Public Const MASTER_FIRST_ROW As Long = 4
Public Const MCOL_PRODUCT As Long = 2      ' B
Public Const MCOL_INSPECTION As Long = 3   ' C
Public Const MCOL_COUNT As Long = 4        ' D
Public Const MCOL_METHOD As Long = 5       ' E
Public Const MCOL_STD1 As Long = 6         ' F
Public Const MCOL_STD2 As Long = 7         ' G
Public Const MCOL_UNIT As Long = 8         ' H
Public Const MCOL_DISPLAY As Long = 9      ' I
Public Const MCOL_NOTE As Long = 10        ' J

Private Const MAX_CANDIDATES As Long = 10  ' 品名の候補を出す数

'--------------------------------------------------------------
' マスタの最終行
'--------------------------------------------------------------
Public Function MasterLastRow(ms As Worksheet) As Long
    MasterLastRow = ms.Cells(ms.Rows.Count, MCOL_PRODUCT).End(xlUp).Row
End Function

'--------------------------------------------------------------
' 品名をそろえる（前後の空白を除き、全角は半角、小文字は大文字にする）
'--------------------------------------------------------------
Private Function ProductKey(productName As String) As String
    ProductKey = UCase(StrConv(Trim(productName), vbNarrow))
End Function

'--------------------------------------------------------------
' マスタの行が指定の品名かどうか
'--------------------------------------------------------------
Public Function IsProductRow(ms As Worksheet, r As Long, productName As String) As Boolean
    IsProductRow = (ProductKey(CellText(ms.Cells(r, MCOL_PRODUCT).Value)) = ProductKey(productName))
End Function

'--------------------------------------------------------------
' 品名の検査名（数が1以上のもの）を、マスタに書いた順に Collection で返す
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
            If inspName <> "" And GetCountAt(ms, r) >= 1 Then
                If Not ContainsText(result, inspName) Then result.Add inspName
            End If
        End If
    Next r

    Set GetInspectionNames = result

End Function

'--------------------------------------------------------------
' 品名・検査のマスタの行。無ければ 0
'--------------------------------------------------------------
Public Function GetMasterRow(productName As String, inspName As String) As Long

    Dim ms As Worksheet
    Dim r As Long

    Set ms = ThisWorkbook.Worksheets(MASTER_SHEET)
    For r = MASTER_FIRST_ROW To MasterLastRow(ms)
        If IsProductRow(ms, r, productName) Then
            If CellText(ms.Cells(r, MCOL_INSPECTION).Value) = inspName Then
                GetMasterRow = r
                Exit Function
            End If
        End If
    Next r
    GetMasterRow = 0

End Function

'--------------------------------------------------------------
' マスタの行の数（整数でなければ 0）
'--------------------------------------------------------------
Public Function GetCountAt(ms As Worksheet, r As Long) As Long
    If IsIntegerIn(ms.Cells(r, MCOL_COUNT).Value, 0, MAX_SAMPLES) Then GetCountAt = CLng(ms.Cells(r, MCOL_COUNT).Value)
End Function

'--------------------------------------------------------------
' マスタの行の基準値欄の表示
'   「基準の表示」があればそれを使い、空欄ならルールと基準値から作る
'--------------------------------------------------------------
Public Function StandardTextAt(ms As Worksheet, r As Long) As String

    Dim std1 As String
    Dim std2 As String
    Dim unitText As String

    StandardTextAt = CellText(ms.Cells(r, MCOL_DISPLAY).Value)
    If StandardTextAt <> "" Then Exit Function

    std1 = CellText(ms.Cells(r, MCOL_STD1).Value)
    std2 = CellText(ms.Cells(r, MCOL_STD2).Value)
    unitText = CellText(ms.Cells(r, MCOL_UNIT).Value)
    StandardTextAt = FormatStandard(CellText(ms.Cells(r, MCOL_METHOD).Value), std1, std2, unitText)

End Function

'--------------------------------------------------------------
' ルールと基準値から基準値欄の文を作る（例：300℃以上で合格）
'--------------------------------------------------------------
Public Function FormatStandard(method As String, std1 As String, std2 As String, unitText As String) As String
    Select Case method
        Case METHOD_MIN:   FormatStandard = std1 & unitText & "以上で合格"
        Case METHOD_MAX:   FormatStandard = std1 & unitText & "以下で合格"
        Case METHOD_RANGE: FormatStandard = std1 & unitText & "以上" & std2 & unitText & "以下で合格"
        Case METHOD_MATCH: FormatStandard = std1 & "が合格"
        Case Else:         FormatStandard = ""
    End Select
End Function

'--------------------------------------------------------------
' 品名の行をすべてチェックする。問題があれば errText に内容を入れて False
'--------------------------------------------------------------
Public Function ValidateProductMaster(productName As String, ByRef errText As String) As Boolean

    Dim ms As Worksheet
    Dim r As Long
    Dim inspName As String
    Dim seen As New Collection

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
            ElseIf ContainsText(seen, inspName) Then
                AddError errText, ms.Cells(r, MCOL_INSPECTION), "同じ品名に検査「" & inspName & "」が2行あります"
            Else
                seen.Add inspName
            End If

            If Not IsIntegerIn(ms.Cells(r, MCOL_COUNT).Value, 0, MAX_SAMPLES) Then
                AddError errText, ms.Cells(r, MCOL_COUNT), "数が 0～" & MAX_SAMPLES & " の整数ではありません"
            ElseIf CLng(ms.Cells(r, MCOL_COUNT).Value) >= 1 Then
                Select Case CellText(ms.Cells(r, MCOL_METHOD).Value)
                    Case METHOD_MIN, METHOD_MAX, METHOD_MATCH
                        If IsBlankValue(ms.Cells(r, MCOL_STD1).Value) Then
                            AddError errText, ms.Cells(r, MCOL_STD1), "基準値1が空です"
                        End If
                    Case METHOD_RANGE
                        If IsBlankValue(ms.Cells(r, MCOL_STD1).Value) Then
                            AddError errText, ms.Cells(r, MCOL_STD1), "基準値1（下限）が空です"
                        ElseIf IsBlankValue(ms.Cells(r, MCOL_STD2).Value) Then
                            AddError errText, ms.Cells(r, MCOL_STD2), "基準値2（上限）が空です"
                        End If
                    Case Else
                        AddError errText, ms.Cells(r, MCOL_METHOD), _
                            "検査のルールは「以上」「以下」「範囲」「一致」のどれかにしてください"
                End Select
            End If
        End If
    Next r

    If errText = "" And GetInspectionNames(productName).Count = 0 Then
        errText = "　品名「" & productName & "」の検査（数が1以上）がマスタにありません" & vbCrLf
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

'--------------------------------------------------------------
' マスタに無い品名の候補（「、」区切り。最大 MAX_CANDIDATES 件）
'   入力が品名の一部になっているもの、品名が入力の一部になっているもの、
'   先頭3文字が同じものを候補にする
'--------------------------------------------------------------
Public Function FindProductCandidates(productName As String) As String

    Dim ms As Worksheet
    Dim r As Long
    Dim key As String
    Dim p As String
    Dim pk As String
    Dim found As New Collection
    Dim v As Variant
    Dim result As String

    Set ms = ThisWorkbook.Worksheets(MASTER_SHEET)
    key = ProductKey(productName)
    If key = "" Then Exit Function

    For r = MASTER_FIRST_ROW To MasterLastRow(ms)
        p = CellText(ms.Cells(r, MCOL_PRODUCT).Value)
        pk = ProductKey(p)
        If pk <> "" And Not ContainsText(found, p) Then
            If InStr(pk, key) > 0 Or InStr(key, pk) > 0 Or (Len(key) >= 3 And Left(pk, 3) = Left(key, 3)) Then
                found.Add p
                If found.Count >= MAX_CANDIDATES Then Exit For
            End If
        End If
    Next r

    For Each v In found
        If result <> "" Then result = result & "、"
        result = result & v
    Next v
    FindProductCandidates = result

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


