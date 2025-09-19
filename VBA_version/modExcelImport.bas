Attribute VB_Name = "modExcelImport"
Option Explicit

'=============================================================
' Excel -> API 取り込み処理
' - tasks シートの変更セルを検出（黄色塗りつぶし or _original 比較）
' - 行情報をディクショナリとして返却
' - 塗りつぶし更新、結果保存などのユーティリティ
'=============================================================

Public Const COLOR_SUCCESS As Long = RGB(0, 176, 240) ' 青
Public Const COLOR_FAILURE As Long = RGB(255, 0, 0) ' 赤
Public Const COLOR_HIGHLIGHT As Long = RGB(255, 255, 0) ' 黄

Public Function CollectTaskRows(ByVal wsTasks As Worksheet, _
                                Optional ByVal wsOriginal As Worksheet) As Collection
    Dim results As Collection
    Set results = New Collection

    Dim headers() As String
    Dim headerCount As Long

    Dim rng As Range
    Dim rowRange As Range

    Set rng = wsTasks.UsedRange
    If rng Is Nothing Then GoTo ExitProc

    For Each rowRange In rng.Rows
        Dim currentRow As Long
        currentRow = rowRange.Row

        If currentRow = 1 Then
            headerCount = rowRange.Columns.Count
            ReDim headers(1 To headerCount)

            Dim headerCell As Range
            Dim idx As Long
            idx = 1
            For Each headerCell In rowRange.Columns
                headers(idx) = Trim$(CStr(headerCell.Value))
                idx = idx + 1
            Next headerCell
            GoTo ContinueLoop
        End If

        If headerCount = 0 Then Exit For

        Dim values As Scripting.Dictionary
        Set values = New Scripting.Dictionary
        values.CompareMode = TextCompare

        Dim yellowCells As Collection
        Set yellowCells = New Collection

        Dim hasValue As Boolean
        Dim colIndex As Long
        Dim headerName As String
        Dim cell As Range

        For colIndex = 1 To headerCount
            headerName = headers(colIndex)
            If Len(headerName) = 0 Then GoTo NextColumn

            Set cell = wsTasks.Cells(currentRow, colIndex)
            Dim cellValue As Variant
            cellValue = cell.Value
            If Not IsEmpty(cellValue) And cellValue <> "" Then
                hasValue = True
            End If

            Dim normalized As Variant
            normalized = NormalizeCellValue(cellValue)
            values(headerName) = normalized

            If IsYellowCell(cell) Then
                yellowCells.Add Array(cell.Row, cell.Column)
            End If

            If Not wsOriginal Is Nothing Then
                Dim originalValue As Variant
                originalValue = NormalizeCellValue(wsOriginal.Cells(currentRow, colIndex).Value)
                If Not ValuesEqual(normalized, originalValue) Then
                    AddUniqueCoordinate yellowCells, cell.Row, cell.Column
                End If
            End If
NextColumn:
        Next colIndex

        If Not hasValue Then GoTo ContinueLoop
        If yellowCells.Count = 0 Then GoTo ContinueLoop

        Dim taskRow As Scripting.Dictionary
        Set taskRow = New Scripting.Dictionary
        taskRow.CompareMode = TextCompare
        taskRow("RowIndex") = currentRow
        taskRow("Values") = values
        taskRow("YellowCells") = yellowCells
        taskRow("TaskId") = GetFirstNonEmpty(values, Array("Id", "ID", "TaskId"))
        taskRow("ProjectId") = GetFirstNonEmpty(values, Array("ProjectId", "project_id"))

        results.Add taskRow
ContinueLoop:
    Next rowRange

ExitProc:
    Set CollectTaskRows = results
End Function

Private Function GetFirstNonEmpty(ByVal dict As Scripting.Dictionary, ByVal keys As Variant) As String
    Dim key As Variant
    For Each key In keys
        If dict.Exists(key) Then
            If Len(Trim$(CStr(dict(key)))) > 0 Then
                GetFirstNonEmpty = Trim$(CStr(dict(key)))
                Exit Function
            End If
        End If
    Next key
    GetFirstNonEmpty = vbNullString
End Function

Private Sub AddUniqueCoordinate(ByVal coords As Collection, ByVal rowIndex As Long, ByVal colIndex As Long)
    Dim item As Variant
    For Each item In coords
        If item(0) = rowIndex And item(1) = colIndex Then
            Exit Sub
        End If
    Next item
    coords.Add Array(rowIndex, colIndex)
End Sub

Private Function ValuesEqual(ByVal a As Variant, ByVal b As Variant) As Boolean
    If IsObject(a) Or IsObject(b) Then
        ValuesEqual = False
        Exit Function
    End If
    If IsEmpty(a) And IsEmpty(b) Then
        ValuesEqual = True
        Exit Function
    End If
    If IsNull(a) And IsNull(b) Then
        ValuesEqual = True
        Exit Function
    End If
    ValuesEqual = (a = b)
End Function

Public Function NormalizeCellValue(ByVal value As Variant) As Variant
    If IsError(value) Then
        NormalizeCellValue = Empty
        Exit Function
    End If

    If IsMissing(value) Then
        NormalizeCellValue = Empty
        Exit Function
    End If

    If IsNull(value) Then
        NormalizeCellValue = Empty
        Exit Function
    End If

    Select Case VarType(value)
        Case vbEmpty
            NormalizeCellValue = Empty
        Case vbString
            Dim trimmed As String
            trimmed = Trim$(CStr(value))
            If Len(trimmed) = 0 Then
                NormalizeCellValue = Empty
            Else
                Dim boolCandidate As Variant
                boolCandidate = NormalizeBooleanString(trimmed)
                If VarType(boolCandidate) = vbBoolean Then
                    NormalizeCellValue = boolCandidate
                    Exit Function
                End If

                Dim parsed As Variant
                parsed = TryParseJson(trimmed)
                If Not IsEmpty(parsed) Then
                    NormalizeCellValue = parsed
                Else
                    NormalizeCellValue = trimmed
                End If
            End If
        Case vbDate
            NormalizeCellValue = Format$(value, "yyyy-mm-dd\Thh:nn:ss")
        Case vbBoolean
            NormalizeCellValue = value
        Case vbDouble, vbSingle, vbCurrency
            If IsNumeric(value) Then
                If CLng(value) = value Then
                    NormalizeCellValue = CLng(value)
                Else
                    NormalizeCellValue = CDbl(value)
                End If
            Else
                NormalizeCellValue = value
            End If
        Case vbInteger, vbLong, vbByte, vbLongLong
            NormalizeCellValue = value
        Case Else
            NormalizeCellValue = value
    End Select
End Function

Private Function TryParseJson(ByVal text As String) As Variant
    On Error GoTo Failed
    If Left$(text, 1) = "{" Or Left$(text, 1) = "[" Then
        TryParseJson = JsonConverter.ParseJson(text)
        Exit Function
    End If
Failed:
    TryParseJson = Empty
End Function

Private Function IsYellowCell(ByVal cell As Range) As Boolean
    On Error GoTo NotYellow
    If cell.Interior.Pattern <> xlSolid Then GoTo NotYellow

    Dim color As Long
    color = cell.Interior.Color

    If color = RGB(255, 255, 0) Then
        IsYellowCell = True
        Exit Function
    End If

    Dim index As Long
    index = cell.Interior.ColorIndex
    If index = 6 Or index = 44 Or index = 36 Or index = 27 Then
        IsYellowCell = True
        Exit Function
    End If
NotYellow:
End Function

Public Sub SetFillColor(ByVal ws As Worksheet, ByVal coords As Collection, ByVal color As Long)
    Dim item As Variant
    For Each item In coords
        ws.Cells(item(0), item(1)).Interior.Color = color
    Next item
End Sub

Public Function SaveWithSuffix(ByVal wb As Workbook, ByVal originalPath As String, _
                               Optional ByVal suffix As String = "_result") As String
    Dim folderPath As String
    Dim fileName As String
    folderPath = Left$(originalPath, InStrRev(originalPath, "\"))
    fileName = Mid$(originalPath, InStrRev(originalPath, "\") + 1)

    Dim dotPos As Long
    dotPos = InStrRev(fileName, ".")
    Dim stem As String
    Dim ext As String
    If dotPos > 0 Then
        stem = Left$(fileName, dotPos - 1)
        ext = Mid$(fileName, dotPos)
    Else
        stem = fileName
        ext = ".xlsm"
    End If

    Dim newName As String
    newName = stem & suffix & ext
    Dim outputPath As String
    outputPath = folderPath & newName

    wb.SaveCopyAs outputPath
    SaveWithSuffix = outputPath
End Function

Public Function BuildUpdatePayload(ByVal flatValues As Scripting.Dictionary) As Scripting.Dictionary
    Set BuildUpdatePayload = UnflattenDictionary(flatValues)
End Function
