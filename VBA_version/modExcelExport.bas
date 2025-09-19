Attribute VB_Name = "modExcelExport"
Option Explicit

'=============================================================
' Excel 出力系ユーティリティ
' - タスクのフラット化結果を "tasks" シートへ書き込み
' - "_original" シートへバックアップ、差分ハイライトを設定
' - "project" シートへの一覧貼り付け
'=============================================================

Private Const TASKS_SHEET As String = "tasks"
Private Const ORIGINAL_SHEET As String = "_original"
Private Const PROJECT_SHEET As String = "project"

Public Sub PopulateWorkbookWithTasks(ByVal targetWorkbook As Workbook, _
                                     ByVal flatRows As Collection, _
                                     ByVal headers As Collection, _
                                     Optional projectRows As Variant)
    Dim wsTasks As Worksheet
    Set wsTasks = EnsureWorksheet(targetWorkbook, TASKS_SHEET)

    Dim wsOriginal As Worksheet
    Set wsOriginal = EnsureWorksheet(targetWorkbook, ORIGINAL_SHEET)

    Dim wsProject As Worksheet
    Set wsProject = EnsureWorksheet(targetWorkbook, PROJECT_SHEET)

    ClearWorksheet wsTasks
    ClearWorksheet wsOriginal

    WriteHeaderAndRows wsTasks, headers, flatRows
    WriteHeaderAndRows wsOriginal, headers, flatRows
    wsOriginal.Visible = xlSheetHidden

    ApplyDifferenceHighlight wsTasks, wsOriginal, headers.Count, flatRows.Count
    UnlockDataCells wsTasks, headers.Count, flatRows.Count
    ProtectTasksSheet wsTasks

    If Not IsMissing(projectRows) Then
        ClearWorksheet wsProject
        WriteProjectRows wsProject, projectRows
    End If
End Sub

Private Function EnsureWorksheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set EnsureWorksheet = wb.Worksheets(sheetName)
    On Error GoTo 0

    If EnsureWorksheet Is Nothing Then
        Set EnsureWorksheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        EnsureWorksheet.Name = sheetName
    End If
End Function

Private Sub ClearWorksheet(ByVal ws As Worksheet)
    With ws
        .Cells.Clear
        .Cells.Validation.Delete
        .Cells.FormatConditions.Delete
        .Protection.AllowEditRanges.Clear
    End With
End Sub

Private Sub WriteHeaderAndRows(ByVal ws As Worksheet, _
                               ByVal headers As Collection, _
                               ByVal rows As Collection)
    Dim colIndex As Long
    Dim header As Variant
    colIndex = 1
    For Each header In headers
        ws.Cells(1, colIndex).Value = header
        colIndex = colIndex + 1
    Next header

    Dim rowIndex As Long
    rowIndex = 2

    Dim rowDict As Scripting.Dictionary
    Dim value As Variant

    For Each rowDict In rows
        colIndex = 1
        For Each header In headers
            If rowDict.Exists(header) Then
                value = SanitizeCellValue(rowDict(header))
                ws.Cells(rowIndex, colIndex).Value = value
            Else
                ws.Cells(rowIndex, colIndex).Value = vbNullString
            End If
            colIndex = colIndex + 1
        Next header
        rowIndex = rowIndex + 1
    Next rowDict
End Sub

Private Function SanitizeCellValue(ByVal value As Variant) As Variant
    If VarType(value) = vbString Then
        SanitizeCellValue = RemoveIllegalCharacters(CStr(value))
    Else
        SanitizeCellValue = value
    End If
End Function

Private Function RemoveIllegalCharacters(ByVal text As String) As String
    Dim i As Long
    Dim code As Integer
    Dim builder As String
    For i = 1 To Len(text)
        code = Asc(Mid$(text, i, 1))
        If code = 9 Or code = 10 Or code = 13 Or code >= 32 Then
            builder = builder & Mid$(text, i, 1)
        End If
    Next i
    RemoveIllegalCharacters = builder
End Function

Private Sub ApplyDifferenceHighlight(ByVal wsTasks As Worksheet, _
                                     ByVal wsOriginal As Worksheet, _
                                     ByVal headerCount As Long, _
                                     ByVal rowCount As Long)
    If headerCount = 0 Then Exit Sub

    Dim lastColLetter As String
    lastColLetter = Split(wsTasks.Cells(1, headerCount).Address(False, False), "1")(0)

    Dim lastRow As Long
    lastRow = rowCount + 1 ' ヘッダー行を含める

    Dim rangeAddress As String
    rangeAddress = "$A$1:$" & lastColLetter & "$" & lastRow

    wsTasks.FormatConditions.Delete

    Dim formula As String
    formula = "=A1<>" & ORIGINAL_SHEET & "!A1"

    Dim fc As FormatCondition
    Set fc = wsTasks.Range(rangeAddress).FormatConditions.Add(Type:=xlExpression, Formula1:=formula)
    fc.Interior.Color = RGB(255, 255, 0)
End Sub

Private Sub UnlockDataCells(ByVal ws As Worksheet, _
                            ByVal headerCount As Long, _
                            ByVal rowCount As Long)
    Dim rng As Range
    If headerCount = 0 Then Exit Sub
    If rowCount < 0 Then rowCount = 0

    Dim lastRow As Long
    lastRow = rowCount + 1
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, headerCount))
    rng.Locked = False
End Sub

Private Sub ProtectTasksSheet(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, _
               AllowFormattingColumns:=False, AllowFormattingRows:=False, _
               AllowInsertingColumns:=False, AllowInsertingRows:=False, _
               AllowDeletingColumns:=False, AllowDeletingRows:=False, _
               AllowSorting:=False, AllowFiltering:=False, AllowUsingPivotTables:=False
    On Error GoTo 0
End Sub

Private Sub WriteProjectRows(ByVal ws As Worksheet, ByVal projectRows As Variant)
    If IsEmpty(projectRows) Then Exit Sub

    Dim rowIndex As Long
    Dim colIndex As Long

    If IsArray(projectRows) Then
        Dim lowerRow As Long
        Dim upperRow As Long
        Dim lowerCol As Long
        Dim upperCol As Long
        lowerRow = LBound(projectRows, 1)
        upperRow = UBound(projectRows, 1)
        lowerCol = LBound(projectRows, 2)
        upperCol = UBound(projectRows, 2)

        For rowIndex = lowerRow To upperRow
            For colIndex = lowerCol To upperCol
                ws.Cells(rowIndex - lowerRow + 1, colIndex - lowerCol + 1).Value = projectRows(rowIndex, colIndex)
            Next colIndex
        Next rowIndex
    ElseIf TypeName(projectRows) = "Collection" Then
        rowIndex = 1
        Dim rowItem As Variant
        For Each rowItem In projectRows
            colIndex = 1
            Dim cellValue As Variant
            For Each cellValue In rowItem
                ws.Cells(rowIndex, colIndex).Value = cellValue
                colIndex = colIndex + 1
            Next cellValue
            rowIndex = rowIndex + 1
        Next rowItem
    End If
End Sub

Public Function BuildFlatRows(ByVal tasks As Variant) As Collection
    Dim output As Collection
    Set output = New Collection

    If IsObject(tasks) Then
        Dim task As Variant
        For Each task In tasks
            Dim dict As Scripting.Dictionary
            Set dict = FlattenDictionary(task)
            output.Add dict
        Next task
    ElseIf IsArray(tasks) Then
        Dim i As Long
        For i = LBound(tasks) To UBound(tasks)
            Dim dict As Scripting.Dictionary
            Set dict = FlattenDictionary(tasks(i))
            output.Add dict
        Next i
    End If

    Set BuildFlatRows = output
End Function
