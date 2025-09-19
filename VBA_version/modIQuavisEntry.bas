Attribute VB_Name = "modIQuavisEntry"
Option Explicit

'=============================================================
' iQUAVIS VBA エントリーポイント例
' - フォームやリボンボタンから利用しやすいラッパー関数
'=============================================================

Public Function CreateClient(Optional ByVal baseUrl As String = vbNullString, _
                             Optional ByVal debugLogging As Boolean = False) As IQuavisClient
    Dim client As IQuavisClient
    Set client = New IQuavisClient
    If Len(Trim$(baseUrl)) > 0 Then
        client.BaseUrl = baseUrl
    End If
    client.DebugLogging = debugLogging
    Set CreateClient = client
End Function

Public Sub AuthenticateClient(ByVal client As IQuavisClient, _
                              ByVal userId As String, _
                              ByVal password As String)
    If client Is Nothing Then Err.Raise vbObjectError + 2000, "AuthenticateClient", "client が指定されていません。"
    client.Login userId, password
End Sub

Public Function FetchProjects(ByVal client As IQuavisClient, _
                              Optional ByVal nameFilter As String = vbNullString) As Variant
    FetchProjects = client.ListProjects(nameFilter)
End Function

Public Function FlattenProjects(ByVal projects As Variant) As Collection
    Dim rows As Collection
    Set rows = New Collection

    If IsObject(projects) Then
        Dim project As Variant
        For Each project In projects
            Dim flat As Scripting.Dictionary
            Set flat = FlattenDictionary(project)
            rows.Add flat
        Next project
    End If

    Set FlattenProjects = rows
End Function

Public Function BuildProjectRows(ByVal flatProjects As Collection) As Variant
    Dim headers As Collection
    Set headers = CollectHeaders(flatProjects)

    Dim rowCount As Long
    rowCount = flatProjects.Count + 1
    Dim colCount As Long
    colCount = headers.Count

    If rowCount = 1 Or colCount = 0 Then
        BuildProjectRows = Empty
        Exit Function
    End If

    Dim data() As Variant
    ReDim data(1 To rowCount, 1 To colCount)

    Dim colIndex As Long
    Dim header As Variant

    colIndex = 1
    For Each header In headers
        data(1, colIndex) = header
        colIndex = colIndex + 1
    Next header

    Dim rowIndex As Long
    rowIndex = 2

    Dim dict As Scripting.Dictionary
    For Each dict In flatProjects
        colIndex = 1
        For Each header In headers
            If dict.Exists(header) Then
                data(rowIndex, colIndex) = dict(header)
            Else
                data(rowIndex, colIndex) = vbNullString
            End If
            colIndex = colIndex + 1
        Next header
        rowIndex = rowIndex + 1
    Next dict

    BuildProjectRows = data
End Function

Public Function LoadProjectsFromWorksheet(ByVal ws As Worksheet) As Variant
    If ws Is Nothing Then
        LoadProjectsFromWorksheet = Empty
        Exit Function
    End If

    Dim used As Range
    Set used = ws.UsedRange
    If used Is Nothing Then
        LoadProjectsFromWorksheet = Empty
        Exit Function
    End If

    Dim headerRow As Range
    Set headerRow = used.Rows(1)
    Dim headers() As String
    Dim colCount As Long
    colCount = headerRow.Columns.Count
    If colCount = 0 Then
        LoadProjectsFromWorksheet = Empty
        Exit Function
    End If

    ReDim headers(1 To colCount)
    Dim i As Long
    For i = 1 To colCount
        headers(i) = Trim$(CStr(headerRow.Cells(1, i).Value))
    Next i

    Dim rows As Collection
    Set rows = New Collection

    Dim rowIdx As Long
    For rowIdx = 2 To used.Rows.Count
        Dim hasValue As Boolean
        Dim dict As Scripting.Dictionary
        Set dict = New Scripting.Dictionary
        dict.CompareMode = TextCompare

        For i = 1 To colCount
            Dim value As Variant
            value = ws.Cells(rowIdx, i).Value
            If Not IsEmpty(value) And value <> "" Then hasValue = True
            dict(headers(i)) = value
        Next i

        If hasValue Then rows.Add dict
    Next rowIdx

    LoadProjectsFromWorksheet = rows
End Function

Public Function ExportTasksToWorkbook(ByVal client As IQuavisClient, _
                                      ByVal project As Variant, _
                                      ByVal targetWorkbook As Workbook, _
                                      Optional ByVal includes As Variant, _
                                      Optional ByVal extraHeaders As Variant, _
                                      Optional ByVal projectRows As Variant) As Long
    If client Is Nothing Then Err.Raise vbObjectError + 2001, "ExportTasksToWorkbook", "client が Nothing です。"
    If targetWorkbook Is Nothing Then Err.Raise vbObjectError + 2002, "ExportTasksToWorkbook", "targetWorkbook が指定されていません。"

    Dim identity As Variant
    identity = client.ProjectIdentity(project)
    Dim projectId As String
    Dim projectName As String
    projectId = identity(0)
    projectName = identity(1)

    If Len(projectId) = 0 Then Err.Raise vbObjectError + 2003, "ExportTasksToWorkbook", "プロジェクト ID を解決できません。"

    Dim tasks As Variant
    tasks = client.ListTasks(projectId, , includes)

    Dim unwrapped As Collection
    Set unwrapped = New Collection
    If IsObject(tasks) Then
        Dim t As Variant
        For Each t In tasks
            unwrapped.Add client.UnwrapTask(t)
        Next t
    End If

    Dim flatRows As Collection
    Set flatRows = BuildFlatRows(unwrapped)

    Dim headers As Collection
    If IsMissing(extraHeaders) Or IsEmpty(extraHeaders) Then
        Set headers = CollectHeaders(flatRows)
    Else
        Set headers = CollectHeaders(flatRows, extraHeaders)
    End If

    PopulateWorkbookWithTasks targetWorkbook, flatRows, headers, projectRows

    ExportTasksToWorkbook = flatRows.Count
End Function

Public Function ApplyTaskUpdates(ByVal client As IQuavisClient, _
                                 ByVal wsTasks As Worksheet, _
                                 Optional ByVal wsOriginal As Worksheet) As Collection
    If client Is Nothing Then Err.Raise vbObjectError + 2004, "ApplyTaskUpdates", "client が Nothing です。"
    If wsTasks Is Nothing Then Err.Raise vbObjectError + 2005, "ApplyTaskUpdates", "tasks シートが指定されていません。"

    Dim rows As Collection
    Set rows = CollectTaskRows(wsTasks, wsOriginal)

    Dim results As Collection
    Set results = New Collection

    Dim rowInfo As Scripting.Dictionary
    For Each rowInfo In rows
        Dim result As Scripting.Dictionary
        Set result = New Scripting.Dictionary
        result.CompareMode = TextCompare
        result("RowIndex") = rowInfo("RowIndex")

        Dim taskId As String
        Dim projectId As String
        taskId = CStr(rowInfo("TaskId"))
        projectId = CStr(rowInfo("ProjectId"))

        Dim values As Scripting.Dictionary
        Set values = rowInfo("Values")

        Dim payload As Scripting.Dictionary
        Set payload = BuildUpdatePayload(values)

        On Error GoTo Failed
        If Len(taskId) = 0 Then Err.Raise vbObjectError + 2006, "ApplyTaskUpdates", "Task ID が空です。"
        If Len(projectId) = 0 Then Err.Raise vbObjectError + 2007, "ApplyTaskUpdates", "Project ID が空です。"

        client.UpdateTask projectId, taskId, payload
        result("Success") = True
        result("Message") = "Updated"
        SetFillColor wsTasks, rowInfo("YellowCells"), COLOR_SUCCESS
        GoTo AppendResult
Failed:
        result("Success") = False
        result("Message") = Err.Description
        SetFillColor wsTasks, rowInfo("YellowCells"), COLOR_FAILURE
        Err.Clear
        On Error GoTo 0
AppendResult:
        results.Add result
        On Error GoTo 0
    Next rowInfo

    Set ApplyTaskUpdates = results
End Function

Public Function SummarizeResults(ByVal results As Collection) As String
    Dim total As Long
    Dim success As Long
    Dim failure As Long

    Dim item As Scripting.Dictionary
    For Each item In results
        total = total + 1
        If item.Exists("Success") And item("Success") = True Then
            success = success + 1
        Else
            failure = failure + 1
        End If
    Next item

    SummarizeResults = "Total: " & total & " / Success: " & success & " / Failure: " & failure
End Function
