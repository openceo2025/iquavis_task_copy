VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IQuavisClient"
Option Explicit

'=============================================================
' IQuavisClient
' ------------------------------------------------------------
' - 軽量な iQUAVIS Web API クライアント
' - /token でのパスワード認証
' - プロジェクト / タスクの取得、タスク更新
' - VBA-JSON (JsonConverter.bas) への依存あり
'=============================================================

Private Const DEFAULT_BASE_URL As String = "http://rdgpm0701/iquavis-api"

Private mBaseUrl As String
Private mAccessToken As String
Private mTimeoutSeconds As Long
Private mDebug As Boolean

Private Sub Class_Initialize()
    mBaseUrl = DEFAULT_BASE_URL
    mTimeoutSeconds = 60
    mDebug = False
End Sub

'------------------------------
' Properties
'------------------------------
Public Property Get BaseUrl() As String
    BaseUrl = mBaseUrl
End Property

Public Property Let BaseUrl(ByVal value As String)
    If Len(Trim$(value)) > 0 Then
        mBaseUrl = value
    Else
        mBaseUrl = DEFAULT_BASE_URL
    End If
End Property

Public Property Get AccessToken() As String
    AccessToken = mAccessToken
End Property

Public Property Get TimeoutSeconds() As Long
    TimeoutSeconds = mTimeoutSeconds
End Property

Public Property Let TimeoutSeconds(ByVal value As Long)
    If value > 0 Then
        mTimeoutSeconds = value
    End If
End Property

Public Property Get DebugLogging() As Boolean
    DebugLogging = mDebug
End Property

Public Property Let DebugLogging(ByVal value As Boolean)
    mDebug = value
End Property

'------------------------------
' Public API
'------------------------------
Public Function Login(ByVal userId As String, ByVal password As String) As String
    Dim url As String
    url = BuildUrl("/token")

    Dim http As Object
    Set http = CreateHttpClient()

    Dim payload As String
    payload = "grant_type=password&username=" & EncodeUriComponent(userId) & _
              "&password=" & EncodeUriComponent(password)

    DebugMessage "POST " & url & " (password grant)"

    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.send payload

    EnsureSuccess http, "Login failed"

    Dim responseJson As Variant
    responseJson = ParseJsonSafe(http.responseText)

    If IsObject(responseJson) Then
        On Error Resume Next
        mAccessToken = CStr(responseJson("access_token"))
        On Error GoTo 0
    End If

    If LenB(mAccessToken) = 0 Then
        Err.Raise vbObjectError + 1000, TypeName(Me), "access_token がレスポンスに含まれていません。"
    End If

    Login = mAccessToken
End Function

Public Function ListProjects(Optional ByVal nameFilter As String) As Variant
    Dim params As Scripting.Dictionary
    Set params = New Scripting.Dictionary
    If Len(Trim$(nameFilter)) > 0 Then
        params.Add "name", nameFilter
    End If

    Dim jsonResponse As Variant
    jsonResponse = ExecuteJsonRequest("GET", "/v1/projects", params, Empty)

    If IsObject(jsonResponse) Then
        ListProjects = jsonResponse
    Else
        ListProjects = Nothing
    End If
End Function

Public Function ListTasks(ByVal projectId As String, _
                          Optional ByVal nameFilter As String, _
                          Optional includeItems As Variant, _
                          Optional ByVal count As Long = 10000) As Variant
    If Len(Trim$(projectId)) = 0 Then
        Err.Raise vbObjectError + 1001, TypeName(Me), "projectId は必須です。"
    End If

    Dim params As Scripting.Dictionary
    Set params = New Scripting.Dictionary

    If Len(Trim$(nameFilter)) > 0 Then params.Add "name", nameFilter
    If count > 0 Then params.Add "count", CStr(count)

    If Not IsMissing(includeItems) Then
        Dim includeValue As String
        includeValue = BuildIncludeString(includeItems)
        If Len(includeValue) > 0 Then params("include") = includeValue
    End If

    Dim path As String
    path = "/v1/projects/" & EncodeUriComponent(projectId) & "/tasks"

    Dim jsonResponse As Variant
    jsonResponse = ExecuteJsonRequest("GET", path, params, Empty)

    If IsObject(jsonResponse) Then
        ListTasks = jsonResponse
    Else
        ListTasks = Nothing
    End If
End Function

Public Function UpdateTask(ByVal projectId As String, _
                           ByVal taskId As String, _
                           ByVal payload As Variant) As Variant
    If Len(Trim$(projectId)) = 0 Then
        Err.Raise vbObjectError + 1002, TypeName(Me), "projectId は必須です。"
    End If
    If Len(Trim$(taskId)) = 0 Then
        Err.Raise vbObjectError + 1003, TypeName(Me), "taskId は必須です。"
    End If

    Dim body As Scripting.Dictionary
    Set body = New Scripting.Dictionary
    body.Add "Task", payload

    Dim path As String
    path = "/v1/projects/" & EncodeUriComponent(projectId) & "/tasks/" & EncodeUriComponent(taskId)

    UpdateTask = ExecuteJsonRequest("PUT", path, Nothing, body)
End Function

Public Function ProjectIdentity(ByVal project As Variant) As Variant
    Dim idValue As String
    Dim nameValue As String

    On Error Resume Next
    idValue = CStr(GetValueFromVariant(project, "Id"))
    If Len(idValue) = 0 Then idValue = CStr(GetValueFromVariant(project, "id"))
    If Len(idValue) = 0 Then idValue = CStr(GetValueFromVariant(project, "ID"))

    nameValue = CStr(GetValueFromVariant(project, "Name"))
    If Len(nameValue) = 0 Then nameValue = CStr(GetValueFromVariant(project, "name"))
    On Error GoTo 0

    Dim result(1) As String
    result(0) = idValue
    result(1) = nameValue
    ProjectIdentity = result
End Function

Public Function UnwrapTask(ByVal item As Variant) As Variant
    On Error GoTo CleanExit
    If IsObject(item) Then
        Dim inner As Variant
        inner = GetValueFromVariant(item, "Task")
        If Not IsEmpty(inner) Then
            UnwrapTask = inner
            Exit Function
        End If
    End If
CleanExit:
    UnwrapTask = item
End Function

'------------------------------
' Internal helpers
'------------------------------
Private Function BuildIncludeString(ByVal includeItems As Variant) As String
    Dim result As String

    If IsObject(includeItems) Then
        Dim item As Variant
        For Each item In includeItems
            If Len(Trim$(CStr(item))) > 0 Then
                result = result & "," & Trim$(CStr(item))
            End If
        Next item
    ElseIf IsArray(includeItems) Then
        Dim i As Long
        For i = LBound(includeItems) To UBound(includeItems)
            If Len(Trim$(CStr(includeItems(i)))) > 0 Then
                result = result & "," & Trim$(CStr(includeItems(i)))
            End If
        Next i
    ElseIf Len(Trim$(CStr(includeItems))) > 0 Then
        result = Trim$(CStr(includeItems))
    End If

    If Len(result) > 0 And Left$(result, 1) = "," Then
        result = Mid$(result, 2)
    End If

    BuildIncludeString = result
End Function

Private Function ExecuteJsonRequest(ByVal method As String, _
                                    ByVal path As String, _
                                    ByVal params As Scripting.Dictionary, _
                                    ByVal body As Variant) As Variant
    Dim http As Object
    Set http = CreateHttpClient()

    Dim url As String
    url = BuildUrl(path)
    If Not params Is Nothing Then
        Dim query As String
        query = BuildQueryString(params)
        If Len(query) > 0 Then
            url = url & "?" & query
        End If
    End If

    DebugMessage method & " " & url

    http.Open method, url, False
    http.setRequestHeader "Content-Type", "application/json"
    If LenB(mAccessToken) > 0 Then
        http.setRequestHeader "Authorization", "Bearer " & mAccessToken
    End If

    Dim payloadText As String
    If IsObject(body) Or IsArray(body) Then
        payloadText = JsonConverter.ConvertToJson(body)
    ElseIf Not IsMissing(body) And Not IsEmpty(body) Then
        payloadText = CStr(body)
    Else
        payloadText = vbNullString
    End If

    If method = "GET" Or Len(payloadText) = 0 Then
        http.send
    Else
        http.send payloadText
    End If

    EnsureSuccess http, method & " " & path & " failed"

    Dim responseText As String
    responseText = CStr(http.responseText)

    If Len(responseText) = 0 Then
        ExecuteJsonRequest = Empty
    Else
        ExecuteJsonRequest = ParseJsonSafe(responseText)
    End If
End Function

Private Function BuildUrl(ByVal path As String) As String
    Dim base As String
    base = mBaseUrl
    If Right$(base, 1) = "/" Then base = Left$(base, Len(base) - 1)
    If Len(path) = 0 Then path = "/"
    If Left$(path, 1) <> "/" Then path = "/" & path
    BuildUrl = base & path
End Function

Private Function BuildQueryString(ByVal params As Scripting.Dictionary) As String
    If params Is Nothing Then
        BuildQueryString = vbNullString
        Exit Function
    End If
    If params.Count = 0 Then
        BuildQueryString = vbNullString
        Exit Function
    End If
    Dim parts() As String
    Dim idx As Long
    ReDim parts(0 To params.Count - 1)

    Dim key As Variant
    For Each key In params.Keys
        parts(idx) = EncodeUriComponent(CStr(key)) & "=" & EncodeUriComponent(CStr(params(key)))
        idx = idx + 1
    Next key
    BuildQueryString = Join(parts, "&")
End Function

Private Function CreateHttpClient() As Object
    Dim http As Object
    On Error Resume Next
    Set http = CreateObject("MSXML2.XMLHTTP60")
    If http Is Nothing Then
        Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    End If
    On Error GoTo 0
    If http Is Nothing Then
        Err.Raise vbObjectError + 1004, TypeName(Me), "HTTP クライアントを作成できません。MSXML2 または WinHTTP への参照設定が必要です。"
    End If
    CreateHttpClient = http
End Function

Private Sub EnsureSuccess(ByVal http As Object, ByVal message As String)
    Dim statusCode As Long
    On Error Resume Next
    statusCode = CLng(http.status)
    On Error GoTo 0

    DebugMessage "-> " & statusCode & " " & Left$(CStr(http.responseText), 200)

    If statusCode < 200 Or statusCode >= 300 Then
        Dim errMsg As String
        errMsg = message & " (HTTP " & statusCode & ")"
        Err.Raise vbObjectError + 1005, TypeName(Me), errMsg
    End If
End Sub

Private Sub DebugMessage(ByVal text As String)
    If mDebug Then
        Debug.Print "[IQuavisClient] " & text
    End If
End Sub

Private Function ParseJsonSafe(ByVal jsonText As String) As Variant
    On Error GoTo Failed
    If Len(Trim$(jsonText)) = 0 Then
        ParseJsonSafe = Empty
        Exit Function
    End If
    ParseJsonSafe = JsonConverter.ParseJson(jsonText)
    Exit Function
Failed:
    DebugMessage "JSON parse failed: " & Err.Description
    ParseJsonSafe = Empty
End Function

Private Function GetValueFromVariant(ByVal source As Variant, ByVal key As String) As Variant
    If IsObject(source) Then
        If TypeOf source Is Scripting.Dictionary Then
            If source.Exists(key) Then GetValueFromVariant = source(key)
        Else
            On Error Resume Next
            GetValueFromVariant = CallByName(source, key, VbGet)
            On Error GoTo 0
        End If
    End If
End Function

Private Function EncodeUriComponent(ByVal value As String) As String
    Dim i As Long
    Dim ch As String
    Dim code As Long

    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        code = AscW(ch)
        Select Case code
            Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95, 126
                EncodeUriComponent = EncodeUriComponent & ch
            Case Else
                EncodeUriComponent = EncodeUriComponent & "%" & Right$("0" & Hex(code), 2)
        End Select
    Next i
End Function
