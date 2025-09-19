Attribute VB_Name = "modJsonHelpers"
Option Explicit

'=============================================================
' JSON / データ変換ユーティリティ
' ------------------------------------------------------------
' - VBA-JSON (JsonConverter) を利用してディクショナリ/コレクションを操作
' - フラット化/ネスト化、ヘッダーの収集などを提供
'=============================================================

Private Const SEP_DEFAULT As String = "."
' VBA-JSON の ConvertToJson で使用するコンパクトな整形モード値
Private Const JSON_WHITESPACE_COMPACT As Long = 2

Public Function FlattenDictionary(ByVal obj As Variant, _
                                  Optional ByVal parentKey As String = vbNullString, _
                                  Optional ByVal separator As String = SEP_DEFAULT) As Scripting.Dictionary
    Dim result As Scripting.Dictionary
    Set result = New Scripting.Dictionary

    If IsObject(obj) Then
        If TypeOf obj Is Scripting.Dictionary Then
            Dim key As Variant
            For Each key In obj.Keys
                AppendFlattened result, key, obj(key), parentKey, separator
            Next key
        ElseIf TypeOf obj Is Collection Then
            Dim serialized As String
            serialized = SafeConvertToJson(obj)
            If Len(parentKey) > 0 Then
                result.Add parentKey, serialized
            Else
                result.Add "value", serialized
            End If
        Else
            ' 任意のクラス: プロパティを列挙できないため JSON 化
            Dim jsonText As String
            jsonText = SafeConvertToJson(obj)
            If Len(parentKey) > 0 Then
                result.Add parentKey, jsonText
            Else
                result.Add "value", jsonText
            End If
        End If
    ElseIf IsArray(obj) Then
        Dim jsonText As String
        jsonText = SafeConvertToJson(obj)
        If Len(parentKey) > 0 Then
            result.Add parentKey, jsonText
        Else
            result.Add "value", jsonText
        End If
    Else
        If Len(parentKey) = 0 Then parentKey = "value"
        result.Add parentKey, obj
    End If

    Set FlattenDictionary = result
End Function

Private Sub AppendFlattened(ByRef result As Scripting.Dictionary, _
                            ByVal key As Variant, _
                            ByVal value As Variant, _
                            ByVal parentKey As String, _
                            ByVal separator As String)
    Dim composedKey As String
    If Len(parentKey) > 0 Then
        composedKey = parentKey & separator & CStr(key)
    Else
        composedKey = CStr(key)
    End If

    If IsObject(value) Then
        If TypeOf value Is Scripting.Dictionary Then
            Dim nested As Scripting.Dictionary
            Set nested = FlattenDictionary(value, composedKey, separator)
            CopyInto result, nested
        ElseIf TypeOf value Is Collection Then
            result.Add composedKey, SafeConvertToJson(value)
        Else
            result.Add composedKey, SafeConvertToJson(value)
        End If
    ElseIf IsArray(value) Then
        result.Add composedKey, SafeConvertToJson(value)
    Else
        result.Add composedKey, value
    End If
End Sub

Private Sub CopyInto(ByRef target As Scripting.Dictionary, ByVal source As Scripting.Dictionary)
    Dim key As Variant
    For Each key In source.Keys
        If Not target.Exists(key) Then
            target.Add key, source(key)
        Else
            target(key) = source(key)
        End If
    Next key
End Sub

Public Function CollectHeaders(ByVal rows As Collection, _
                               Optional extraHeaders As Variant) As Collection
    Dim seen As Scripting.Dictionary
    Set seen = New Scripting.Dictionary
    seen.CompareMode = TextCompare

    Dim header As Variant
    If Not IsMissing(extraHeaders) Then
        If IsArray(extraHeaders) Then
            For Each header In extraHeaders
                If Len(Trim$(CStr(header))) > 0 Then
                    seen(Trim$(CStr(header))) = True
                End If
            Next header
        ElseIf IsObject(extraHeaders) Then
            For Each header In extraHeaders
                If Len(Trim$(CStr(header))) > 0 Then
                    seen(Trim$(CStr(header))) = True
                End If
            Next header
        ElseIf Len(Trim$(CStr(extraHeaders))) > 0 Then
            seen(Trim$(CStr(extraHeaders))) = True
        End If
    End If

    Dim preferred As Variant
    preferred = Array("Id", "Name", "Type", "StartDate", "EndDate", "ProjectId", "TaskDomainId")

    Dim rowItem As Variant
    Dim dict As Scripting.Dictionary
    For Each rowItem In rows
        If TypeOf rowItem Is Scripting.Dictionary Then
            Set dict = rowItem
            For Each header In dict.Keys
                If Not seen.Exists(header) Then seen.Add header, True
            Next header
        End If
    Next rowItem

    Dim ordered As Collection
    Set ordered = New Collection

    For Each header In preferred
        If seen.Exists(header) Then
            ordered.Add header
            seen.Remove header
        End If
    Next header

    Dim remaining As Variant
    remaining = seen.Keys
    If Not IsEmpty(remaining) Then
        QuickSortStrings remaining, LBound(remaining), UBound(remaining)
        Dim idx As Long
        For idx = LBound(remaining) To UBound(remaining)
            ordered.Add remaining(idx)
        Next idx
    End If

    Set CollectHeaders = ordered
End Function

Private Sub QuickSortStrings(ByRef arr As Variant, ByVal first As Long, ByVal last As Long)
    Dim i As Long, j As Long
    Dim pivot As String
    Dim temp As Variant

    i = first
    j = last
    pivot = CStr(arr((first + last) \ 2))

    Do While i <= j
        Do While CStr(arr(i)) < pivot
            i = i + 1
        Loop
        Do While CStr(arr(j)) > pivot
            j = j - 1
        Loop
        If i <= j Then
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp
            i = i + 1
            j = j - 1
        End If
    Loop

    If first < j Then QuickSortStrings arr, first, j
    If i < last Then QuickSortStrings arr, i, last
End Sub

Public Function SanitizeFileName(ByVal name As String) As String
    Dim sanitized As String
    sanitized = Trim$(name)
    sanitized = Replace(sanitized, "\", "_")
    sanitized = Replace(sanitized, "/", "_")
    sanitized = Replace(sanitized, ":", "_")
    sanitized = Replace(sanitized, "*", "_")
    sanitized = Replace(sanitized, "?", "_")
    sanitized = Replace(sanitized, """", "_")
    sanitized = Replace(sanitized, "<", "_")
    sanitized = Replace(sanitized, ">", "_")
    sanitized = Replace(sanitized, "|", "_")
    Do While Right$(sanitized, 1) = "." Or Right$(sanitized, 1) = " "
        sanitized = Left$(sanitized, Len(sanitized) - 1)
        If Len(sanitized) = 0 Then Exit Do
    Loop
    If Len(sanitized) = 0 Then sanitized = "project"
    SanitizeFileName = sanitized
End Function

Public Function NextAvailablePath(ByVal baseDir As String, ByVal baseName As String) As String
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject

    Dim rootName As String
    Dim extension As String
    rootName = fso.GetBaseName(baseName)
    extension = "." & fso.GetExtensionName(baseName)

    Dim candidate As String
    candidate = fso.BuildPath(baseDir, rootName & extension)

    Dim counter As Long
    counter = 1
    Do While fso.FileExists(candidate)
        candidate = fso.BuildPath(baseDir, rootName & " (" & counter & ")" & extension)
        counter = counter + 1
    Loop

    NextAvailablePath = candidate
End Function

Public Function UnflattenDictionary(ByVal flat As Scripting.Dictionary, _
                                    Optional ByVal separator As String = SEP_DEFAULT) As Scripting.Dictionary
    Dim result As Scripting.Dictionary
    Set result = New Scripting.Dictionary

    Dim key As Variant
    For Each key In flat.Keys
        Dim value As Variant
        value = flat(key)
        If IsEmpty(value) Or IsNull(value) Then
            GoTo ContinueLoop
        End If

        Dim parts() As String
        parts = Split(CStr(key), separator)

        Dim cursor As Scripting.Dictionary
        Set cursor = result

        Dim i As Long
        For i = LBound(parts) To UBound(parts) - 1
            If Not cursor.Exists(parts(i)) Then
                Dim child As Scripting.Dictionary
                Set child = New Scripting.Dictionary
                cursor.Add parts(i), child
                Set cursor = child
            ElseIf TypeOf cursor(parts(i)) Is Scripting.Dictionary Then
                Set cursor = cursor(parts(i))
            Else
                Dim replacement As Scripting.Dictionary
                Set replacement = New Scripting.Dictionary
                cursor(parts(i)) = replacement
                Set cursor = replacement
            End If
        Next i

        Dim leafKey As String
        leafKey = parts(UBound(parts))
        cursor(leafKey) = value
ContinueLoop:
    Next key

    Set UnflattenDictionary = result
End Function

Public Function SafeConvertToJson(ByVal value As Variant) As String
    On Error GoTo Fallback
    SafeConvertToJson = JsonConverter.ConvertToJson(value, Whitespace:=JSON_WHITESPACE_COMPACT)
    Exit Function
Fallback:
    SafeConvertToJson = CStr(value)
End Function

Public Function NormalizeBooleanString(ByVal text As String) As Variant
    Dim lowered As String
    lowered = LCase$(Trim$(text))
    Select Case lowered
        Case "true", "1", "yes"
            NormalizeBooleanString = True
        Case "false", "0", "no"
            NormalizeBooleanString = False
        Case Else
            NormalizeBooleanString = text
    End Select
End Function
