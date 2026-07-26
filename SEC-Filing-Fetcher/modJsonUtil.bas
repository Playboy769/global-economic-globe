Attribute VB_Name = "modJsonUtil"
Option Explicit

' Lightweight, targeted JSON reader. Not a general parser -- it only supports the
' two shapes SEC's JSON endpoints actually use: extracting a named value (object /
' array / string / number) from an object, and splitting a JSON array into its
' raw element substrings. Written by hand instead of parsing the whole document,
' because SEC's companyfacts JSON can run tens of MB and only a handful of
' concepts are ever needed out of it.

Public Function FindKeyValueStart(ByRef json As String, ByVal key As String, Optional ByVal startPos As Long = 1) As Long
    Dim searchStr As String
    searchStr = """" & key & """"
    Dim p As Long
    p = InStr(startPos, json, searchStr, vbBinaryCompare)
    Dim jsonLen As Long
    jsonLen = Len(json)
    Do While p > 0
        Dim q As Long
        q = p + Len(searchStr)
        Do While q <= jsonLen And Mid$(json, q, 1) = " "
            q = q + 1
        Loop
        If q <= jsonLen And Mid$(json, q, 1) = ":" Then
            q = q + 1
            Do While q <= jsonLen And (Mid$(json, q, 1) = " " Or Mid$(json, q, 1) = vbLf Or Mid$(json, q, 1) = vbCr Or Mid$(json, q, 1) = vbTab)
                q = q + 1
            Loop
            FindKeyValueStart = q
            Exit Function
        End If
        p = InStr(p + 1, json, searchStr, vbBinaryCompare)
    Loop
    FindKeyValueStart = 0
End Function

Private Function ExtractBalanced(ByRef json As String, ByVal startPos As Long) As String
    Dim openCh As String, closeCh As String
    openCh = Mid$(json, startPos, 1)
    If openCh = "{" Then closeCh = "}" Else closeCh = "]"
    Dim depth As Long
    depth = 0
    Dim i As Long, jsonLen As Long
    jsonLen = Len(json)
    Dim insideStr As Boolean
    insideStr = False
    i = startPos
    Do While i <= jsonLen
        Dim c As String
        c = Mid$(json, i, 1)
        If insideStr Then
            If c = "\" Then
                i = i + 1
            ElseIf c = """" Then
                insideStr = False
            End If
        Else
            If c = """" Then
                insideStr = True
            ElseIf c = openCh Then
                depth = depth + 1
            ElseIf c = closeCh Then
                depth = depth - 1
                If depth = 0 Then
                    ExtractBalanced = Mid$(json, startPos, i - startPos + 1)
                    Exit Function
                End If
            End If
        End If
        i = i + 1
    Loop
    ExtractBalanced = ""
End Function

Private Function ExtractQuotedRaw(ByRef json As String, ByVal startPos As Long) As String
    Dim i As Long, jsonLen As Long
    jsonLen = Len(json)
    i = startPos + 1
    Do While i <= jsonLen
        Dim c As String
        c = Mid$(json, i, 1)
        If c = "\" Then
            i = i + 2
        ElseIf c = """" Then
            ExtractQuotedRaw = Mid$(json, startPos, i - startPos + 1)
            Exit Function
        Else
            i = i + 1
        End If
    Loop
    ExtractQuotedRaw = ""
End Function

' Returns the raw JSON text of the value for "key" inside json (an object body).
' For objects/arrays this includes the surrounding {} / []; for strings it
' includes the surrounding quotes; for numbers/true/false/null it's the bare token.
Public Function ExtractJsonValueRaw(ByRef json As String, ByVal key As String, Optional ByVal startPos As Long = 1) As String
    Dim vStart As Long
    vStart = FindKeyValueStart(json, key, startPos)
    If vStart = 0 Then
        ExtractJsonValueRaw = ""
        Exit Function
    End If
    Dim ch As String
    ch = Mid$(json, vStart, 1)
    Select Case ch
        Case "{", "["
            ExtractJsonValueRaw = ExtractBalanced(json, vStart)
        Case """"
            ExtractJsonValueRaw = ExtractQuotedRaw(json, vStart)
        Case Else
            Dim e As Long, jsonLen As Long
            jsonLen = Len(json)
            e = vStart
            Do While e <= jsonLen
                Dim c As String
                c = Mid$(json, e, 1)
                If c = "," Or c = "}" Or c = "]" Then Exit Do
                e = e + 1
            Loop
            ExtractJsonValueRaw = Trim$(Mid$(json, vStart, e - vStart))
    End Select
End Function

' Strips surrounding quotes and unescapes a raw JSON string token.
Public Function JsonUnescape(ByVal s As String) As String
    Dim inner As String
    inner = s
    If Len(inner) >= 2 And Left$(inner, 1) = """" And Right$(inner, 1) = """" Then
        inner = Mid$(inner, 2, Len(inner) - 2)
    End If
    inner = Replace(inner, "\""", """")
    inner = Replace(inner, "\/", "/")
    inner = Replace(inner, "\n", vbLf)
    inner = Replace(inner, "\t", vbTab)
    inner = Replace(inner, "\\", "\")
    JsonUnescape = inner
End Function

' Splits a raw JSON array (including [ ]) into a Collection of raw element strings.
Public Function SplitJsonArrayElements(ByRef arrJson As String) As Collection
    Dim result As New Collection
    If Len(arrJson) < 2 Then
        Set SplitJsonArrayElements = result
        Exit Function
    End If
    Dim i As Long, n As Long
    i = 2
    n = Len(arrJson) - 1
    Do While i <= n
        Do While i <= n And (Mid$(arrJson, i, 1) = " " Or Mid$(arrJson, i, 1) = "," Or Mid$(arrJson, i, 1) = vbLf Or Mid$(arrJson, i, 1) = vbCr Or Mid$(arrJson, i, 1) = vbTab)
            i = i + 1
        Loop
        If i > n Then Exit Do
        Dim c As String
        c = Mid$(arrJson, i, 1)
        Dim elem As String
        If c = "{" Or c = "[" Then
            elem = ExtractBalanced(arrJson, i)
            If Len(elem) = 0 Then Exit Do
            i = i + Len(elem)
        ElseIf c = """" Then
            elem = ExtractQuotedRaw(arrJson, i)
            If Len(elem) = 0 Then Exit Do
            i = i + Len(elem)
        Else
            Dim s As Long
            s = i
            Do While i <= n And Mid$(arrJson, i, 1) <> "," And Mid$(arrJson, i, 1) <> "]"
                i = i + 1
            Loop
            elem = Trim$(Mid$(arrJson, s, i - s))
        End If
        result.Add elem
    Loop
    Set SplitJsonArrayElements = result
End Function

' Convenience: extract key "key" from json object, unescape if it was a string.
Public Function ExtractJsonString(ByRef json As String, ByVal key As String) As String
    Dim raw As String
    raw = ExtractJsonValueRaw(json, key)
    If Len(raw) >= 2 And Left$(raw, 1) = """" Then
        ExtractJsonString = JsonUnescape(raw)
    Else
        ExtractJsonString = raw
    End If
End Function
