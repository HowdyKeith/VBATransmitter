Option Explicit

'***************************************************************
' JSONParser Module
' Purpose: Handles JSON parsing and serialization for AppLauncher, including IoT devices
' Uses late binding to avoid requiring Microsoft Scripting Runtime
'***************************************************************

Private Const LOG_DIR As String = "C:\SmartTraffic\"
Private Const LOG_PREFIX As String = "server_log_"

'==============================
' Public API
'==============================

' Parse a JSON string into late-bound structures:
'   - Object  => Scripting.Dictionary (keys as String, values as Variant)
'   - Array   => Collection (1-based)
'   - String  => String
'   - Number  => Double (or Long if you cast after)
'   - true/false => Boolean
'   - null   => Null
Public Function ParseJSON(ByVal jsonText As String) As Variant
    On Error GoTo fail

    Dim pos As Long
    pos = 1
    jsonText = Trim$(jsonText)
    
    If Len(jsonText) = 0 Then
        DebugLog "ParseJSON: Empty JSON string"
        Set ParseJSON = CreateObject("Scripting.Dictionary")
        Exit Function
    End If

    SkipWS jsonText, pos
    Dim v As Variant
    v = ParseValue(jsonText, pos)

    ' Ensure no trailing content
    SkipWS jsonText, pos
    If pos <= Len(jsonText) Then
        DebugLog "ParseJSON: Trailing content at position " & pos
        Set ParseJSON = CreateObject("Scripting.Dictionary")
        Exit Function
    End If

    ' If top-level is Null or non-object, return empty dictionary
    If IsNull(v) Or Not IsObject(v) Then
        Set ParseJSON = CreateObject("Scripting.Dictionary")
    Else
        Set ParseJSON = v
    End If

    DebugLog "Successfully parsed JSON"
    Exit Function

fail:
    DebugLog "Error in ParseJSON: " & Err.description
    Set ParseJSON = CreateObject("Scripting.Dictionary")
End Function


' Serialize a late-bound JSON structure back to text.
' Accepts Dictionary, Collection, or simple types (String/Numeric/Boolean/Null)
Public Function SerializeJSON(ByVal obj As Variant) As String
    On Error GoTo fail
    SerializeJSON = SerializeValue(obj)
    Exit Function

fail:
    DebugLog "Error in SerializeJSON: " & Err.description
    SerializeJSON = "{}"
End Function

'==============================
' Core parsing helpers
'==============================

Private Function ParseObject(ByVal s As String, ByRef pos As Long) As Object
    On Error GoTo fail

    ' Expect '{'
    If Mid$(s, pos, 1) <> "{" Then GoTo fail
    pos = pos + 1
    SkipWS s, pos

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    ' Empty object?
    If pos <= Len(s) And Mid$(s, pos, 1) = "}" Then
        pos = pos + 1
        Set ParseObject = dict
        Exit Function
    End If

    Do
        ' Key must be string
        Dim key As String
        key = ParseString(s, pos)
        If Len(key) = 0 And (Err.Number <> 0) Then GoTo fail

        SkipWS s, pos
        ' Expect ':'
        If Mid$(s, pos, 1) <> ":" Then GoTo fail
        pos = pos + 1

        SkipWS s, pos
        Dim val As Variant
        val = ParseValue(s, pos)

        If IsObject(val) Then
            Set dict(key) = val
        Else
            dict(key) = val
        End If

        SkipWS s, pos
        Dim c As String
        c = Mid$(s, pos, 1)

        If c = "}" Then
            pos = pos + 1
            Set ParseObject = dict
            Exit Function
        ElseIf c = "," Then
            pos = pos + 1
            SkipWS s, pos
            ' Loop to next pair
        Else
            GoTo fail
        End If
    Loop

    ' Unreachable
    Exit Function

fail:
    DebugLog "Error in ParseObject at pos " & CStr(pos) & ": " & Err.description
    Set ParseObject = Nothing
End Function

Private Function ParseArray(ByVal s As String, ByRef pos As Long) As Object
    On Error GoTo fail

    ' Expect '['
    If Mid$(s, pos, 1) <> "[" Then GoTo fail
    pos = pos + 1
    SkipWS s, pos

    Dim coll As Collection
    Set coll = New Collection

    ' Empty array?
    If pos <= Len(s) And Mid$(s, pos, 1) = "]" Then
        pos = pos + 1
        Set ParseArray = coll
        Exit Function
    End If

    Do
        Dim v As Variant
        v = ParseValue(s, pos)
        If IsObject(v) Then
            coll.Add v
        Else
            coll.Add v
        End If

        SkipWS s, pos
        Dim c As String
        c = Mid$(s, pos, 1)

        If c = "]" Then
            pos = pos + 1
            Set ParseArray = coll
            Exit Function
        ElseIf c = "," Then
            pos = pos + 1
            SkipWS s, pos
            ' Continue
        Else
            GoTo fail
        End If
    Loop

    Exit Function

fail:
    DebugLog "Error in ParseArray at pos " & CStr(pos) & ": " & Err.description
    Set ParseArray = Nothing
End Function

Private Function ParseValue(ByVal s As String, ByRef pos As Long) As Variant
    On Error GoTo fail

    SkipWS s, pos
    If pos > Len(s) Then GoTo fail

    Dim c As String
    c = Mid$(s, pos, 1)

    Select Case c
        Case "{"
            Dim o As Object
            Set o = ParseObject(s, pos)
            If o Is Nothing Then GoTo fail
            Set ParseValue = o

        Case "["
            Dim a As Object
            Set a = ParseArray(s, pos)
            If a Is Nothing Then GoTo fail
            Set ParseValue = a

        Case """"
            ParseValue = ParseString(s, pos)

        Case "-", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
            ParseValue = ParseNumber(s, pos)

        Case Else
            ' true / false / null (case-insensitive)
            Dim look As String
            look = LCase$(Mid$(s, pos, 5)) ' safe length read

            If LCase$(Mid$(s, pos, 4)) = "true" Then
                ParseValue = True
                pos = pos + 4
            ElseIf Left$(look, 5) = "false" Then
                ParseValue = False
                pos = pos + 5
            ElseIf LCase$(Mid$(s, pos, 4)) = "null" Then
                ParseValue = Null
                pos = pos + 4
            Else
                GoTo fail
            End If
    End Select

    Exit Function

fail:
    DebugLog "Error in ParseValue at pos " & CStr(pos) & ": " & Err.description
    ParseValue = Null
End Function

Private Function ParseString(ByVal s As String, ByRef pos As Long) As String
    On Error GoTo fail

    ' Expect opening quote
    If Mid$(s, pos, 1) <> """" Then GoTo fail
    pos = pos + 1

    Dim res As String
    Dim c As String

    Do While pos <= Len(s)
        c = Mid$(s, pos, 1)

        If c = """" Then
            pos = pos + 1
            ParseString = res
            Exit Function

        ElseIf c = "\" Then
            pos = pos + 1
            If pos > Len(s) Then GoTo fail

            c = Mid$(s, pos, 1)
            Select Case c
                Case """", "\", "/"
                    res = res & c
                Case "n"
                    res = res & vbLf
                Case "t"
                    res = res & vbTab
                Case "r"
                    res = res & vbCr
                Case "b"
                    res = res & vbBack
                Case "f"
                    res = res & vbFormFeed
                Case "u"
                    ' Basic \uXXXX handling (BMP only)
                    Dim hex4 As String
                    If pos + 4 > Len(s) Then GoTo fail
                    hex4 = Mid$(s, pos + 1, 4)
                    If Not IsHex4(hex4) Then GoTo fail
                    res = res & ChrW$(CLng("&H" & hex4))
                    pos = pos + 4
                Case Else
                    ' Unknown escape; keep literal
                    res = res & c
            End Select
            pos = pos + 1

        Else
            res = res & c
            pos = pos + 1
        End If
    Loop

    ' Missing closing quote
    GoTo fail

fail:
    DebugLog "Error in ParseString at pos " & CStr(pos) & ": " & Err.description
    ParseString = ""
End Function

Private Function ParseNumber(ByVal s As String, ByRef pos As Long) As Double
    On Error GoTo fail

    Dim startPos As Long
    startPos = pos

    Dim c As String
    Dim hasDigits As Boolean

    ' Optional sign
    If Mid$(s, pos, 1) = "-" Then pos = pos + 1

    ' Integer part (0 or non-zero digit followed by digits)
    If pos <= Len(s) Then
        c = Mid$(s, pos, 1)
        If c = "0" Then
            hasDigits = True
            pos = pos + 1
        ElseIf c >= "1" And c <= "9" Then
            hasDigits = True
            Do While pos <= Len(s)
                c = Mid$(s, pos, 1)
                If c >= "0" And c <= "9" Then
                    pos = pos + 1
                Else
                    Exit Do
                End If
            Loop
        End If
    End If

    ' Fraction
    If pos <= Len(s) And Mid$(s, pos, 1) = "." Then
        pos = pos + 1
        Do While pos <= Len(s)
            c = Mid$(s, pos, 1)
            If c >= "0" And c <= "9" Then
                hasDigits = True
                pos = pos + 1
            Else
                Exit Do
            End If
        Loop
    End If

    ' Exponent
    If pos <= Len(s) Then
        c = Mid$(s, pos, 1)
        If c = "e" Or c = "E" Then
            pos = pos + 1
            If pos <= Len(s) Then
                c = Mid$(s, pos, 1)
                If c = "+" Or c = "-" Then pos = pos + 1
            End If
            Dim expHad As Boolean
            Do While pos <= Len(s)
                c = Mid$(s, pos, 1)
                If c >= "0" And c <= "9" Then
                    expHad = True
                    pos = pos + 1
                Else
                    Exit Do
                End If
            Loop
            If Not expHad Then GoTo fail
        End If
    End If

    If Not hasDigits Then GoTo fail

    Dim numText As String
    numText = Mid$(s, startPos, pos - startPos)
    ' Val() is tolerant and fast; CDbl on an empty string would error
    ParseNumber = val(numText)
    Exit Function

fail:
    DebugLog "Error in ParseNumber at pos " & CStr(pos) & ": " & Err.description
    ParseNumber = 0#
End Function

'==============================
' Serialization helpers
'==============================

Private Function SerializeValue(ByVal v As Variant) As String
    On Error GoTo fail

    If IsObject(v) Then
        Dim t As String
        t = TypeName(v)
        If t = "Dictionary" Or t = "Scripting.Dictionary" Then
            SerializeValue = SerializeDict(v)
        ElseIf t = "Collection" Then
            SerializeValue = SerializeCollection(v)
        Else
            DebugLog "SerializeValue: Unsupported object type " & t
            SerializeValue = "null"
        End If

    ElseIf IsNull(v) Then
        SerializeValue = "null"

    ElseIf VarType(v) = vbBoolean Then
        SerializeValue = IIf(v, "true", "false")

    ElseIf IsNumeric(v) Then
        ' Avoid localized decimal issues by using CStr on invariant-like values
        SerializeValue = CStr(v)

    Else
        SerializeValue = """" & EscapeJSONString(CStr(v)) & """"
    End If
    Exit Function

fail:
    DebugLog "Error in SerializeValue: " & Err.description
    SerializeValue = "null"
End Function

Private Function SerializeDict(ByVal d As Object) As String
    On Error GoTo fail
    Dim sb As String
    sb = "{"

    Dim k As Variant
    Dim first As Boolean
    first = True

    For Each k In d.Keys
        If Not first Then sb = sb & ","
        sb = sb & """" & EscapeJSONString(CStr(k)) & """:" & SerializeValue(d(k))
        first = False
    Next k

    sb = sb & "}"
    SerializeDict = sb
    Exit Function

fail:
    DebugLog "Error in SerializeDict: " & Err.description
    SerializeDict = "{}"
End Function

Private Function SerializeCollection(ByVal c As Collection) As String
    On Error GoTo fail
    Dim sb As String
    sb = "["

    Dim i As Long
    For i = 1 To c.count
        If i > 1 Then sb = sb & ","
        sb = sb & SerializeValue(c.item(i))
    Next i

    sb = sb & "]"
    SerializeCollection = sb
    Exit Function

fail:
    DebugLog "Error in SerializeCollection: " & Err.description
    SerializeCollection = "[]"
End Function

Private Function EscapeJSONString(ByVal s As String) As String
    ' Minimal correct escaping for JSON strings
    s = Replace$(s, "\", "\\")
    s = Replace$(s, """", "\""")
    s = Replace$(s, vbCr, "\r")
    s = Replace$(s, vbLf, "\n")
    s = Replace$(s, vbTab, "\t")
    EscapeJSONString = s
End Function

'==============================
' Utilities
'==============================

Private Sub SkipWS(ByVal s As String, ByRef pos As Long)
    Do While pos <= Len(s)
        Select Case Mid$(s, pos, 1)
            Case " ", vbTab, vbCr, vbLf
                pos = pos + 1
            Case Else
                Exit Do
        End Select
    Loop
End Sub

Private Function IsHex4(ByVal s As String) As Boolean
    Dim i As Long, ch As String
    If Len(s) <> 4 Then Exit Function
    For i = 1 To 4
        ch = Mid$(s, i, 1)
        If InStr(1, "0123456789abcdefABCDEF", ch, vbBinaryCompare) = 0 Then
            IsHex4 = False
            Exit Function
        End If
    Next i
    IsHex4 = True
End Function

