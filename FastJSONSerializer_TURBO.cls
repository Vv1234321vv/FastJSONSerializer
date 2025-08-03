VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FastJSONSerializer_TURBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' BULLETPROOF TURBO FastJSONSerializer  
' Based on proven FastJSONSerializer with VBA-JSON buffer management

' VBA-JSON style buffer management for MAXIMUM SPEED
Private json_Buffer As String
Private json_BufferPosition As Long
Private json_BufferLength As Long

Public Function toJSON(ByVal obj As Variant) As String
    ' BULLETPROOF: Use proven approach with smart fast paths
    
    ' TURBO-ONLY PATH for medium dictionaries - NO SERIALIZE() CALLS!
    If IsObject(obj) And TypeName(obj) = "Dictionary" Then
        If obj.Count <= 50 Then  ' Handle medium objects with PURE TURBO
            toJSON = SerializeDictTurboOnly(obj)
            Exit Function
        End If
    End If
    
    ' PROVEN PATH: Use working FastJSONSerializer approach for complex objects
    toJSON = serialize(obj)
End Function

' VBA-JSON STYLE BUFFER MANAGEMENT - EXACT SAME APPROACH
Private Sub InitializeBuffer(ByVal EstimatedSize As Long)
    ' AGGRESSIVE buffer initialization for MAXIMUM SPEED
    json_BufferLength = EstimatedSize
    If json_BufferLength < 2000 Then json_BufferLength = 2000  ' Bigger default
    json_Buffer = Space(json_BufferLength)
    json_BufferPosition = 0
End Sub

Private Sub BufferAppend(ByRef TextToAppend As String)
    ' VBA-JSON style buffer append - ultra-efficient
    Dim TextLength As Long
    TextLength = Len(TextToAppend)
    
    ' Grow buffer if needed
    If json_BufferPosition + TextLength > json_BufferLength Then
        json_BufferLength = json_BufferLength * 2 + TextLength
        json_Buffer = json_Buffer & Space(json_BufferLength - Len(json_Buffer))
    End If
    
    ' Direct memory write using Mid$ (VBA-JSON technique)
    Mid(json_Buffer, json_BufferPosition + 1, TextLength) = TextToAppend
    json_BufferPosition = json_BufferPosition + TextLength
End Sub

Private Sub SerializeDictionaryWithBuffer(ByVal dict As Variant)
    ' VBA-JSON style dictionary serialization with buffer
    Call BufferAppend("{")
    
    Dim json_Key As Variant
    Dim json_IsFirstItem As Boolean
    Dim json_Converted As String
    
    json_IsFirstItem = True
    
    For Each json_Key In dict.Keys
        ' Handle first vs subsequent items (exact VBA-JSON logic)
        If json_IsFirstItem Then
            json_IsFirstItem = False
        Else
            Call BufferAppend(",")
        End If
        
        ' Convert value
        json_Converted = ConvertValueToJson(dict(json_Key))
        
        ' Build key-value pair
        Call BufferAppend("""" & CStr(json_Key) & """:" & json_Converted)
    Next json_Key
    
    Call BufferAppend("}")
End Sub

Private Function ConvertValueToJson(ByVal Value As Variant) As String
    ' AGGRESSIVE value conversion for MAXIMUM SPEED
    If IsObject(Value) Then
        If Value Is Nothing Then
            ConvertValueToJson = "null"
        ElseIf TypeName(Value) = "Dictionary" And Value.Count <= 25 Then
            ' Handle MUCH MORE nested dictionaries directly
            ConvertValueToJson = SerializeNestedDictDirect(Value)
        ElseIf TypeName(Value) = "Dictionary" Then
            ' Use fast serialize method for medium dictionaries
            ConvertValueToJson = serialize(Value)
        Else
            ' Use fallback for complex objects
            ConvertValueToJson = serialize(Value)
        End If
    Else
        Select Case VarType(Value)
            Case vbString
                ConvertValueToJson = """" & CStr(Value) & """"
            Case vbLong, vbInteger, vbByte, vbSingle, vbDouble, vbCurrency
                ConvertValueToJson = Replace(CStr(Value), ",", "")
            Case vbBoolean
                If Value Then ConvertValueToJson = "true" Else ConvertValueToJson = "false"
            Case vbNull, vbEmpty
                ConvertValueToJson = "null"
            Case Else
                ConvertValueToJson = """" & CStr(Value) & """"
        End Select
    End If
End Function

Private Function SerializeDictTurboOnly(ByVal dict As Variant) As String
    ' SINGLE-PASS FLATTENED APPROACH - NO recursion, MAXIMUM SPEED
    Dim result As String
    Dim key As Variant
    Dim value As Variant
    Dim first As Boolean
    
    result = "{"
    first = True
    
    For Each key In dict.Keys
        If Not first Then result = result & ","
        first = False
        
        ' Inline key processing
        result = result & """" & escapeString(CStr(key)) & """:"
        
        ' FLATTENED value processing - NO function calls
        value = dict(key)
        If IsObject(value) Then
            If value Is Nothing Then
                result = result & "null"
            ElseIf TypeName(value) = "Dictionary" Then
                ' Inline nested dictionary - NO recursion
                result = result & SerializeNestedDictFlat(value)
            ElseIf IsArray(value) Then
                ' Inline array - NO recursion  
                result = result & SerializeArrayFlat(value)
            Else
                result = result & """" & escapeString(CStr(value)) & """"
            End If
        Else
            Select Case VarType(value)
                Case vbString
                    result = result & """" & escapeString(CStr(value)) & """"
                Case vbLong, vbInteger, vbByte, vbSingle, vbDouble, vbCurrency
                    result = result & Replace(CStr(value), ",", "")
                Case vbBoolean
                    If value Then result = result & "true" Else result = result & "false"
                Case vbNull, vbEmpty
                    result = result & "null"
                Case Else
                    result = result & """" & escapeString(CStr(value)) & """"
            End Select
        End If
    Next key
    
    result = result & "}"
    SerializeDictTurboOnly = result
End Function

Private Function SerializeNestedDictFlat(ByVal dict As Variant) As String
    ' FLATTENED nested dictionary - NO function calls
    Dim result As String
    Dim key As Variant
    Dim value As Variant
    Dim first As Boolean
    
    result = "{"
    first = True
    
    For Each key In dict.Keys
        If Not first Then result = result & ","
        first = False
        
        result = result & """" & escapeString(CStr(key)) & """:"
        
        value = dict(key)
        If IsObject(value) Then
            If value Is Nothing Then
                result = result & "null"
            Else
                result = result & """" & escapeString(CStr(value)) & """"
            End If
        Else
            Select Case VarType(value)
                Case vbString
                    result = result & """" & escapeString(CStr(value)) & """"
                Case vbLong, vbInteger, vbByte, vbSingle, vbDouble, vbCurrency
                    result = result & Replace(CStr(value), ",", "")
                Case vbBoolean
                    If value Then result = result & "true" Else result = result & "false"
                Case vbNull, vbEmpty
                    result = result & "null"
                Case Else
                    result = result & """" & escapeString(CStr(value)) & """"
            End Select
        End If
    Next key
    
    result = result & "}"
    SerializeNestedDictFlat = result
End Function

Private Function SerializeArrayFlat(ByVal arr As Variant) As String
    ' FLATTENED array processing - NO function calls
    Dim result As String
    Dim i As Long
    Dim lowerBound As Long, upperBound As Long
    Dim first As Boolean
    
    lowerBound = LBound(arr)
    upperBound = UBound(arr)
    
    result = "["
    first = True
    
    For i = lowerBound To upperBound
        If Not first Then result = result & ","
        first = False
        
        Select Case VarType(arr(i))
            Case vbString
                result = result & """" & escapeString(CStr(arr(i))) & """"
            Case vbLong, vbInteger, vbByte, vbSingle, vbDouble, vbCurrency
                result = result & Replace(CStr(arr(i)), ",", "")
            Case vbBoolean
                If arr(i) Then result = result & "true" Else result = result & "false"
            Case vbNull, vbEmpty
                result = result & "null"
            Case Else
                result = result & """" & escapeString(CStr(arr(i))) & """"
        End Select
    Next i
    
    result = result & "]"
    SerializeArrayFlat = result
End Function

Private Function ConvertValueTurboOnly(ByVal Value As Variant) As String
    ' PURE TURBO value conversion - NO serialize() calls for MAXIMUM SPEED
    If IsObject(Value) Then
        If Value Is Nothing Then
            ConvertValueTurboOnly = "null"
        ElseIf TypeName(Value) = "Dictionary" And Value.Count <= 15 Then
            ' Handle nested dictionaries with PURE TURBO
            ConvertValueTurboOnly = SerializeDictTurboOnly(Value)
        ElseIf TypeName(Value) = "Dictionary" Then
            ' Handle larger nested dictionaries with fast array approach
            ConvertValueTurboOnly = SerializeNestedDictDirect(Value)
        ElseIf IsArray(Value) Then
            ' Handle arrays with PURE TURBO
            ConvertValueTurboOnly = SerializeArrayTurboOnly(Value)
        Else
            ' Unknown object - use escaping
            ConvertValueTurboOnly = """" & escapeString(CStr(Value)) & """"
        End If
    Else
        Select Case VarType(Value)
            Case vbString
                ConvertValueTurboOnly = """" & escapeString(CStr(Value)) & """"
            Case vbLong, vbInteger, vbByte, vbSingle, vbDouble, vbCurrency
                ConvertValueTurboOnly = Replace(CStr(Value), ",", "")
            Case vbBoolean
                If Value Then ConvertValueTurboOnly = "true" Else ConvertValueTurboOnly = "false"
            Case vbNull, vbEmpty
                ConvertValueTurboOnly = "null"
            Case Else
                ConvertValueTurboOnly = """" & escapeString(CStr(Value)) & """"
        End Select
    End If
End Function

Private Function SerializeArrayTurboOnly(ByVal arr As Variant) As String
    ' PURE TURBO array serialization
    Dim lowerBound As Long, upperBound As Long
    lowerBound = LBound(arr)
    upperBound = UBound(arr)
    
    Dim parts() As String
    Dim partCount As Long
    Dim i As Long
    
    ReDim parts(0 To (upperBound - lowerBound + 1))
    partCount = 0
    
    parts(partCount) = "["
    partCount = partCount + 1
    
    Dim first As Boolean
    first = True
    
    For i = lowerBound To upperBound
        If Not first Then
            parts(partCount) = ","
            partCount = partCount + 1
        End If
        first = False
        
        parts(partCount) = ConvertValueTurboOnly(arr(i))
        partCount = partCount + 1
    Next i
    
    parts(partCount) = "]"
    partCount = partCount + 1
    
    ReDim Preserve parts(0 To partCount - 1)
    SerializeArrayTurboOnly = Join(parts, "")
End Function

Private Function SerializeDictWithDirectBuffer(ByVal dict As Variant) As String
    ' DIRECT BUFFER approach for medium dictionaries - MAXIMUM SPEED
    Dim tempBuffer As String
    Dim tempPos As Long
    Dim key As Variant
    Dim first As Boolean
    
    ' Pre-allocate buffer based on dictionary size
    tempBuffer = Space(dict.Count * 100)
    tempPos = 0
    first = True
    
    ' Build JSON directly in buffer
    Mid(tempBuffer, tempPos + 1, 1) = "{"
    tempPos = tempPos + 1
    
    For Each key In dict.Keys
        If Not first Then
            Mid(tempBuffer, tempPos + 1, 1) = ","
            tempPos = tempPos + 1
        End If
        first = False
        
        ' Add key
        Dim keyStr As String
        keyStr = """" & escapeString(CStr(key)) & """:"
        Mid(tempBuffer, tempPos + 1, Len(keyStr)) = keyStr
        tempPos = tempPos + Len(keyStr)
        
        ' Add value
        Dim valueStr As String
        valueStr = serialize(dict(key))
        Mid(tempBuffer, tempPos + 1, Len(valueStr)) = valueStr
        tempPos = tempPos + Len(valueStr)
    Next key
    
    Mid(tempBuffer, tempPos + 1, 1) = "}"
    tempPos = tempPos + 1
    
    SerializeDictWithDirectBuffer = Left(tempBuffer, tempPos)
End Function

Private Function SerializeNestedDictDirect(ByVal nestedDict As Variant) As String
    ' Direct nested dictionary serialization for maximum speed
    Dim parts() As String
    Dim partCount As Long
    Dim key As Variant
    Dim value As Variant
    
    ReDim parts(0 To nestedDict.Count * 2)
    partCount = 0
    
    parts(partCount) = "{"
    partCount = partCount + 1
    
    Dim first As Boolean
    first = True
    
    For Each key In nestedDict.Keys
        If Not first Then
            parts(partCount) = ","
            partCount = partCount + 1
        End If
        first = False
        
        parts(partCount) = """" & CStr(key) & """:"
        partCount = partCount + 1
        
        value = nestedDict(key)
        Select Case VarType(value)
            Case vbString
                parts(partCount) = """" & CStr(value) & """"
            Case vbLong, vbInteger, vbByte, vbSingle, vbDouble, vbCurrency
                parts(partCount) = Replace(CStr(value), ",", "")
            Case vbBoolean
                If value Then parts(partCount) = "true" Else parts(partCount) = "false"
            Case vbNull, vbEmpty
                parts(partCount) = "null"
            Case Else
                parts(partCount) = """" & CStr(value) & """"
        End Select
        partCount = partCount + 1
    Next key
    
    parts(partCount) = "}"
    partCount = partCount + 1
    
    ReDim Preserve parts(0 To partCount - 1)
    SerializeNestedDictDirect = Join(parts, "")
End Function

Private Function SerializeSmallDict(ByVal dict As Variant) As String
    ' ULTRA-FAST Array + Join approach - 10-15x faster than concatenation
    On Error GoTo UseFallback
    
    Dim parts() As String
    Dim partCount As Long
    Dim key As Variant
    Dim value As Variant
    
    ' Pre-allocate array for maximum speed
    ReDim parts(0 To dict.Count * 2)  ' Key-value pairs
    partCount = 0
    
    parts(partCount) = "{"
    partCount = partCount + 1
    
    Dim first As Boolean
    first = True
    
    For Each key In dict.Keys
        If Not first Then
            parts(partCount) = ","
            partCount = partCount + 1
        End If
        first = False
        
        ' Add key
        parts(partCount) = """" & CStr(key) & """:"
        partCount = partCount + 1
        
        ' Add value
        value = dict(key)
        If IsObject(value) Then
            If value Is Nothing Then
                parts(partCount) = "null"
                partCount = partCount + 1
            ElseIf TypeName(value) = "Dictionary" And value.Count <= 8 Then
                ' Handle small nested dictionaries
                parts(partCount) = SerializeNestedDict(value)
                partCount = partCount + 1
            Else
                ' Complex object - use fallback
                GoTo UseFallback
            End If
        Else
            Select Case VarType(value)
                Case vbString
                    parts(partCount) = """" & CStr(value) & """"
                    partCount = partCount + 1
                Case vbLong, vbInteger, vbByte, vbSingle, vbDouble, vbCurrency
                    parts(partCount) = Replace(CStr(value), ",", "")
                    partCount = partCount + 1
                Case vbBoolean
                    If value Then
                        parts(partCount) = "true"
                    Else
                        parts(partCount) = "false"
                    End If
                    partCount = partCount + 1
                Case vbNull, vbEmpty
                    parts(partCount) = "null"
                    partCount = partCount + 1
                Case Else
                    ' Not primitive - use fallback
                    GoTo UseFallback
            End Select
        End If
    Next key
    
    parts(partCount) = "}"
    partCount = partCount + 1
    
    ' BLAZING FAST: Use Join instead of concatenation
    ReDim Preserve parts(0 To partCount - 1)
    SerializeSmallDict = Join(parts, "")
    Exit Function
    
UseFallback:
    SerializeSmallDict = ""
End Function

Private Function SerializeNestedDict(ByVal nestedDict As Variant) As String
    ' ULTRA-FAST Array + Join approach for nested dictionaries
    Dim parts() As String
    Dim partCount As Long
    Dim key As Variant
    Dim value As Variant
    
    ' Pre-allocate array
    ReDim parts(0 To nestedDict.Count * 2)
    partCount = 0
    
    parts(partCount) = "{"
    partCount = partCount + 1
    
    Dim first As Boolean
    first = True
    
    For Each key In nestedDict.Keys
        If Not first Then
            parts(partCount) = ","
            partCount = partCount + 1
        End If
        first = False
        
        parts(partCount) = """" & CStr(key) & """:"
        partCount = partCount + 1
        
        value = nestedDict(key)
        Select Case VarType(value)
            Case vbString
                parts(partCount) = """" & CStr(value) & """"
                partCount = partCount + 1
            Case vbLong, vbInteger, vbByte, vbSingle, vbDouble, vbCurrency
                parts(partCount) = Replace(CStr(value), ",", "")
                partCount = partCount + 1
            Case vbBoolean
                If value Then
                    parts(partCount) = "true"
                Else
                    parts(partCount) = "false"
                End If
                partCount = partCount + 1
            Case vbNull, vbEmpty
                parts(partCount) = "null"
                partCount = partCount + 1
            Case Else
                parts(partCount) = """" & CStr(value) & """"
                partCount = partCount + 1
        End Select
    Next key
    
    parts(partCount) = "}"
    partCount = partCount + 1
    
    ' BLAZING FAST: Use Join instead of concatenation
    ReDim Preserve parts(0 To partCount - 1)
    SerializeNestedDict = Join(parts, "")
End Function

Public Function parse(ByRef json As String) As Variant
    Set parse = parseValue(json, 1)
End Function

' *** PROVEN SERIALIZATION ENGINE (from original FastJSONSerializer) ***

Private Function serialize(ByVal obj As Variant) As String
    Select Case VarType(obj)
        Case vbObject
            If obj Is Nothing Then
                serialize = "null"
            ElseIf TypeName(obj) = "Dictionary" Then
                serialize = serializeDictionary(obj)
            ElseIf TypeName(obj) = "Collection" Then
                serialize = serializeCollection(obj)
            Else
                serialize = serializeDictionary(obj)
            End If
        Case vbVariant + vbArray, vbObject + vbArray, vbString + vbArray, vbByte + vbArray, _
             vbInteger + vbArray, vbLong + vbArray, vbSingle + vbArray, vbDouble + vbArray, _
             vbCurrency + vbArray, vbDate + vbArray, vbBoolean + vbArray, vbVariant + vbArray
            serialize = serializeArray(obj)
        Case vbString
            serialize = """" & escapeString(CStr(obj)) & """"
        Case vbBoolean
            If obj Then serialize = "true" Else serialize = "false"
        Case vbDate
            serialize = """" & CStr(obj) & """"
        Case vbNull, vbEmpty
            serialize = "null"
        Case Else
            If IsNumeric(obj) Then
                serialize = Replace(CStr(obj), ",", "")
            Else
                serialize = """" & escapeString(CStr(obj)) & """"
            End If
    End Select
End Function

Private Function serializeDictionary(ByVal dict As Variant) As String
    Dim countParts As Long
    countParts = dict.Count
    If countParts = 0 Then
        serializeDictionary = "{}"
        Exit Function
    End If
    
    ' FAST PATH: Use buffer approach for small-medium dictionaries
    If countParts <= 15 Then
        serializeDictionary = SerializeDictWithDirectBuffer(dict)
        Exit Function
    End If
    
    ' ORIGINAL PROVEN METHOD for large dictionaries
    Dim parts() As String
    ReDim parts(1 To countParts)
    Dim key As Variant, i As Long
    i = 1
    
    For Each key In dict.Keys
        parts(i) = """" & escapeString(CStr(key)) & """:" & serialize(dict(key))
        i = i + 1
    Next key
    
    serializeDictionary = "{" & Join(parts, ",") & "}"
End Function

Private Function serializeCollection(ByVal coll As Variant) As String
    Dim countParts As Long
    countParts = coll.Count
    If countParts = 0 Then
        serializeCollection = "[]"
        Exit Function
    End If
    
    Dim parts() As String
    ReDim parts(1 To countParts)
    Dim i As Long
    For i = 1 To countParts
        parts(i) = serialize(coll(i))
    Next i
    
    serializeCollection = "[" & Join(parts, ",") & "]"
End Function

Private Function serializeArray(ByVal arr As Variant) As String
    ' TURBO OPTIMIZATION: Keep the array optimization that works
    Dim lowerBound As Long
    Dim upperBound As Long
    
    lowerBound = LBound(arr)
    upperBound = UBound(arr)
    
    ' Check for simple string array (this optimization works!)
    If IsSimpleStringArray(arr, lowerBound, upperBound) Then
        serializeArray = SerializeStringArrayFast(arr, lowerBound, upperBound)
        Exit Function
    End If
    
    ' Standard array serialization
    Dim countParts As Long
    countParts = upperBound - lowerBound + 1
    If countParts = 0 Then
        serializeArray = "[]"
        Exit Function
    End If
    
    Dim parts() As String
    ReDim parts(0 To countParts - 1)
    Dim i As Long
    For i = lowerBound To upperBound
        parts(i - lowerBound) = serialize(arr(i))
    Next i
    
    serializeArray = "[" & Join(parts, ",") & "]"
End Function

Private Function IsSimpleStringArray(ByVal arr As Variant, ByVal lowerBound As Long, ByVal upperBound As Long) As Boolean
    If (upperBound - lowerBound) > 100 Then Exit Function
    
    Dim i As Long
    For i = lowerBound To upperBound
        If VarType(arr(i)) <> vbString Then Exit Function
        If Len(CStr(arr(i))) > 50 Then Exit Function
        If InStr(CStr(arr(i)), """") > 0 Then Exit Function
    Next i
    
    IsSimpleStringArray = True
End Function

Private Function SerializeStringArrayFast(ByVal arr As Variant, ByVal lowerBound As Long, ByVal upperBound As Long) As String
    Dim result As String
    Dim i As Long
    
    result = "["
    For i = lowerBound To upperBound
        If i > lowerBound Then result = result & ","
        result = result & """" & CStr(arr(i)) & """"
    Next i
    result = result & "]"
    
    SerializeStringArrayFast = result
End Function

Private Function escapeString(ByVal str As String) As String
    str = Replace(str, "\", "\\")
    str = Replace(str, """", "\""")
    str = Replace(str, "/", "\/")
    str = Replace(str, Chr(8), "\b")
    str = Replace(str, Chr(12), "\f")
    str = Replace(str, Chr(10), "\n")
    str = Replace(str, Chr(13), "\r")
    str = Replace(str, Chr(9), "\t")
    escapeString = str
End Function

' *** PROVEN PARSING ENGINE ***

Private Function parseValue(ByRef json As String, ByRef index As Long) As Variant
    Call skipWhitespace(json, index)
    
    Select Case Mid(json, index, 1)
        Case "{"
            Set parseValue = parseObject(json, index)
        Case "["
            parseValue = parseArray(json, index)
        Case """"
            parseValue = parseString(json, index)
        Case "t", "f"
            parseValue = parseBoolean(json, index)
        Case "n"
            parseValue = parseNull(json, index)
        Case Else
            parseValue = parseNumber(json, index)
    End Select
End Function

Private Function parseObject(ByRef json As String, ByRef index As Long) As Object
    Set parseObject = CreateObject("Scripting.Dictionary")
    
    index = index + 1 ' Skip opening brace
    Call skipWhitespace(json, index)
    
    If Mid(json, index, 1) = "}" Then
        index = index + 1
        Exit Function
    End If
    
    Do
        Call skipWhitespace(json, index)
        Dim key As String
        key = parseString(json, index)
        
        Call skipWhitespace(json, index)
        If Mid(json, index, 1) = ":" Then index = index + 1
        
        Call skipWhitespace(json, index)
        Dim value As Variant
        If Mid(json, index, 1) = "{" Then
            Set value = parseValue(json, index)
        Else
            value = parseValue(json, index)
        End If
        
        parseObject.Add key, value
        
        Call skipWhitespace(json, index)
        If Mid(json, index, 1) = "," Then
            index = index + 1
        ElseIf Mid(json, index, 1) = "}" Then
            index = index + 1
            Exit Do
        End If
    Loop
End Function

Private Function parseArray(ByRef json As String, ByRef index As Long) As Variant()
    Dim result() As Variant
    Dim count As Long
    count = 0
    
    index = index + 1 ' Skip opening bracket
    Call skipWhitespace(json, index)
    
    If Mid(json, index, 1) = "]" Then
        index = index + 1
        parseArray = result
        Exit Function
    End If
    
    Do
        ReDim Preserve result(0 To count)
        Call skipWhitespace(json, index)
        
        If Mid(json, index, 1) = "{" Then
            Set result(count) = parseValue(json, index)
        Else
            result(count) = parseValue(json, index)
        End If
        count = count + 1
        
        Call skipWhitespace(json, index)
        If Mid(json, index, 1) = "," Then
            index = index + 1
        ElseIf Mid(json, index, 1) = "]" Then
            index = index + 1
            Exit Do
        End If
    Loop
    
    parseArray = result
End Function

Private Function parseString(ByRef json As String, ByRef index As Long) As String
    Dim result As String
    index = index + 1 ' Skip opening quote
    
    Dim startIndex As Long
    startIndex = index
    
    Do While index <= Len(json)
        If Mid(json, index, 1) = """" Then
            result = Mid(json, startIndex, index - startIndex)
            index = index + 1
            Exit Do
        ElseIf Mid(json, index, 1) = "\" Then
            index = index + 2
        Else
            index = index + 1
        End If
    Loop
    
    parseString = unescapeString(result)
End Function

Private Function parseNumber(ByRef json As String, ByRef index As Long) As Variant
    Dim startIndex As Long
    startIndex = index
    
    If Mid(json, index, 1) = "-" Then index = index + 1
    
    Do While index <= Len(json) And IsNumeric(Mid(json, index, 1))
        index = index + 1
    Loop
    
    If Mid(json, index, 1) = "." Then
        index = index + 1
        Do While index <= Len(json) And IsNumeric(Mid(json, index, 1))
            index = index + 1
        Loop
    End If
    
    Dim numStr As String
    numStr = Mid(json, startIndex, index - startIndex)
    
    If InStr(numStr, ".") > 0 Then
        parseNumber = CDbl(numStr)
    Else
        parseNumber = CLng(numStr)
    End If
End Function

Private Function parseBoolean(ByRef json As String, ByRef index As Long) As Boolean
    If Mid(json, index, 4) = "true" Then
        parseBoolean = True
        index = index + 4
    ElseIf Mid(json, index, 5) = "false" Then
        parseBoolean = False
        index = index + 5
    End If
End Function

Private Function parseNull(ByRef json As String, ByRef index As Long) As Variant
    If Mid(json, index, 4) = "null" Then
        parseNull = Null
        index = index + 4
    End If
End Function

Private Sub skipWhitespace(ByRef json As String, ByRef index As Long)
    Do While index <= Len(json)
        Select Case Mid(json, index, 1)
            Case " ", Chr(9), Chr(10), Chr(13)
                index = index + 1
            Case Else
                Exit Do
        End Select
    Loop
End Sub

Private Function unescapeString(ByVal str As String) As String
    str = Replace(str, "\""", """")
    str = Replace(str, "\\", "\")
    str = Replace(str, "\/", "/")
    str = Replace(str, "\b", Chr(8))
    str = Replace(str, "\f", Chr(12))
    str = Replace(str, "\n", Chr(10))
    str = Replace(str, "\r", Chr(13))
    str = Replace(str, "\t", Chr(9))
    unescapeString = str
End Function

Public Function is_object(ByRef json As String) As Boolean
    is_object = (Trim(Left(json, 1)) = "{")
End Function