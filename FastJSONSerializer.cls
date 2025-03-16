VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FastJSONSerializer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Use generic Object types for late binding
Private current_token As Long
Private dec_sep As String
Private tokens() As String
Private tokenCount As Long
Private tokenCapacity As Long

' *** Public Interface ***

Public Function parse(ByRef json As String) As Variant
    tokenize_json json
    Dim token As String
    token = fetch_token
    If Left(token, 1) = "{" Then
        Set parse = parse_dict
    Else
        parse = parse_next(token)
    End If
End Function

Public Function is_object(ByRef json As String) As Boolean
    is_object = (Left(json, 1) = "{")
End Function

' *** End of Public Interface ***

' Use the built-in Class_Initialize event for initialization
Private Sub Class_Initialize()
    ' Initialize decimal separator based on locale
    dec_sep = Mid(Format(1000, "#,##0.00"), 6, 1)
    tokenCapacity = 100 ' Initial capacity for tokens
    tokenCount = 0
    ReDim tokens(1 To tokenCapacity) ' Pre-allocate space for tokens
End Sub

Private Function parse_next(token As String) As Variant
    Select Case Left(token, 1)
        Case "{"
            Set parse_next = parse_dict
        Case "["
            parse_next = parse_array
        Case """"
            parse_next = parse_string(token)
        Case "t"
            parse_next = True
        Case "f"
            parse_next = False
        Case "n"
            parse_next = Null
        Case Else ' must be a number
            parse_next = cdbl2(token)
    End Select
End Function

Private Function cdbl2(value As Variant) As Variant
    If dec_sep = "," Then
        value = Replace(value, ".", ",")
    End If
    cdbl2 = CDbl(value)
End Function

Private Function parse_string(token As String) As String
    ' Remove surrounding quotes and unescape \n and \t
    parse_string = Replace(Replace(Mid(token, 2, Len(token) - 2), "\n", vbCrLf), "\t", vbTab)
End Function

Private Function parse_dict() As Object
    Dim token As String
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    
    Do
        token = fetch_token
        If token = "}" Then
            Set parse_dict = result
            Exit Function
        End If
        ' Use the token (after unescaping) as the key and the next token as its value.
        result.Add parse_string(token), parse_next(fetch_token)
    Loop
End Function

Private Function parse_array() As Variant
    Dim num As Long
    Dim token As String
    Dim result() As Variant
    num = 0
    Do
        token = fetch_token
        If token = "]" Then
            If num > 0 Then ReDim Preserve result(1 To num)
            parse_array = result
            Exit Function
        End If
        num = num + 1
        If (Not Not result) = 0 Then
            ReDim result(1 To 32)
        ElseIf num > UBound(result) Then
            ReDim Preserve result(1 To 2 * (num - 1))
        End If
        result(num) = parse_next(token)
    Loop
End Function

' Manual JSON Tokenizer with Pre-allocated Array
Private Sub tokenize_json(ByRef json As String)
    Dim i As Long
    Dim token As String
    Dim inString As Boolean
    Dim escape As Boolean
    inString = False
    escape = False
    i = 1
    current_token = 0
    token = ""
    
    ' Loop through the JSON string character by character
    Do While i <= Len(json)
        Dim c As String
        c = Mid(json, i, 1)
        
        If c = """" And Not escape Then
            ' End of a string token
            If inString Then
                AddToken token
                token = ""
                inString = False
            Else
                inString = True
            End If
        ElseIf c = "\" And inString Then
            ' Escape character inside string
            escape = True
        ElseIf inString Then
            token = token & c
            escape = False
        ElseIf c = "{" Or c = "}" Or c = "[" Or c = "]" Or c = ":" Or c = "," Then
            ' Separate JSON symbols (no tokenization needed)
            If Len(token) > 0 Then
                AddToken token
                token = ""
            End If
            AddToken c
        ElseIf Not (c = vbTab Or c = vbCr Or c = vbLf Or c = " " Or c = vbTab) Then
            ' Handle other non-special characters like numbers and booleans
            token = token & c
        End If
        
        i = i + 1
    Loop
    ' Add any remaining token
    If Len(token) > 0 Then AddToken token
End Sub

Private Sub AddToken(ByVal value As String)
    ' Resize the array if necessary
    If tokenCount >= tokenCapacity Then
        tokenCapacity = tokenCapacity * 2 ' Double the capacity
        ReDim Preserve tokens(1 To tokenCapacity)
    End If
    
    ' Add the token to the array
    tokenCount = tokenCount + 1
    tokens(tokenCount) = value
End Sub

Private Function fetch_token() As String
    fetch_token = tokens(current_token)
    current_token = current_token + 1
End Function

' --- JSON Writing (Serialization) Functions ---

Public Function toJSON(ByVal value As Variant) As String
    toJSON = serialize_value(value)
End Function

Private Function serialize_value(ByVal value As Variant) As String
    Select Case VarType(value)
        Case vbString
            serialize_value = """" & Replace(value, """", "\""") & """"
        Case vbBoolean
            If value Then
                serialize_value = "true"
            Else
                serialize_value = "false"
            End If
        Case vbDouble, vbInteger, vbLong, vbSingle
            serialize_value = CStr(value)
        Case vbNull
            serialize_value = "null"
        Case Else
            If TypeName(value) = "Dictionary" Then
                serialize_value = serialize_dict(value)
            ElseIf IsArray(value) Then
                serialize_value = serialize_array(value)
            Else
                serialize_value = """" & CStr(value) & """"
            End If
    End Select
End Function

Private Function serialize_dict(ByVal dict As Object) As String
    Dim countParts As Long
    countParts = dict.count
    If countParts = 0 Then
        serialize_dict = "{}"
        Exit Function
    End If
    Dim parts() As String
    ReDim parts(1 To countParts)
    Dim key As Variant, i As Long: i = 1
    For Each key In dict.Keys
        parts(i) = serialize_value(key) & ":" & serialize_value(dict(key))
        i = i + 1
    Next key
    serialize_dict = "{" & Join(parts, ",") & "}"
End Function

Private Function serialize_array(ByVal arr As Variant) As String
    Dim lb As Long, ub As Long
    lb = LBound(arr)
    ub = UBound(arr)
    Dim parts() As String
    ReDim parts(lb To ub)
    Dim i As Long
    For i = lb To ub
        parts(i) = serialize_value(arr(i))
    Next i
    serialize_array = "[" & Join(parts, ",") & "]"
End Function

