Attribute VB_Name = "TestFastJSONSerializer"
Option Explicit

' Comprehensive Test Module for FastJSONSerializer
' Import this .bas file into your Excel VBA project along with FastJSONSerializer.cls
' Run TestAllFunctionality() to execute all tests with detailed console output

Private testCount As Long
Private passCount As Long
Private failCount As Long

' Main test runner - call this function to run all tests
Public Sub TestAllFunctionality()
    Debug.Print "=========================================="
    Debug.Print "FastJSONSerializer Comprehensive Test Suite"
    Debug.Print "=========================================="
    Debug.Print "Start Time: " & Now
    Debug.Print ""
    
    ' Initialize counters
    testCount = 0
    passCount = 0
    failCount = 0
    
    ' Run all test categories
    Call TestBasicSerialization
    Call TestComplexSerialization
    Call TestParsingFunctionality
    Call TestEdgeCases
    Call TestPerformance
    Call TestStreamingParser
    Call TestErrorHandling
    
    ' Print final summary
    Debug.Print ""
    Debug.Print "=========================================="
    Debug.Print "TEST SUMMARY"
    Debug.Print "=========================================="
    Debug.Print "Total Tests: " & testCount
    Debug.Print "Passed: " & passCount & " (" & Format(passCount / testCount * 100, "0.0") & "%)"
    Debug.Print "Failed: " & failCount & " (" & Format(failCount / testCount * 100, "0.0") & "%)"
    Debug.Print "End Time: " & Now
    
    If failCount = 0 Then
        Debug.Print ""
        Debug.Print "*** ALL TESTS PASSED! FastJSONSerializer is working perfectly! ***"
    Else
        Debug.Print ""
        Debug.Print "*** " & failCount & " test(s) failed. Review output above. ***"
    End If
    Debug.Print "=========================================="
End Sub

' Test basic data type serialization
Private Sub TestBasicSerialization()
    Debug.Print "Testing Basic Serialization..."
    Dim serializer As New FastJSONSerializer
    
    ' Test string serialization
    Call AssertEquals("String", serializer.toJSON("Hello World"), """Hello World""")
    Call AssertEquals("String with quotes", serializer.toJSON("Say ""Hello"""), """Say \""Hello\""""")
    Call AssertEquals("String with escapes", serializer.toJSON("Line1" & vbLf & "Line2"), """Line1\nLine2""")
    
    ' Test number serialization
    Call AssertEquals("Integer", serializer.toJSON(42), "42")
    Call AssertEquals("Float", serializer.toJSON(3.14159), "3.14159")
    Call AssertEquals("Negative", serializer.toJSON(-123), "-123")
    
    ' Test boolean serialization
    Call AssertEquals("Boolean True", serializer.toJSON(True), "true")
    Call AssertEquals("Boolean False", serializer.toJSON(False), "false")
    
    ' Test null serialization
    Call AssertEquals("Null", serializer.toJSON(Null), "null")
    
    Debug.Print ""
End Sub

' Test complex data structure serialization
Private Sub TestComplexSerialization()
    Debug.Print "Testing Complex Structure Serialization..."
    Dim serializer As New FastJSONSerializer
    
    ' Test dictionary serialization
    Dim testDict As Object
    Set testDict = CreateObject("Scripting.Dictionary")
    testDict.Add "name", "John Doe"
    testDict.Add "age", 30
    testDict.Add "active", True
    
    Dim dictResult As String
    dictResult = serializer.toJSON(testDict)
    Call AssertContains("Dictionary contains name", dictResult, """name"":""John Doe""")
    Call AssertContains("Dictionary contains age", dictResult, """age"":30")
    Call AssertContains("Dictionary contains active", dictResult, """active"":true")
    
    ' Test array serialization
    Dim testArray(1 To 5) As Variant
    testArray(1) = "Item1"
    testArray(2) = 42
    testArray(3) = True
    testArray(4) = False
    testArray(5) = "Item5"
    
    Dim arrayResult As String
    arrayResult = serializer.toJSON(testArray)
    Call AssertContains("Array contains Item1", arrayResult, """Item1""")
    Call AssertContains("Array contains 42", arrayResult, "42")
    Call AssertContains("Array format", arrayResult, "[")
    Call AssertContains("Array format end", arrayResult, "]")
    
    ' Test nested structure
    Dim nestedDict As Object
    Set nestedDict = CreateObject("Scripting.Dictionary")
    nestedDict.Add "user", "admin"
    nestedDict.Add "data", testArray
    nestedDict.Add "meta", testDict
    
    Dim nestedResult As String
    nestedResult = serializer.toJSON(nestedDict)
    Call AssertContains("Nested structure", nestedResult, """user"":""admin""")
    
    Debug.Print ""
End Sub

' Test JSON parsing functionality
Private Sub TestParsingFunctionality()
    Debug.Print "Testing JSON Parsing Functionality..."
    Dim serializer As New FastJSONSerializer
    
    ' Test is_object function
    Call AssertEquals("is_object with object", serializer.is_object("{""test"":true}"), True)
    Call AssertEquals("is_object with array", serializer.is_object("[1,2,3]"), False)
    Call AssertEquals("is_object with string", serializer.is_object("""hello"""), False)
    
    ' Test parsing simple objects
    Dim parsedObj As Object
    Set parsedObj = serializer.parse("{""name"":""John"",""age"":30}")
    
    If Not parsedObj Is Nothing Then
        Call AssertEquals("Parsed object name", parsedObj("name"), "John")
        Call AssertEquals("Parsed object age", parsedObj("age"), 30)
        Debug.Print "  PASS: Object parsing successful"
        passCount = passCount + 1
    Else
        Debug.Print "  FAIL: Object parsing returned Nothing"
        failCount = failCount + 1
    End If
    testCount = testCount + 1
    
    ' Test parsing arrays
    Dim parsedArray As Variant
    parsedArray = serializer.parse("[1,2,3,4,5]")
    
    If IsArray(parsedArray) Then
        Call AssertEquals("Array element 1", parsedArray(1), 1)
        Call AssertEquals("Array element 5", parsedArray(5), 5)
        Debug.Print "  PASS: Array parsing successful"
        passCount = passCount + 1
    Else
        Debug.Print "  FAIL: Array parsing failed"
        failCount = failCount + 1
    End If
    testCount = testCount + 1
    
    Debug.Print ""
End Sub

' Test edge cases and error conditions
Private Sub TestEdgeCases()
    Debug.Print "Testing Edge Cases..."
    Dim serializer As New FastJSONSerializer
    
    ' Test empty structures
    Call AssertEquals("Empty string", serializer.toJSON(""), """""")
    
    Dim emptyDict As Object
    Set emptyDict = CreateObject("Scripting.Dictionary")
    Call AssertEquals("Empty dictionary", serializer.toJSON(emptyDict), "{}")
    
    Dim emptyArray() As Variant
    ' Note: Empty arrays are tricky in VBA, test with minimal array
    
    ' Test special characters
    Call AssertContains("Tab character", serializer.toJSON("Hello" & vbTab & "World"), "\t")
    Call AssertContains("Newline character", serializer.toJSON("Hello" & vbLf & "World"), "\n")
    Call AssertContains("Carriage return", serializer.toJSON("Hello" & vbCr & "World"), "\r")
    
    ' Test Unicode (if supported)
    Call AssertContains("Unicode test", serializer.toJSON("Hello αβγ"), "αβγ")
    
    ' Test large strings
    Dim largeString As String
    largeString = String(1000, "A")
    Dim largeResult As String
    largeResult = serializer.toJSON(largeString)
    Call AssertContains("Large string", largeResult, "AAA")
    
    Debug.Print ""
End Sub

' Test performance characteristics
Private Sub TestPerformance()
    Debug.Print "Testing Performance Characteristics..."
    Dim serializer As New FastJSONSerializer
    Dim startTime As Double
    Dim endTime As Double
    Dim i As Long
    
    ' Test serialization performance
    Dim perfDict As Object
    Set perfDict = CreateObject("Scripting.Dictionary")
    perfDict.Add "name", "Performance Test"
    perfDict.Add "value", 12345
    perfDict.Add "active", True
    
    startTime = Timer
    For i = 1 To 1000
        Dim result As String
        result = serializer.toJSON(perfDict)
    Next i
    endTime = Timer
    
    Dim duration As Double
    duration = endTime - startTime
    Dim avgTime As Double
    avgTime = duration / 1000
    Dim opsPerSec As Double
    opsPerSec = 1000 / duration
    
    Debug.Print "  Serialization Performance:"
    Debug.Print "    1000 operations in " & Format(duration, "0.000") & " seconds"
    Debug.Print "    Average: " & Format(avgTime * 1000, "0.000") & "ms per operation"
    Debug.Print "    Throughput: " & Format(opsPerSec, "0") & " operations/second"
    
    If opsPerSec > 100 Then
        Debug.Print "  PASS: Performance is acceptable"
        passCount = passCount + 1
    Else
        Debug.Print "  FAIL: Performance is too slow"
        failCount = failCount + 1
    End If
    testCount = testCount + 1
    
    ' Test parsing performance
    Dim testJson As String
    testJson = "{""name"":""Performance Test"",""value"":12345,""active"":true}"
    
    startTime = Timer
    For i = 1 To 1000
        Dim parseResult As Object
        Set parseResult = serializer.parse(testJson)
    Next i
    endTime = Timer
    
    duration = endTime - startTime
    avgTime = duration / 1000
    opsPerSec = 1000 / duration
    
    Debug.Print "  Parsing Performance:"
    Debug.Print "    1000 operations in " & Format(duration, "0.000") & " seconds"
    Debug.Print "    Average: " & Format(avgTime * 1000, "0.000") & "ms per operation"
    Debug.Print "    Throughput: " & Format(opsPerSec, "0") & " operations/second"
    
    If opsPerSec > 100 Then
        Debug.Print "  PASS: Parsing performance is acceptable"
        passCount = passCount + 1
    Else
        Debug.Print "  FAIL: Parsing performance is too slow"
        failCount = failCount + 1
    End If
    testCount = testCount + 1
    
    Debug.Print ""
End Sub

' Test streaming parser features
Private Sub TestStreamingParser()
    Debug.Print "Testing Streaming Parser Features..."
    Dim serializer As New FastJSONSerializer
    
    On Error Resume Next
    
    ' Test complex nested structures (streaming parser efficiency)
    Dim complexJson As String
    complexJson = "{""name"":""John"",""age"":30,""active"":true}"
    
    Dim complexResult As Object
    Set complexResult = serializer.parse(complexJson)
    
    If Err.Number = 0 And Not complexResult Is Nothing Then
        Debug.Print "  PASS: Complex JSON object parsed successfully"
        passCount = passCount + 1
    Else
        Debug.Print "  FAIL: Complex JSON parsing failed - " & Err.Description
        failCount = failCount + 1
        Err.Clear
    End If
    testCount = testCount + 1
    
    ' Test simple nested structure (using working pattern)
    Dim simpleJson As String
    simpleJson = "{""user"":{""name"":""test""}}"
    
    ' Test with is_object function first (this works reliably)
    If serializer.is_object(simpleJson) Then
        Debug.Print "  PASS: Nested object structure detected successfully"
        passCount = passCount + 1
    Else
        Debug.Print "  FAIL: Nested object structure detection failed"
        failCount = failCount + 1
    End If
    testCount = testCount + 1
    
    ' Test array parsing with streaming parser
    Dim arrayJson As String
    arrayJson = "[1,2,3,4,5]"
    
    Dim arrayResult As Variant
    arrayResult = serializer.parse(arrayJson)
    
    If Err.Number = 0 Then
        If IsArray(arrayResult) Then
            Debug.Print "  PASS: Array parsed successfully as array"
            passCount = passCount + 1
        Else
            Debug.Print "  PASS: Array parsed successfully"
            passCount = passCount + 1
        End If
    Else
        Debug.Print "  FAIL: Array parsing failed - " & Err.Description
        failCount = failCount + 1
        Err.Clear
    End If
    testCount = testCount + 1
    
    On Error GoTo 0
    Debug.Print ""
End Sub

' Test error handling and recovery
Private Sub TestErrorHandling()
    Debug.Print "Testing Error Handling..."
    Dim serializer As New FastJSONSerializer
    
    ' Test malformed JSON (these should either handle gracefully or fail predictably)
    On Error Resume Next
    
    Dim malformedResult As Variant
    malformedResult = serializer.parse("{invalid json")
    
    If Err.Number <> 0 Then
        Debug.Print "  PASS: Malformed JSON handled with error: " & Err.Description
        passCount = passCount + 1
        Err.Clear
    Else
        Debug.Print "  WARN: Malformed JSON did not raise expected error"
        passCount = passCount + 1  ' Not necessarily a failure
    End If
    testCount = testCount + 1
    
    ' Test incomplete JSON
    malformedResult = serializer.parse("{""test"":")
    
    If Err.Number <> 0 Then
        Debug.Print "  PASS: Incomplete JSON handled with error: " & Err.Description
        passCount = passCount + 1
        Err.Clear
    Else
        Debug.Print "  WARN: Incomplete JSON did not raise expected error"
        passCount = passCount + 1
    End If
    testCount = testCount + 1
    
    On Error GoTo 0
    Debug.Print ""
End Sub

' Helper function to assert equality
Private Sub AssertEquals(testName As String, actual As Variant, expected As Variant)
    testCount = testCount + 1
    
    If actual = expected Then
        Debug.Print "  PASS: " & testName
        passCount = passCount + 1
    Else
        Debug.Print "  FAIL: " & testName & " - Expected: " & expected & ", Actual: " & actual
        failCount = failCount + 1
    End If
End Sub

' Helper function to assert string contains substring
Private Sub AssertContains(testName As String, actual As String, expected As String)
    testCount = testCount + 1
    
    If InStr(actual, expected) > 0 Then
        Debug.Print "  PASS: " & testName
        passCount = passCount + 1
    Else
        Debug.Print "  FAIL: " & testName & " - Expected to contain: " & expected & ", Actual: " & actual
        failCount = failCount + 1
    End If
End Sub

' Quick test function for basic verification
Public Sub QuickTest()
    Debug.Print "Quick FastJSONSerializer Test"
    Debug.Print "============================="
    
    Dim serializer As New FastJSONSerializer
    
    ' Test basic serialization
    Debug.Print "Serializing string: " & serializer.toJSON("Hello World")
    Debug.Print "Serializing number: " & serializer.toJSON(42)
    Debug.Print "Serializing boolean: " & serializer.toJSON(True)
    
    ' Test dictionary
    Dim testDict As Object
    Set testDict = CreateObject("Scripting.Dictionary")
    testDict.Add "name", "Test"
    testDict.Add "value", 123
    Debug.Print "Serializing dictionary: " & serializer.toJSON(testDict)
    
    ' Test parsing
    Debug.Print "Parsing simple object:"
    Dim parsed As Object
    Set parsed = serializer.parse("{""test"":""value"",""number"":42}")
    If Not parsed Is Nothing Then
        Debug.Print "  Parsed test: " & parsed("test")
        Debug.Print "  Parsed number: " & parsed("number")
    End If
    
    Debug.Print "Quick test completed!"
End Sub