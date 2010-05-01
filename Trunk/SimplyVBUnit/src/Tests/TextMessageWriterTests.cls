VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextMessageWriterTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Copyright 2009 Kelly Ethridge
'
' Licensed under the Apache License, Version 2.0 (the "License");
' you may not use this file except in compliance with the License.
' You may obtain a copy of the License at
'
'     http://www.apache.org/licenses/LICENSE-2.0
'
' Unless required by applicable law or agreed to in writing, software
' distributed under the License is distributed on an "AS IS" BASIS,
' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
' See the License for the specific language governing permissions and
' limitations under the License.
'
' Module: TextMessageWriterTests
'
Option Explicit
Implements ITestCaseSource

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SVBUnit.TestCaseDataBuilder)
    Select Case Test.MethodName
        Case "WriteActualValue_WithValues_WritesExpected"
            Test.Use True, "True"
            Test.Use False, "False"
            Test.Use "actual", Quote("actual")
            Test.Use 1.234, "1.234#"

            
        Case "WriteExpectedValue_WithValues_WritesExpected"
            Test.Use True, "True"
            Test.Use False, "False"
            Test.Use "expected", Quote("expected")
            Test.Use 1.234, "1.234#"
            Test.Use CDec(1), "1D"
            Test.Use CDec(1.2), "1.2D"
            Test.Use 1@, "1@"
            Test.Use 1.1@, "1.1@"
            
            
        Case "WriteValue_WithValues_WritesExpected"
            Test.Use "", Quote("")
            Test.Use Null, "Null"
            Test.Use Empty, "Empty"
            Test.Use Array(1, 2, 3), "< 1, 2, 3 >"
            Test.Use Nothing, "Nothing"
            Test.Use , "Missing"
            Test.Use "string", Quote("string")
            Test.Use NewCollection("a", 1, "b"), "< ""a"", 1, ""b"" >"
            Test.Use NewCollection(Array(1, 2, 3), Array("a", "b")), "< < 1, 2, 3 >, < ""a"", ""b"" > >"
            Test.Use 123456, "123456"
            Test.Use 123.45, "123.45#"
            Test.Use 123#, "123.0#"
            Test.Use 0#, "0.0#"
            Test.Use 0!, "0.0!"
            Test.Use 123!, "123.0!"
            Test.Use 123.456!, "123.456!"
            Test.Use 1234@, "1234@"
            Test.Use 1234.5@, "1234.5@"
            Test.Use -1234@, "-1234@"
            Test.Use -1234.5@, "-1234.5@"
            Test.Use 0@, "0@"
            Test.Use CDec(0), "0D"
            Test.Use CDec("123451234512345"), "123451234512345D"
            Test.Use CDec("12345.1234"), "12345.1234D"
            Test.Use Array(), "<empty>"
            Test.Use Array(1, "name", Array("hi", "bye"), 2#), "< 1, ""name"", < ""hi"", ""bye"" >, 2.0# >"

            Dim List As New Collection
            List.Add 3
            List.Add 2
            List.Add 1
            Test.Use List, "< 3, 2, 1 >"
            
            Dim Unk As IUnknown
            Test.Use Unk, "Nothing"
            Test.Use New TextMessageWriterTests, "<TextMessageWriterTests>"
            
            Set List = New Collection
            Dim i As Long
            For i = 1 To 20
                List.Add i
            Next i
            Test.Use List, "< 1, 2, 3, 4, 5, 6, 7, 8, 9, 10... >"
            
            Dim Multi() As Long
            ReDim Multi(0 To 1, 2 To 3)
            Test.Use Multi, "< < 0, 0 >, < 0, 0 > >"

        Case "DisplayStringDifferences_WithDifferentMismatchIndexes_WritesCorrectCarrotLine"
            Test.Use "abcd", "1bcd", 1, "^"
            Test.Use "abcd", "a1cd", 2, "-^"
            Test.Use "abcd", "ab1d", 3, "--^"
            Test.Use "abcd", "abc1", 4, "---^"
            
        Case "DisplayStringDifferences_WithClipping_WritesExpected"
            Dim Expected As String
            Expected = "Expected: ""aaaaaaa...""" & vbCrLf & _
                       "But was : ""baaaaaa...""" & vbCrLf & _
                       "  ---------^" & vbCrLf
            Test.Use "aaaaaaaaaaa", "baaaaaaaaaa", 1, Expected
            
            Expected = "Expected: ""...aaaaaaa""" & vbCrLf & _
                       "But was : ""...aaaaaab""" & vbCrLf & _
                       "  ------------------^" & vbCrLf
            Test.Use "aaaaaaaaaaa", "aaaaaaaaaab", 11, Expected
            
            Expected = "Expected: ""...aaa1...""" & vbCrLf & _
                       "But was : ""...aaaa""" & vbCrLf & _
                       "  ---------------^" & vbCrLf
            Test.Use "aaaaaaaaaaaaaaa1aaaaaaaaaaaaaa", "aaaaaaaaaaaaaaaa", 16, Expected
        
        Case "WriteTypeOfPrecidate_WithValues_WritesExpected"
            Test.Use CInt(1), "type of <Integer> "
            Test.Use CLng(1), "type of <Long> "
            Test.Use CDbl(1), "type of <Double> "
            Dim IntArray() As Integer
            Test.Use IntArray, "type of <Integer()> "
            Test.Use Array(), "type of <Variant()> "
        
    End Select
End Sub


Public Sub WriteValue_WithValues_WritesExpected(ByRef ActualValue As Variant, ByVal Expected As String)
    Dim w As New TextMessageWriter
    
    Call w.WriteValue(ActualValue)
    
    Dim Actual As String
    Actual = w.ToString
    Assert.IsTrue Actual = Expected, "  Expected: " & Expected & vbCrLf & "  But was:  " & Actual & vbCrLf
End Sub

Public Sub WriteActualValue_WithValues_WritesExpected(ByRef ActualValue As Variant, ByVal Expected As String)
    Dim w As New TextMessageWriter
    
    Call w.WriteActualValue(ActualValue)
    
    Dim Actual As String
    Actual = w.ToString
    Assert.IsTrue Actual = Expected, "  Expected: " & Expected & vbCrLf & "  But was:  " & Actual & vbCrLf
End Sub

Public Sub WriteExpectedValue_WithValues_WritesExpected(ByRef ExpectedValue As Variant, ByVal Expected As String)
    Dim w As New TextMessageWriter
    
    Call w.WriteExpectedValue(ExpectedValue)
    
    Dim Actual As String
    Actual = w.ToString
    Assert.IsTrue Actual = Expected, "Expected: " & Expected & vbCrLf & " But was: " & Actual & vbCrLf
End Sub

Public Sub DisplayStringDifferences_WithDifferentMismatchIndexes_WritesCorrectCarrotLine(ByVal ExpectedString As String, ByVal ActualString As String, ByVal MismatchIndex As Long, ByVal CarrotLine As String)
    Dim Writer As New TextMessageWriter
    
    Writer.DisplayStringDifferences ExpectedString, ActualString, MismatchIndex, False, False

    Dim Actual As String
    Actual = Writer.ToString
    Dim Expected As String
    Expected = "Expected: " & Quote(ExpectedString) & vbCrLf & _
               "But was : " & Quote(ActualString) & vbCrLf & _
               "  ---------" & CarrotLine & vbCrLf
    Assert.That Actual, Iz.EqualTo(Expected).NoClip
End Sub

Public Sub DisplayStringDifferences_WithClipping_WritesExpected(ByVal ExpectedString As String, ByVal ActualString As String, ByVal MismatchIndex As Long, ByVal Expected As String)
    Dim Writer As New TextMessageWriter
    Writer.MaxDisplayLength = 10

    Writer.DisplayStringDifferences ExpectedString, ActualString, MismatchIndex, False, True

    Dim Actual As String
    Actual = Writer.ToString
    Assert.That Actual, Iz.EqualTo(Expected).NoClip
End Sub

'Public Sub DisplayFloatDifferences_WithValues_ReturnsExpected()
'    Dim Writer As New TextMessageWriter
'    Dim Tol As Tolerance
'    Set Tol = Sim.NewTolerance(1)
'
'    Writer.DisplayFloatDifferences ExpectedValue, ActualValue, Tol
'
'    Dim Actual As String
'    Actual = Writer.ToString
'    Assert.That Actual, Iz.EqualTo(Expected)
'End Sub

Public Sub DisplayStringDifferences_WithIgnoreCase_WritesIgnoringCase()
    Dim Writer As New TextMessageWriter

    Writer.DisplayStringDifferences "abcde", "bbcde", 1, True, False

    Dim Actual As String
    Actual = Writer.ToString
    Const Expected As String = "Expected: ""abcde"", ignoring case" & vbCrLf & _
                               "But was : ""bbcde""" & vbCrLf & _
                               "  ---------^" & vbCrLf
    Assert.That Actual, Iz.EqualTo(Expected).NoClip
End Sub

Public Sub DisplayConstraintDifference_UsingConstraint_WritesExpectedButWasClause()
    Dim Stub As New StubConstraint
    Stub.ActualValue = True
    Stub.Description = False
    Dim w As New TextMessageWriter
    
    Call w.DisplayConstraintDifference(Stub)
    
    Dim Actual As String
    Actual = w.ToString
    Assert.IsTrue Actual = "Expected: False" & vbCrLf & "But was : True" & vbCrLf
End Sub

Public Sub WriteLine_WithNoArgs_AppendsCrLfToMessage()
    Dim w As New TextMessageWriter
    
    Call w.WriteLine
    
    Dim Actual As String
    Actual = w.ToString
    Assert.IsTrue Actual = vbCrLf
End Sub

Public Sub WriteLine_WithText_AppendsTextThenCrLf()
    Dim w As New TextMessageWriter
    
    Call w.WriteLine("Text")
    
    Dim Actual As String
    Actual = w.ToString
    Assert.IsTrue Actual = "Text" & vbCrLf
End Sub

Public Sub WriteLine_WithArgs_AppendsFormattedArgs()
    Dim w As New TextMessageWriter
    
    Call w.WriteLine("{0} text", "Some")
    
    Dim Actual As String
    Actual = w.ToString
    Assert.IsTrue Actual = "Some text" & vbCrLf
End Sub

Public Sub DisplayDifference_WithTrueAndFalse_WritesExpectedButWasClause()
    Dim w As New TextMessageWriter
    
    Call w.DisplayDifference(True, False)
    
    Dim Actual As String
    Actual = w.ToString
    Const Expected As String = "Expected: True" & vbCrLf & "But was : False" & vbCrLf
    Assert.IsTrue Actual = Expected
End Sub

Public Sub DisplayDifference_WithFalseAndTrue_WritesExpectedButWasClause()
    Dim w As New TextMessageWriter
    
    Call w.DisplayDifference(False, True)
    
    Dim Actual As String
    Actual = w.ToString
    Const Expected As String = "Expected: False" & vbCrLf & "But was : True" & vbCrLf
    Assert.IsTrue Actual = Expected
End Sub

Public Sub WriteText_WhenCalledMultipleTimes_AppendTextToMessage()
    Dim w As New TextMessageWriter
    
    Call w.WriteText("hi")
    Call w.WriteText("bye")
    
    Dim Actual As String
    Actual = w.ToString
    Assert.IsTrue Actual = "hibye"
End Sub

Public Sub WriteText_WithArgs_AppendsFormattedArgs()
    Dim w As New TextMessageWriter
    
    Call w.WriteText("{0} text", "Some")
    
    Dim Actual As String
    Actual = w.ToString
    Assert.IsTrue Actual = "Some text"
End Sub

Public Sub WriteText_WithVariantArgs_ReturnsFormattedText()
    Dim w As New TextMessageWriter
    Dim v1 As Variant
    v1 = "hello"
    
    Call w.WriteText("{0},{1}", v1, "bye")
    
    Dim Actual As String
    Actual = w.ToString
    Assert.IsTrue Actual = "hello,bye"
End Sub

Public Sub WriteLine_WithVariantArgs_ReturnsFormattedText()
    Dim w As New TextMessageWriter
    Dim v1 As Variant
    v1 = "hello"
    
    Call w.WriteLine("{0},{1}", v1, "bye")
    
    Dim Actual As String
    Actual = w.ToString
    Assert.IsTrue Actual = "hello,bye" & vbCrLf
End Sub

Public Sub WritePredicate_WithValue_AppendValueToMessage()
    Dim w As New TextMessageWriter
    
    Call w.WritePredicate("predicate")
    
    Dim Actual As String
    Actual = w.ToString
    Assert.That Actual, Iz.EqualTo("predicate ")
End Sub

Public Sub WriteErrorInfo_WithErrorNumber_WritesNumberOnly()
    Dim Info As ErrorInfo
    Set Info = Sim.NewErrorInfo(1)
    Dim w As New TextMessageWriter
    
    Call w.WriteErrorInfo(Info)
    
    Dim Actual As String
    Actual = w.ToString
    Const Expected As String = "error 1"
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub WriteErrorInfo_WithErrorNumberAndDescription_WritesNumberAndDescription()
    Dim Info As ErrorInfo
    Set Info = Sim.NewErrorInfo(1, , "description")
    Dim w As New TextMessageWriter
    
    Call w.WriteErrorInfo(Info)
    
    Dim Actual As String
    Actual = w.ToString
    Const Expected As String = "error 1, description"
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub WriteErrorInfo_WithNumberDescriptionSource_WritesNumberDescriptionSource()
    Dim Info As ErrorInfo
    Set Info = Sim.NewErrorInfo(1, "src", "description")
    Dim w As New TextMessageWriter
    
    Call w.WriteErrorInfo(Info)
    
    Dim Actual As String
    Actual = w.ToString
    Const Expected As String = "error 1, description from [src]"
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub WriteErrorInfo_WithNumberAndSource_WritesNumberAndSource()
    Dim Info As ErrorInfo
    Set Info = Sim.NewErrorInfo(1, "src")
    Dim w As New TextMessageWriter
    
    Call w.WriteErrorInfo(Info)
    
    Dim Actual As String
    Actual = w.ToString
    Const Expected As String = "error 1 from [src]"
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub WriteConnector_WithValue_WrapsValueWithSpaces()
    Dim w As New TextMessageWriter
    
    Call w.WriteConnector("and")
    
    Dim Actual As String
    Actual = w.ToString
    Const Expected As String = " and "
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub WriteTypeOfPredicate_WithValues_WritesExpected(ByRef Value As Variant, ByRef Expected As String)
    Dim w As New TextMessageWriter
    
    w.WriteTypeOfPredicate Value
    
    Dim Actual As String
    Actual = w.ToString
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub



























