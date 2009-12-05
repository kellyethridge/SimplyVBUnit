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
            Call Test.Use(True, "True")
            Call Test.Use(False, "False")
            Call Test.Use("actual", Quote("actual"))
            Call Test.Use(1.234, "1.234")
            
            
        Case "WriteExpectedValue_WithValues_WritesExpected"
            Call Test.Use(True, "True")
            Call Test.Use(False, "False")
            Call Test.Use("expected", Quote("expected"))
            Call Test.Use(1.234, "1.234")
            
        Case "WriteValue_WithValues_WritesExpected"
            Call Test.Use("", "<Empty String>")
            Call Test.Use(Null, "Null")
            Call Test.Use(Empty, "Empty")
            Call Test.Use(Array(1, 2, 3), "< 1, 2, 3 >")
            Call Test.Use(Nothing, "Nothing")
            Call Test.Use(, "Missing")
            Call Test.Use("string", Quote("string"))
            
            Dim List As New Collection
            List.Add 3
            List.Add 2
            List.Add 1
            Call Test.Use(List, "< 3, 2, 1 >")
            
            Dim Unk As IUnknown
            Call Test.Use(Unk, "Nothing")
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
































