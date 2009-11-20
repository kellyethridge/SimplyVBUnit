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
            Call Test.Using(True, "True")
            Call Test.Using(False, "False")
            Call Test.Using("actual", Quote("actual"))
            Call Test.Using(1.234, "1.234")
            
            
        Case "WriteExpectedValue_WithValues_WritesExpected"
            Call Test.Using(True, "True")
            Call Test.Using(False, "False")
            Call Test.Using("expected", Quote("expected"))
            Call Test.Using(1.234, "1.234")
            
        Case "WriteValue_WithValues_WritesExpected"
            Call Test.Using("", "<Empty String>")
            Call Test.Using(Null, "Null")
            Call Test.Using(Empty, "Empty")
            Call Test.Using(Array(1, 2, 3), "< 1, 2, 3 >")
            Call Test.Using(Nothing, "Nothing")
            Call Test.Using(, "Missing")
            Call Test.Using("string", Quote("string"))
            
            Dim List As New Collection
            List.Add 3
            List.Add 2
            List.Add 1
            Call Test.Using(List, "< 3, 2, 1 >")
            
            Dim Unk As IUnknown
            Call Test.Using(Unk, "Nothing")
    End Select
End Sub


Public Sub WriteValue_WithValues_WritesExpected(ByRef ActualValue As Variant, ByVal Expected As String)
    Dim w As New TextMessageWriter
    
    w.WriteValue ActualValue
    
    Dim Actual As String
    Actual = w.ToString
    Assert.IsTrue Actual = Expected, "  Expected: " & Expected & vbCrLf & "  But was:  " & Actual & vbCrLf
End Sub

Public Sub WriteActualValue_WithValues_WritesExpected(ByRef ActualValue As Variant, ByVal Expected As String)
    Dim w As New TextMessageWriter
    
    w.WriteActualValue ActualValue
    
    Dim Actual As String
    Actual = w.ToString
    Assert.IsTrue Actual = Expected, "  Expected: " & Expected & vbCrLf & "  But was:  " & Actual & vbCrLf
End Sub

Public Sub WriteExpectedValue_WithValues_WritesExpected(ByRef ExpectedValue As Variant, ByVal Expected As String)
    Dim w As New TextMessageWriter
    
    w.WriteExpectedValue ExpectedValue
    
    Dim Actual As String
    Actual = w.ToString
    Assert.IsTrue Actual = Expected, "Expected: " & Expected & vbCrLf & " But was: " & Actual & vbCrLf
End Sub

Public Sub DisplayConstraintDifference_UsingConstraint_WritesExpectedButWasClause()
    Dim Stub As New StubConstraint
    Stub.ActualValue = True
    Stub.ExpectedValue = False
    Dim w As New TextMessageWriter
    
    w.DisplayConstraintDifference Stub
    
    Dim Actual As String
    Actual = w.ToString
    Assert.IsTrue Actual = "Expected: False" & vbCrLf & "But was : True" & vbCrLf
End Sub

Public Sub WriteLine_WithNoArgs_AppendsCrLfToMessage()
    Dim w As New TextMessageWriter
    
    w.WriteLine
    
    Dim Actual As String
    Actual = w.ToString
    Assert.IsTrue Actual = vbCrLf
End Sub

Public Sub WriteLine_WithText_AppendsTextThenCrLf()
    Dim w As New TextMessageWriter
    
    w.WriteLine "Text"
    
    Dim Actual As String
    Actual = w.ToString
    Assert.IsTrue Actual = "Text" & vbCrLf
End Sub

Public Sub WriteLine_WithArgs_AppendsFormattedArgs()
    Dim w As New TextMessageWriter
    
    w.WriteLine "{0} text", "Some"
    
    Dim Actual As String
    Actual = w.ToString
    Assert.IsTrue Actual = "Some text" & vbCrLf
End Sub

Public Sub DisplayDifference_WithTrueAndFalse_WritesExpectedButWasClause()
    Dim w As New TextMessageWriter
    
    w.DisplayDifference True, False
    
    Dim Actual As String
    Actual = w.ToString
    Const Expected As String = "Expected: True" & vbCrLf & "But was : False" & vbCrLf
    Assert.IsTrue Actual = Expected
End Sub

Public Sub DisplayDifference_WithFalseAndTrue_WritesExpectedButWasClause()
    Dim w As New TextMessageWriter
    
    w.DisplayDifference False, True
    
    Dim Actual As String
    Actual = w.ToString
    Const Expected As String = "Expected: False" & vbCrLf & "But was : True" & vbCrLf
    Assert.IsTrue Actual = Expected
End Sub

Public Sub WriteText_WhenCalledMultipleTimes_AppendTextToMessage()
    Dim w As New TextMessageWriter
    
    w.WriteText "hi"
    w.WriteText "bye"
    
    Dim Actual As String
    Actual = w.ToString
    Assert.IsTrue Actual = "hibye"
End Sub

Public Sub WriteText_WithArgs_AppendsFormattedArgs()
    Dim w As New TextMessageWriter
    
    w.WriteText "{0} text", "Some"
    
    Dim Actual As String
    Actual = w.ToString
    Assert.IsTrue Actual = "Some text"
End Sub

Public Sub WriteText_WithVariantArgs_ReturnsFormattedText()
    Dim w As New TextMessageWriter
    Dim v1 As Variant
    v1 = "hello"
    
    w.WriteText "{0},{1}", v1, "bye"
    
    Dim Actual As String
    Actual = w.ToString
    Assert.IsTrue Actual = "hello,bye"
End Sub

Public Sub WriteLine_WithVariantArgs_ReturnsFormattedText()
    Dim w As New TextMessageWriter
    Dim v1 As Variant
    v1 = "hello"
    
    w.WriteLine "{0},{1}", v1, "bye"
    
    Dim Actual As String
    Actual = w.ToString
    Assert.IsTrue Actual = "hello,bye" & vbCrLf
End Sub


