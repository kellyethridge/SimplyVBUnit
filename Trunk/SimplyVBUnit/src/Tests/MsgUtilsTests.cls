VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MsgUtilsTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Copyright 2010 Kelly Ethridge
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
' Module: MsgUtilsTests
'
Option Explicit
Implements ITestCaseSource

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SVBUnit.TestCaseDataBuilder)
    Select Case Test.MethodName
        Case "FindMismatchIndex_WithValues_ReturnsExpected"
            Call Test.Use("abcd", "abcd", True, 0)
            Call Test.Use("abcd", "abce", True, 4)
            Call Test.Use("abcD", "abcd", True, 0)
            Call Test.Use("abcD", "abcd", False, 4)
            Call Test.Use("abcd", "abcde", False, 5)
            Call Test.Use("abcde", "abcd", False, 5)
        
        Case "ClipString_WithValues_ReturnsExpected"
            Call Test.Use("string", 1, 6, "string")
            Call Test.Use("string", 1, 5, "st...")
            Call Test.Use("string", 2, 6, "......")
            Call Test.Use("string", 2, 7, "...t...")
            Call Test.Use("1234567890", 1, 9, "123456...")
            Call Test.Use("1234567890", 4, 10, "...4567890")
            Call Test.Use("1234567890", 5, 15, "...567890")
            Call Test.Use("12345678901", 1, 10, "1234567...")
            Call Test.Use("12345678901", 2, 10, "...2345...")
            Call Test.Use("12345678901", 3, 10, "...3456...")
            
        Case "ClipExpectedAndActual_WithUnclippedExpected_ReturnsClippedExpected"
            Call Test.Use("1234567890", "1234567890", 1)
            Call Test.Use("12345678901", "1234567...", 1)
            Call Test.Use("12345678901", "...5678901", 11)
            
        Case "ClipExpectedAndActual_WithUnclippedActual_ReturnsClippedExpected"
            Call Test.Use("1234567890", "1234567890", 1)
            Call Test.Use("12345678901", "1234567...", 1)
            Call Test.Use("12345678901", "...5678901", 11)
        
        Case "EscapeControlChars_WithValues_ReturnsExpected"
            Call Test.Use("abc", "abc")
            Call Test.Use("ab" & vbCr & "cd", "ab\rcd")
            Call Test.Use("ab" & vbCrLf & "cd", "ab\r\ncd")
            Call Test.Use("abc" & Chr$(14) & "de", "abc\x0Ede")
            Call Test.Use("", "")
            
    End Select
End Sub

Public Sub FindMismatchIndex_WithValues_ReturnsExpected(ByVal ExpectedString As String, ByVal ActualString As String, ByVal IgnoreCase As Boolean, ByVal Expected As Long)
    Dim Utils As New MsgUtils
    
    Dim Actual As Long
    Actual = Utils.FindMismatchIndex(ExpectedString, ActualString, IgnoreCase)
    
    assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ClipString_WithValues_ReturnsExpected(ByVal SourceString As String, ByVal StartIndex As Long, ByVal MaxStringLength As Long, ByVal Expected As String)
    Dim Utils As New MsgUtils
    
    Dim Actual As String
    Actual = Utils.ClipString(SourceString, StartIndex, MaxStringLength)
    
    assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ClipExpectedAndActual_WithUnclippedExpected_ReturnsClippedExpected(ByVal UnclippedExpected As String, ByVal Expected As String, ByVal MismatchIndex As Long)
    Dim Utils As New MsgUtils
    
    Dim Actual As String
    Actual = UnclippedExpected
    Call Utils.ClipExpectedAndActual(Actual, UnclippedExpected, 10, MismatchIndex)
    
    assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ClipExpectedAndActual_WithUnclippedActual_ReturnsClippedExpected(ByVal UnclippedActual As String, ByVal Expected As String, ByVal MismatchIndex As Long)
    Dim Utils As New MsgUtils
    
    Dim Actual As String
    Actual = UnclippedActual
    Call Utils.ClipExpectedAndActual(UnclippedActual, Actual, 10, MismatchIndex)
    
    assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub EscapeControlChars_WithValues_ReturnsExpected(ByVal Source As String, ByVal Expected As String)
    Dim Utils As New MsgUtils
    
    Dim Actual As String
    Actual = Utils.EscapeControlChars(Source)
    
    assert.That Actual, Iz.EqualTo(Expected)
End Sub
