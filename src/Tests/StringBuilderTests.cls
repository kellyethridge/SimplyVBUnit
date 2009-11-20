VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringBuilderTests"
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
' Module: StringBuilderTests
'
Option Explicit

Public Sub Constructor_WhenCalled_SetsStringToEmpty()
    Dim sb As New StringBuilder
    
    Dim Actual As String
    Actual = sb.ToString
    
    Assert.IsTrue Actual = ""
End Sub

Public Sub Constructor_WhenCalled_SetsStringToEmpty2()
    Dim sb As New StringBuilder
    
    Dim Actual As Long
    Actual = sb.Length
    
    Assert.IsTrue Actual = 0
End Sub

Public Sub Append_WithEmptyString_AppendsNothing()
    Dim sb As New StringBuilder
    
    sb.Append ""
    
    Dim Actual As String
    Actual = sb.ToString
    Assert.IsTrue Actual = ""
End Sub

Public Sub Append_WithText_AppendsToString()
    Dim sb As New StringBuilder
    
    sb.Append "Text"
    
    Dim Actual As String
    Actual = sb.ToString
    Assert.IsTrue Actual = "Text"
End Sub

Public Sub Append_WithText_AppendsToString2()
    Dim sb As New StringBuilder
    
    sb.Append "Text"
    
    Dim Actual As Long
    Actual = sb.Length
    Assert.IsTrue Actual = 4
End Sub

Public Sub Append_WithMultipleCalls_AppendsAllText()
    Dim sb As New StringBuilder
    
    sb.Append "one"
    sb.Append "two"
    
    Dim Actual As String
    Actual = sb.ToString
    Assert.IsTrue Actual = "onetwo"
End Sub

Public Sub Append_WithLongString_AppendsLongString()
    Dim sb As New StringBuilder
    
    Const LONG_STRING As String = "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    sb.Append LONG_STRING
    
    Dim Actual As String
    Actual = sb.ToString
    Assert.IsTrue Actual = LONG_STRING
End Sub

Public Sub AppendFormat_With1Argument_ReplacesIndexWithArgument()
    Dim sb As New StringBuilder
    
    sb.AppendFormat "replace {0} text", "this"
    
    Dim Actual As String
    Actual = sb.ToString
    Assert.IsTrue Actual = "replace this text"
End Sub

Public Sub AppendFormat_With2Arguments_Replaces2IndexesWithArguments()
    Dim sb As New StringBuilder
    
    sb.AppendFormat "replace {0} and {1} text", "this", "that"
    
    Dim Actual As String
    Actual = sb.ToString
    Assert.IsTrue Actual = "replace this and that text"
End Sub

Public Sub AppendFormat_WithEscapedIndex_LeavesIndex()
    Dim sb As New StringBuilder
    
    sb.AppendFormat "{{0}} should be {0}.", "'this'"
    
    Dim Actual As String
    Actual = sb.ToString
    Assert.IsTrue Actual = "{0} should be 'this'."
End Sub

Public Sub Length_SettingLessThanCurrentLength_TruncatesString()
    Dim sb As New StringBuilder
    sb.Append "1234567890"
    
    sb.Length = 5
    
    Dim Actual As String
    Actual = sb.ToString
    Assert.IsTrue Actual = "12345"
End Sub

Public Sub Length_SetLessThanZero_RaisesError()
    Dim sb As New StringBuilder
    
    On Error GoTo errTrap
    sb.Length = -1
    On Error GoTo 0
    Assert.Fail "An error should be raised."

errTrap:
    Dim Actual As ErrorInfo
    Set Actual = ErrorInfo.FromErr(Err)
    Dim Expected As ErrorInfo
    Set Expected = Sim.NewErrorInfo(ErrorCode.ArgumentOutOfRange, "StringBuilder.Length", "Length cannot be negative.")
    
    Assert.IsTrue Actual.Equals(Expected)
End Sub

Public Sub AppendFormatArray_With1Argument_ReplacesIndexWithArgument()
    Dim sb As New StringBuilder
    Dim Args() As Variant
    Args = Array("this")
    
    sb.AppendFormatArray "replace {0} text", Args
    
    
    Dim Actual As String
    Actual = sb.ToString
    Assert.IsTrue Actual = "replace this text"
End Sub

Public Sub AppendFormatArray_With2Arguments_Replaces2IndexesWithArguments()
    Dim sb As New StringBuilder
    Dim Args() As Variant
    Args = Array("this", "that")
    
    sb.AppendFormatArray "replace {0} and {1} text", Args

    Dim Actual As String
    Actual = sb.ToString
    Assert.IsTrue Actual = "replace this and that text"
End Sub

Public Sub AppendFormatArray_WithEscapedIndex_LeavesIndex()
    Dim sb As New StringBuilder
    Dim Args() As Variant
    Args = Array("'this'")
    
    sb.AppendFormatArray "{{0}} should be {0}.", Args

    Dim Actual As String
    Actual = sb.ToString
    Assert.IsTrue Actual = "{0} should be 'this'."
End Sub

Public Sub Length_SetLargerThanCurrentSize_FillsWithSpace()
    Dim sb As New StringBuilder
    sb.Append "hello"
    
    sb.Length = 10
    
    Dim Actual As String
    Actual = sb.ToString
    Assert.IsTrue Actual = "hello     "
End Sub

Public Sub AppendFormat_WithVariantArgs_ReturnsFormattedString()
    Dim v1 As Variant
    v1 = "hello"
    Dim sb As New StringBuilder
    
    sb.AppendFormat "{0},{1}", v1, "bye"
    
    Dim Actual As String
    Actual = sb.ToString
    Assert.IsTrue Actual = "hello,bye"
End Sub








