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
        Case "ClipString_WithValues_ReturnsExpected"
            Call Test.Use("string", 1, 6, "string")
            Call Test.Use("string", 1, 5, "st...")
            Call Test.Use("string", 2, 6, "......")
            Call Test.Use("string", 2, 7, "...t...")
            Call Test.Use("1234567890", 1, 9, "123456...")
            Call Test.Use("1234567890", 4, 10, "...4567890")
            Call Test.Use("1234567890", 5, 15, "...567890")
            
    End Select
End Sub


Public Sub FindMismatchIndex_WithEqualStrings_ReturnsZero()
    Dim Utils As New MsgUtils
    
    Const SAME_STRING As String = "string"
    Dim Actual As Long
    Actual = Utils.FindMismatchIndex(SAME_STRING, SAME_STRING, False)
    
    Const NO_MISMATCH_FOUND As Long = 0
    Assert.That Actual, Iz.EqualTo(NO_MISMATCH_FOUND)
End Sub

Public Sub FindMismatchIndex_WithFirstCharacterDifferent_ReturnsIndexOfFirstCharacter()
    Dim Utils As New MsgUtils
    
    Dim Actual As Long
    Actual = Utils.FindMismatchIndex("aaaa", "baaa", False)
    
    Const INDEX_OF_FIRST_CHARACTER As Long = 1
    Assert.That Actual, Iz.EqualTo(INDEX_OF_FIRST_CHARACTER)
End Sub

Public Sub FindMismatchIndex_WithLastCharacterDifferent_ReturnsIndexOfLastCharacter()
    Dim Utils As New MsgUtils
    
    Dim Actual As Long
    Actual = Utils.FindMismatchIndex("aaaa", "aaab", False)
    
    Const INDEX_OF_LAST_CHARACTER As Long = 4
    Assert.That Actual, Iz.EqualTo(INDEX_OF_LAST_CHARACTER)
End Sub

Public Sub FindMismatchIndex_WithCaseDifference_ReturnsIndexOfCaseDifference()
    Dim Utils As New MsgUtils
    
    Dim Actual As Long
    Actual = Utils.FindMismatchIndex("aaaa", "aAaa", False)
    
    Const INDEX_OF_CASE_DIFFERENCE As Long = 2
    Assert.That Actual, Iz.EqualTo(INDEX_OF_CASE_DIFFERENCE)
End Sub

Public Sub FindMismatchIndex_WithCaseDifferenceIgnoringCase_ReturnsZero()
    Dim Utils As New MsgUtils
    
    Dim Actual As Long
    Actual = Utils.FindMismatchIndex("aaaa", "aAaa", True)
    
    Const NO_MISMATCH_FOUND As Long = 0
    Assert.That Actual, Iz.EqualTo(NO_MISMATCH_FOUND)
End Sub

Public Sub FindMismatchIndex_WithDifferenceLengthStrings_ReturnsIndexOfFirstCharacterPastShortedString()
    Dim Utils As New MsgUtils
    
    Dim Actual As Long
    Actual = Utils.FindMismatchIndex("aaaa", "aaaaa", False)
    
    Assert.That Actual, Iz.EqualTo(5)
End Sub

Public Sub FindMismatchIndex_WithDifferenceLengthStrings_ReturnsIndexOfFirstCharacterPastShortedString2()
    Dim Utils As New MsgUtils
    
    Dim Actual As Long
    Actual = Utils.FindMismatchIndex("aaaaa", "aaaa", False)
    
    Assert.That Actual, Iz.EqualTo(5)
End Sub

Public Sub ClipString_WithValues_ReturnsExpected(ByVal SourceString As String, ByVal StartIndex As Long, ByVal MaxStringLength As Long, ByVal Expected As String)
    Dim Utils As New MsgUtils
    
    Dim Actual As String
    Actual = Utils.ClipString(SourceString, StartIndex, MaxStringLength)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub


