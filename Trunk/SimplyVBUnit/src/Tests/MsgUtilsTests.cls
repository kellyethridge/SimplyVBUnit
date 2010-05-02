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
            Call Test.Use("abc" & Chr$(14) & "de", "abc\x000Ede")
            Test.Use "abc" & ChrW$(8194) & "de", "abc\x2002de"
            Call Test.Use("", "")
        
        Case "GetTypeRepresentation_WithValues_ReturnsExpected"
            Test.Use CInt(1), "<Integer>"
            Dim Arr As Variant
            Arr = Array(1, 2)
            Test.Use Arr, "<Variant(0 To 1)>"
            
            ReDim Arr(1 To 2) As Long
            Test.Use Arr, "<Long(1 To 2)>"
            
            ReDim Arr(1 To 2, 3 To 4) As Long
            Test.Use Arr, "<Long(1 To 2, 3 To 4)>"
            
            Test.Use New Collection, "<Collection>"
        
        Case "GetArrayIndicesAsString_WithValues_ReturnsExpected"
            Dim Longs() As Long
            Longs = MakeLongArray(1, 1, 2, 3)
            Test.Use Longs, "1,2,3"
            
            Erase Longs
            Test.Use Longs, ""
            
        Case "GetArrayIndicesFromCollectionIndex_WithValues_ReturnsExpected"
            Test.Use New Collection, 1, NewLongs(0)
            Test.Use NewCollection(1, 2), 2, MakeLongArray(0, 2)
            Test.Use MakeLongArray(0, 1, 2), 1, MakeLongArray(0, 1)
            Test.Use Array(), 0, NewLongs(0)
            
            Dim Long1() As Long
            ReDim Long1(0 To 2, 0 To 2, 0 To 3)
            Test.Use Long1, 0, MakeLongArray(0, 0, 0, 0)
            Test.Use Long1, 5, MakeLongArray(0, 2, 1, 0)
            Test.Use Long1, 8, MakeLongArray(0, 2, 2, 0)
            Test.Use Long1, 14, MakeLongArray(0, 2, 1, 1)
            Test.Use Long1, 24, MakeLongArray(0, 0, 2, 2)
            Test.Use Long1, 11, MakeLongArray(0, 2, 0, 1)
            Test.Use Long1, 19, MakeLongArray(0, 1, 0, 2)
            Test.Use Long1, 25, MakeLongArray(0, 1, 2, 2)
            Test.Use Long1, 31, MakeLongArray(0, 1, 1, 3)
            
    End Select
End Sub

' Dim Long1(0 To 2, 0 To 2, 0 To 3)
'(0,0,0) =     0; Address: 105034064
'(1,0,0) =     1; Address: 105034068
'(2,0,0) =     2; Address: 105034072
'(0,1,0) =     3; Address: 105034076
'(1,1,0) =     4; Address: 105034080
'(2,1,0) =     5; Address: 105034084
'(0,2,0) =     6; Address: 105034088
'(1,2,0) =     7; Address: 105034092
'(2,2,0) =     8; Address: 105034096
'(0,0,1) =     9; Address: 105034100
'(1,0,1) =    10; Address: 105034104
'(2,0,1) =    11; Address: 105034108
'(0,1,1) =    12; Address: 105034112
'(1,1,1) =    13; Address: 105034116
'(2,1,1) =    14; Address: 105034120
'(0,2,1) =    15; Address: 105034124
'(1,2,1) =    16; Address: 105034128
'(2,2,1) =    17; Address: 105034132
'(0,0,2) =    18; Address: 105034136
'(1,0,2) =    19; Address: 105034140
'(2,0,2) =    20; Address: 105034144
'(0,1,2) =    21; Address: 105034148
'(1,1,2) =    22; Address: 105034152
'(2,1,2) =    23; Address: 105034156
'(0,2,2) =    24; Address: 105034160
'(1,2,2) =    25; Address: 105034164
'(2,2,2) =    26; Address: 105034168
'(0,0,3) =    27; Address: 105034172
'(1,0,3) =    28; Address: 105034176
'(2,0,3) =    29; Address: 105034180
'(0,1,3) =    30; Address: 105034184
'(1,1,3) =    31; Address: 105034188
'(2,1,3) =    32; Address: 105034192
'(0,2,3) =    33; Address: 105034196
'(1,2,3) =    34; Address: 105034200
'(2,2,3) =    35; Address: 105034204

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

Public Sub GetTypeRepresentation_WithValues_ReturnsExpected(ByRef Value As Variant, ByVal Expected As String)
    Dim Utils As New MsgUtils
    
    Dim Actual As String
    Actual = Utils.GetTypeRepresentation(Value)
    
    assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetArrayIndicesAsString_WithValues_ReturnsExpected(ByRef Values() As Long, ByVal Expected As String)
    Dim Utils As New MsgUtils
    
    Dim Actual As String
    Actual = Utils.GetArrayIndicesAsString(Values)
    
    assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetArrayIndicesFromCollectionIndex_WithValues_ReturnsExpected(ByRef Value As Variant, ByVal Index As Long, ByRef Expected() As Long)
    Dim Utils As New MsgUtils
    
    Dim Actual() As Long
    Actual = Utils.GetArrayIndicesFromCollectionIndex(Value, Index)
    
    assert.That Actual, Iz.EqualTo(Expected)
End Sub
