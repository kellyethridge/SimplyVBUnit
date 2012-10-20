VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EqualConstraintTests"
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
' Module: EqualConstraintTests
'
Option Explicit
Implements ITestCaseSource

Private Sub ITestCaseSource_GetTestCases(ByVal Test As TestCaseCollector)
    Dim Long1() As Long
    Dim Long2() As Long
    Dim Array1  As Variant
    Dim Array2  As Variant

    Select Case Test.MethodName
        Case "WriteMessageTo_WithValues_WritesExpected"
            Dim Expected As String
            Expected = "  Expected: True" & vbCrLf & _
                       "  But was : False" & vbCrLf
            Test.Use True, False, Expected
            
            Expected = "  Expected string length 1 but was 2. Strings differ at index 2." & vbCrLf & _
                       "  Expected: ""a""" & vbCrLf & _
                       "  But was : ""ab""" & vbCrLf & _
                       "  ------------^" & vbCrLf
            Test.Use "a", "ab", Expected
        
            Array1 = Array(1, 2, 3, 4)
            Array2 = Array(1, 2, 3, 5)
            Expected = "  Expected and actual are both <Variant(0 To 3)>" & vbCrLf & _
                       "  Values differ at index [3]" & vbCrLf & _
                       "  Expected: 4" & vbCrLf & _
                       "  But was : 5" & vbCrLf
            Test.Use Array1, Array2, Expected
            
            ReDim Long1(0 To 1, 0 To 2)
            ReDim Long2(0 To 1, 0 To 2)
            Long1(1, 2) = 1
            Expected = "  Expected and actual are both <Long(0 To 1, 0 To 2)>" & vbCrLf & _
                       "  Values differ at index [1,2]" & vbCrLf & _
                       "  Expected: 1" & vbCrLf & _
                       "  But was : 0" & vbCrLf
            Test.Use Long1, Long2, Expected
            
            ReDim Long1(0 To 1)
            ReDim int1(0 To 1) As Integer
            Expected = "  Expected is <Long(0 To 1)>, actual is <Integer(0 To 1)>" & vbCrLf & _
                       "  Values differ at index [0]" & vbCrLf & _
                       "  Expected: 0" & vbCrLf & _
                       "  But was : 0" & vbCrLf
            
            Test.Use Long1, int1, Expected
            
            Expected = "  Expected and actual are both <Variant(0 To 2)>" & vbCrLf & _
                       "  Values differ at index [1]" & vbCrLf & _
                       "    Expected and actual are both <Long(0 To 3)>" & vbCrLf & _
                       "    Values differ at index [2]" & vbCrLf & _
                       "  Expected: 5" & vbCrLf & _
                       "  But was : 3" & vbCrLf
            Test.Use Array(1, NewLongs(1, 2, 5, 6), 3), Array(1, NewLongs(1, 2, 3, 4), 3), Expected
            
            Expected = "  Expected and actual are both <VBA.Collection> with 3 elements" & vbCrLf & _
                       "  Values differ at index [2]" & vbCrLf & _
                       "    Expected and actual are both <VBA.Collection> with 4 elements" & vbCrLf & _
                       "    Values differ at index [3]" & vbCrLf & _
                       "  Expected: 5" & vbCrLf & _
                       "  But was : 3" & vbCrLf
            Test.Use NewCollection(1, NewCollection(1, 2, 5, 6), 3), NewCollection(1, NewCollection(1, 2, 3, 4), 3), Expected
            
            Expected = "  Expected is <VBA.Collection> with 3 elements, actual is <VBA.Collection> with 2 elements" & vbCrLf & _
                       "  Values differ at index [3]" & vbCrLf & _
                       "  Missing: < 3 >" & vbCrLf
            Test.Use NewCollection(1, 2, 3), NewCollection(1, 2), Expected
            
            Expected = "  Expected is <VBA.Collection> with 2 elements, actual is <VBA.Collection> with 3 elements" & vbCrLf & _
                       "  Values differ at index [3]" & vbCrLf & _
                       "  Extra: < 3 >" & vbCrLf
            Test.Use NewCollection(1, 2), NewCollection(1, 2, 3), Expected
            
            ReDim Array1(0 To 1, 0 To 2) As Variant
            ReDim Array2(0 To 1, 0 To 2) As Variant
            Array1(1, 2) = Array(1, 2, 3)
            Array2(1, 2) = Array(1, 3, 3)
            Expected = "  Expected and actual are both <Variant(0 To 1, 0 To 2)>" & vbCrLf & _
                       "  Values differ at index [1,2]" & vbCrLf & _
                       "    Expected and actual are both <Variant(0 To 2)>" & vbCrLf & _
                       "    Values differ at index [1]" & vbCrLf & _
                       "  Expected: 2" & vbCrLf & _
                       "  But was : 3" & vbCrLf
            Test.Use Array1, Array2, Expected
            
            Expected = "  Expected is <VBA.Collection> with 3 elements, actual is <Variant(0 To 1)>" & vbCrLf & _
                       "  Values differ at expected index [1], actual index [0]" & vbCrLf & _
                       "  Expected: 1" & vbCrLf & _
                       "  But was : 2" & vbCrLf
            Test.Use NewCollection(1, 2, 3), Array(2, 2), Expected
    End Select
    
End Sub

Public Sub WriteMessageTo_WithValues_WritesExpected(ByRef ConstraintExpected As Variant, ByRef ConstraintActual As Variant, ByVal Expected As String)
    Dim c As IConstraint
    Set c = Sim.NewEqualConstraint(ConstraintExpected).Strict
    c.Matches ConstraintActual
    
    Dim w As New TextMessageWriter
    c.WriteMessageTo w
    
    Dim Actual As String
    Actual = w.ToString
    Assert.That Actual, Iz.EqualTo(Expected).NoClip
End Sub

Public Sub WriteMessageTo_WithMultiDimensionalArraysAsCollection_WritesExpected()
    ReDim Array1(0 To 1, 0 To 2) As Variant
    ReDim Array2(1 To 2, 1 To 3) As Variant
    Array1(1, 2) = Array(1, 2, 3)
    Array2(2, 3) = Array(1, 3, 3)
    Dim c As IConstraint
    Set c = Sim.NewEqualConstraint(Array1).AsCollection
    c.Matches Array2
    Dim w As New TextMessageWriter
    c.WriteMessageTo w
    
    Dim Actual As String
    Actual = w.ToString
    
    Dim Expected As String
    Expected = "  Expected is <Variant(0 To 1, 0 To 2)>, actual is <Variant(1 To 2, 1 To 3)>" & vbCrLf & _
               "  Values differ at expected index [1,2], actual index [2,3]" & vbCrLf & _
               "    Expected and actual are both <Variant(0 To 2)>" & vbCrLf & _
               "    Values differ at index [1]" & vbCrLf & _
               "  Expected: 2" & vbCrLf & _
               "  But was : 3" & vbCrLf
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Resolve_WhenCalled_ReturnsSelf()
    Dim c As IResolvable
    Set c = Sim.NewEqualConstraint(1)
    
    Dim Actual As IConstraint
    Set Actual = c.Resolve
    
    Assert.IsTrue Actual Is c
End Sub




Private Function MakeMissing(Optional ByVal Value As Variant) As Variant
    MakeMissing = Value
End Function



