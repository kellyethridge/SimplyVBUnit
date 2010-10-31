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

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SVBUnit.TestCaseDataBuilder)
    Dim Long1() As Long
    Dim Long2() As Long
    Dim Array1  As Variant
    Dim Array2  As Variant
    Dim Col1    As Collection
    Dim Col2    As Collection

    Select Case Test.MethodName
        Case "Matches_WithValues"
            Test.Use True, True, True
            Test.Use True, False, False
            Test.Use False, False, True
            Test.Use False, True, False
            Test.Use True, "True", False
            Test.Use "True", True, False
            Test.Use True, CInt(True), False
            Test.Use CInt(True), True, False
            Test.Use "False", False, False
            Test.Use False, "False", False
            Test.Use False, 0, False
            Test.Use 0, False, False
            Test.Use "a", "a", True
            Test.Use "A", "a", False
            Test.Use "a", "A", False
            Test.Use "1", 1, False
            Test.Use 1&, "1", False
            Test.Use 1&, 1, True
            Test.Use 1&, 1&, True
            Test.Use 1&, 1@, True
            Test.Use 1&, 1!, True
            Test.Use 1&, 1#, True
            Test.Use 1&, CDec(1), True
            Test.Use 1&, CByte(1), True
            Test.Use 1@, "1", False
            Test.Use 1@, 1&, True
            Test.Use 1@, 1!, True
            Test.Use 1@, 1#, True
            Test.Use 1@, CDec(1), True
            Test.Use 1@, 1@, True
            Test.Use 1#, 1&, True
            Test.Use 1#, 1!, True
            Test.Use 1#, 1@, True
            Test.Use 1#, CDec(1), True
            Test.Use 1#, 1#, True
            Test.Use 1#, "1.0", False
            Test.Use CDec(1), 1, True
            Test.Use CDec(1), 1#, True
            Test.Use CDec(1), 1!, True
            Test.Use CDec(1), 1@, True
            Test.Use CDec(1), CDec(1), True
            Test.Use CDec(1), "1", False
            Test.Use 1, 2, False
            Test.Use #1/1/2001#, #1/1/2001#, True
            Test.Use #1/1/2001 8:00:00 AM#, #1/1/2001 8:00:01 AM#, False
            Test.Use Nothing, Nothing, True
            Test.Use Nothing, New FakeTest, False
            Test.Use New FakeTest, Nothing, False
            Test.Use Nothing, 1, False
            Test.Use 1, Nothing, False
            Test.Use 1999999999, 2E+09!, False
            Test.Use 2E+09!, 1999999999, False
            Test.Use 12345678901234.5, 12345678901234.5@, True
            Test.Use 0.0001, 0.0001@, True
            Test.Use 0.0001@, 0.0002@, False
            Test.Use #1/1/2001#, "1/1/2001", False
            Test.Use Empty, Empty, True
            Test.Use Empty, 1, False
            Test.Use 1, Empty, False
            Test.Use Null, Null, True
            Test.Use Null, 1, False
            Test.Use 1, Null, False
            Test.Use MakeMissing, MakeMissing, True
            Test.Use MakeMissing, 1, False
            Test.Use 1, MakeMissing, False
            
            Dim Obj As New FakeTest
            Test.Use Obj, Obj, True
            Test.Use Array(), Array(), True
            Test.Use Array(1), Array(0), False
            Test.Use NewLongsLb(1, 1, 2, 3), NewLongsLb(1, 1, 2, 3), True
            Test.Use NewLongsLb(0, 1, 2, 3), NewLongsLb(1, 1, 2, 3), False
            Test.Use Array(1, "2"), Array("1", 2), False
            Test.Use Array(1), Array(1, 2), False
            Test.Use Array(Obj), Array(Obj), True
            Test.Use Array(1, 2, 3), Array(1#, 2#, 3#), True
            Test.Use NewLongs(1, 2, 3), NewDoubles(1, 2, 3), True
            Test.Use NewLongs(1, 1), NewLongs(1, 1), True
            Test.Use NewLongs(1, 2), NewLongs(1, 3), False
            Test.Use NewLongs(1, 1), NewLongs(1, 1, 1), False
            Test.Use NewLongs(1, 2, 3), NewDoubles(1, 2, 3), True
            
            ReDim Long1(0 To 1, 0 To 1)
            ReDim Long2(0 To 1, 0 To 1, 0 To 1)
            Test.Use Long1, Long2, False
            
            Test.Use NewCollection(1, 2), NewCollection(1, 2), True
            Test.Use NewCollection(1), NewCollection(1, 2), False
            Test.Use NewCollection(1, 2), NewCollection(1), False
            Test.Use Array(NewLongs(1, 2), Array("a", "b")), Array(NewLongs(1, 2), Array("a", "b")), True
            Test.Use Array(NewLongs(1, 2), Array("a", "b")), Array(NewLongs(1, 2), Array("a", "c")), False
            Test.Use NewCollection(1, 2), NewLongs(1, 2), True
            Test.Use NewCollection(1, 2), NewLongs(1), False
            Test.Use NewLongs(1, 2), NewCollection(1, 2), True
            Test.Use NewLongs(1, 2), NewCollection(1), False
            
        Case "Matches_CaseInsensitive"
            Test.Use "lower", "lower", True
            Test.Use "lower", "LOWER", True
            Test.Use Array("HELLO", "hello"), Array("hello", "HELLO"), True
            
        Case "Matches_WithTolerances"
            Test.Use 1#, 1.1, 0.11, True
            Test.Use 1#, 1.1!, 0.11, True
            Test.Use 1#, CDec(1.1), 0.11, True
            Test.Use 1#, 1.1@, 0.11, True
            Test.Use 1@, 1.1, 0.11, True
            Test.Use 1@, 1.1!, 0.11, True
            Test.Use 1@, CDec(1.1), 0.11, True
            Test.Use 1@, 1.1@, 0.11, True
            Test.Use CDec(1), 1.1, 0.11, True
            Test.Use CDec(1), 1.1!, 0.11, True
            Test.Use CDec(1), 1.1@, 0.11, True
            Test.Use CDec(1), CDec(1.1), 0.11, True
            
            Test.Use 1#, 1.11, 0.1, False
            Test.Use 1#, 1.11!, 0.1, False
            Test.Use 1#, CDec(1.11), 0.1, False
            Test.Use 1#, 1.11@, 0.1, False
            Test.Use 1@, 1.11, 0.1, False
            Test.Use 1@, 1.11!, 0.1, False
            Test.Use 1@, CDec(1.11), 0.1, False
            Test.Use 1@, 1.11@, 0.1, False
            Test.Use CDec(1), 1.11, 0.1, False
            Test.Use CDec(1), 1.11!, 0.1, False
            Test.Use CDec(1), 1.11@, 0.1, False
            Test.Use CDec(1), CDec(1.11), 0.1, False
            
            Test.Use 100, 101, 1, True
            Test.Use 100, 102, 1, False
            Test.Use #1/1/2001 8:00:00 AM#, #1/1/2001 8:00:02 AM#, 1, False
            Test.Use #1/1/2001 8:00:00 AM#, #1/1/2001 8:00:02 AM#, 2, True
            
        Case "Matches_WithTolerancesOfTotalDays"
            Test.Use #1/1/2001#, #1/2/2001#, 1, True
            Test.Use #1/2/2001#, #1/1/2001#, 1, True
            Test.Use #1/1/2001#, #1/3/2001#, 1, False
            Test.Use #1/3/2001#, #1/1/2001#, 1, False
            Test.Use #1/1/2001#, #1/2/2001 12:00:01 AM#, 1, False
            Test.Use #1/2/2001 12:00:01 AM#, #1/1/2001#, 1, False
            
        Case "Matches_WithTolerancesOfTotalHours"
            Test.Use #1/1/2001 8:00:00 AM#, #1/1/2001 9:00:00 AM#, 1, True
            Test.Use #1/1/2001 9:00:00 AM#, #1/1/2001 8:00:00 AM#, 1, True
            Test.Use #1/1/2001 8:00:00 AM#, #1/1/2001 9:00:01 AM#, 1, False
            Test.Use #1/1/2001 9:00:01 AM#, #1/1/2001 8:00:00 AM#, 1, False
            
        Case "Matches_WithTolerancesOfTotalMinutes"
            Test.Use #1/1/2001 8:00:00 AM#, #1/1/2001 8:00:00 AM#, 1, True
            Test.Use #1/1/2001 8:00:00 AM#, #1/1/2001 8:01:00 AM#, 1, True
            Test.Use #1/1/2001 8:00:00 AM#, #1/1/2001 8:01:01 AM#, 1, False
            
        Case "Matches_WithTolerancesOfTotalSeconds"
            Test.Use #1/1/2001 8:00:00 AM#, #1/1/2001 8:00:00 AM#, 1, True
            Test.Use #1/1/2001 8:00:00 AM#, #1/1/2001 8:00:01 AM#, 1, True
            Test.Use #1/1/2001 8:00:00 AM#, #1/1/2001 8:00:02 AM#, 1, False
            Test.Use #1/1/100#, #12/31/9999 11:59:59 PM#, 312413759999#, True
            
        Case "Matches_WithArraysAsCollections"
            Test.Use Array(1, 2, 3), Array(1, 2, 3), True
            Test.Use NewLongsLb(-10, 1, 2, 3, 4), NewLongsLb(10, 1, 2, 3, 4), True
            
            ReDim Long1(0 To 1, 1 To 2)
            ReDim Long2(2 To 3, 3 To 4)
            Test.Use Long1, Long2, True
        
            ReDim Long1(0 To 1, 1 To 2)
            ReDim Long2(2 To 3, 3 To 4)
            Long1(0, 1) = 1
            Test.Use Long1, Long2, False
        
            ReDim Long1(0 To 1, 1 To 2)
            ReDim Long2(0 To 3)
            Test.Use Long1, Long2, True
            
        Case "Matches_WithExternalComparer"
            Test.Use 1, 3, True
            Test.Use 1, 2, False
            Test.Use 2, 4, True
            
        Case "Matches_WithStrict"
            Test.Use CLng(1), CInt(1), False
            Test.Use CLng(1), CLng(1), True
            
            ReDim Array1(0 To 1) As Long
            ReDim Array2(0 To 1) As Integer
            Test.Use Array1, Array2, False
        
        Case "Matches_WithToleranceInPercent"
            Test.Use 1000, 900, 10, True
            Test.Use 1000, 899, 10, False
            Test.Use CDec(1000), CDec(900), 10, True
            Test.Use CDec(1000), CDec(899), 10, False
            Test.Use 1000#, 900#, 10, True
            Test.Use 1000#, 899#, 10, False
            Test.Use 1000@, 900@, 10, True
            Test.Use 1000@, 899@, 10, False
            Test.Use 0, 0, 10, True
            Test.Use 0, 1, 10, False
            Test.Use CDec(0), CDec(0), 10, True
            Test.Use CDec(0), CDec(1), 10, False
            Test.Use 0#, 0#, 10, True
            Test.Use 0#, 1#, 10, False
            Test.Use 0@, 0@, 10, True
            Test.Use 0@, 1@, 10, False
        
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
            Expected = "  Expected is <Long(0 To 1)>, actual is <Integer(0 To 1)>" & vbCrLf
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

  
Public Sub Matches_WithValues(ByRef ConstraintExpected As Variant, ByRef ConstraintActual As Variant, ByVal Expected As Boolean)
    Dim c As EqualConstraint
    Set c = Sim.NewEqualConstraint(ConstraintExpected)
    
    Dim Actual As Boolean
    Actual = c.Matches(ConstraintActual)
    
    Assert.IsTrue Actual = Expected
End Sub

Public Sub Matches_CaseInsensitive(ByRef ConstraintExpected As Variant, ByRef ConstraintActual As Variant, ByVal Expected As Boolean)
    Dim c As EqualConstraint
    Set c = Sim.NewEqualConstraint(ConstraintExpected)
    c.IgnoreCase
    
    Dim Actual As Boolean
    Actual = c.Matches(ConstraintActual)
    
    Assert.IsTrue Actual = Expected
End Sub

Public Sub Matches_WithTolerances(ByRef ConstraintExpected As Variant, ByRef ConstraintActual As Variant, ByVal Tolerance As Double, ByVal Expected As Boolean)
    Dim c As EqualConstraint
    Set c = Sim.NewEqualConstraint(ConstraintExpected)
    c.Within Tolerance
    
    Dim Actual As Boolean
    Actual = c.Matches(ConstraintActual)
    
    Assert.IsTrue Actual = Expected
End Sub

Public Sub Matches_WithTolerancesOfTotalDays(ByVal ConstraintExpected As Date, ByVal ConstraintActual As Date, ByVal Tolerance As Long, ByVal Expected As Boolean)
    Dim c As EqualConstraint
    Set c = Sim.NewEqualConstraint(ConstraintExpected)
    c.Within(Tolerance).TotalDays
    
    Dim Actual As Boolean
    Actual = c.Matches(ConstraintActual)
    
    Assert.IsTrue Actual = Expected
End Sub

Public Sub Matches_WithTolerancesOfTotalHours(ByVal ConstraintExpected As Date, ByVal ConstraintActual As Date, ByVal Tolerance As Long, ByVal Expected As Boolean)
    Dim c As EqualConstraint
    Set c = Sim.NewEqualConstraint(ConstraintExpected)
    c.Within(Tolerance).TotalHours
    
    Dim Actual As Boolean
    Actual = c.Matches(ConstraintActual)
    
    Assert.IsTrue Actual = Expected
End Sub

Public Sub Matches_WithTolerancesOfTotalMinutes(ByVal ConstraintExpected As Date, ByVal ConstraintActual As Date, ByVal Tolerance As Long, ByVal Expected As Boolean)
    Dim c As EqualConstraint
    Set c = Sim.NewEqualConstraint(ConstraintExpected)
    c.Within(Tolerance).TotalMinutes
    
    Dim Actual As Boolean
    Actual = c.Matches(ConstraintActual)
    
    Assert.IsTrue Actual = Expected
End Sub

Public Sub Matches_WithTolerancesOfTotalSeconds(ByVal ConstraintExpected As Date, ByVal ConstraintActual As Date, ByVal Tolerance As Double, ByVal Expected As Boolean)
    Dim c As EqualConstraint
    Set c = Sim.NewEqualConstraint(ConstraintExpected)
    c.Within(Tolerance).TotalSeconds
    
    Dim Actual As Boolean
    Actual = c.Matches(ConstraintActual)
    
    Assert.IsTrue Actual = Expected
End Sub

Public Sub Matches_WithArraysAsCollections(ByRef ConstraintExpected As Variant, ByRef ConstraintActual As Variant, ByVal Expected As Boolean)
    Dim c As EqualConstraint
    Set c = Sim.NewEqualConstraint(ConstraintExpected)
    c.AsCollection
    
    Dim Actual As Boolean
    Actual = c.Matches(ConstraintActual)
    
    Assert.IsTrue Actual = Expected
End Sub

Public Sub Matches_WithExternalComparer(ByRef ConstraintExpected As Variant, ByRef ConstraintActual As Variant, ByVal Expected As Boolean)
    Dim c As EqualConstraint
    Set c = Sim.NewEqualConstraint(ConstraintExpected)
    c.Using New OddEvenEqualityComparer
    
    Dim Actual As Boolean
    Actual = c.Matches(ConstraintActual)
    
    Assert.IsTrue Actual = Expected
End Sub

Public Sub Matches_WithStrict(ByRef ConstraintExpected As Variant, ByRef ConstraintActual As Variant, ByVal Expected As Boolean)
    Dim c As EqualConstraint
    Set c = Sim.NewEqualConstraint(ConstraintExpected)
    c.Strict

    Dim Actual As Boolean
    Actual = c.Matches(ConstraintActual)

    Assert.IsTrue Actual = Expected
End Sub

Public Sub Matches_WithToleranceInPercent(ByRef ConstraintExpected As Variant, ByRef ConstraintActual As Variant, ByRef Tolerance As Variant, ByVal Expected As Boolean)
    Dim c As EqualConstraint
    Set c = Sim.NewEqualConstraint(ConstraintExpected).Within(Tolerance).Percent
    
    Dim Actual As Boolean
    Actual = c.Matches(ConstraintActual)
    
    Assert.IsTrue Actual = Expected
End Sub

Public Sub WriteMessageTo_WithValues_WritesExpected(ByRef ConstraintExpected As Variant, ByRef ConstraintActual As Variant, ByVal Expected As String)
    Dim c As IConstraint
    Set c = Sim.NewEqualConstraint(ConstraintExpected)
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
    Dim c As EqualConstraint
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



