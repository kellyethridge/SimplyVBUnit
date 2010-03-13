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
    Dim Multi1() As Long
    Dim Multi2() As Long
    Dim Arr1 As Variant
    Dim Arr2 As Variant

    Select Case Test.MethodName
        Case "Matches_WithValues"
            Call Test.Use(True, True, True)
            Call Test.Use(True, False, False)
            Call Test.Use(False, False, True)
            Call Test.Use(False, True, False)
            Call Test.Use(True, "True", False)
            Call Test.Use(True, CInt(True), False)
            Call Test.Use("a", "a", True)
            Call Test.Use("A", "a", False)
            Call Test.Use("a", "A", False)
            Call Test.Use("1", 1, False)
            Call Test.Use("True", True, False)
            Call Test.Use(CInt(1), CInt(1), True)
            Call Test.Use(CLng(1), CInt(1), True)
            Call Test.Use(CByte(1), CInt(1), True)
            Call Test.Use(CDbl(1), CInt(1), True)
            Call Test.Use(CSng(1), CInt(1), True)
            Call Test.Use(CCur(1), CInt(1), True)
            Call Test.Use(CInt(1), CLng(1), True)
            Call Test.Use(CDec(1), CLng(1), True)
            Call Test.Use(1, 2, False)
            Call Test.Use(#1/1/2001#, #1/1/2001#, True)
            Call Test.Use(#1/1/2001 8:00:00 AM#, #1/1/2001 8:00:01 AM#, False)
            Call Test.Use(Nothing, Nothing, True)
            Call Test.Use(Nothing, New FakeTest, False)
            Call Test.Use(New FakeTest, Nothing, False)
            Call Test.Use(Nothing, 1, False)
            Call Test.Use(1, Nothing, False)
            Call Test.Use(1999999999, 2E+09!, False)
            Call Test.Use(2E+09!, 1999999999, False)
            Call Test.Use(12345678901234.5, 12345678901234.5@, True)
            Call Test.Use(0.0001, 0.0001@, True)
            Call Test.Use(#1/1/2001#, "1/1/2001", False)
            Call Test.Use(Empty, Empty, True)
            Call Test.Use(Empty, 1, False)
            Call Test.Use(1, Empty, False)
            Call Test.Use(Null, Null, True)
            Call Test.Use(Null, 1, False)
            Call Test.Use(1, Null, False)
            Call Test.Use(MakeMissing, MakeMissing, True)
            Call Test.Use(MakeMissing, 1, False)
            Call Test.Use(1, MakeMissing, False)
            
            Dim Obj As New FakeTest
            Call Test.Use(Obj, Obj, True)
            Call Test.Use(Array(), Array(), True)
            Call Test.Use(Array(1), Array(0), False)
            Call Test.Use(MakeLongArray(1, 1, 2, 3), MakeLongArray(1, 1, 2, 3), True)
            Call Test.Use(MakeLongArray(0, 1, 2, 3), MakeLongArray(1, 1, 2, 3), False)
            Call Test.Use(Array(1, "2"), Array("1", 2), False)
            Call Test.Use(Array(1), Array(1, 2), False)
            Call Test.Use(Array(Obj), Array(Obj), True)
            Call Test.Use(Array(1, 2, 3), Array(1#, 2#, 3#), True)
            
            ReDim Arr1(0 To 1) As Long
            ReDim Arr2(0 To 1) As Integer
            Call Test.Use(Arr1, Arr2, True)
            
            ReDim Multi1(1, 1)
            ReDim Multi2(1, 1)
            Call Test.Use(Multi1, Multi2, True)
            
            ReDim Multi1(1, 1)
            ReDim Multi2(1, 1)
            Multi1(1, 1) = 1
            Call Test.Use(Multi1, Multi2, False)
            
            ReDim Multi1(1, 1)
            ReDim Multi2(1, 2)
            Call Test.Use(Multi1, Multi2, False)
            
            ReDim Multi1(0 To 1, 0 To 1)
            ReDim Multi2(0 To 1, 0 To 1, 0 To 1)
            Call Test.Use(Multi1, Multi2, False)
            
            Dim Col1 As Collection
            Dim Col2 As Collection
            Set Col1 = New Collection
            Call Col1.Add(1)
            Set Col2 = New Collection
            Call Col2.Add(1)
            Call Test.Use(Col1, Col2, True)
            
            Set Col1 = New Collection
            Set Col2 = New Collection
            Call Col1.Add(1)
            Call Col1.Add(2)
            Call Col2.Add(1)
            Call Test.Use(Col1, Col2, False)
            
            Arr1 = Array(Array(1, 2), Array("a", "b"))
            Arr2 = Array(Array(1, 2), Array("a", "b"))
            Call Test.Use(Arr1, Arr2, True)
            
            Arr1 = Array(Array(1, 2), Array("a", "b"))
            Arr2 = Array(Array(1, 2), Array("a", "c"))
            Call Test.Use(Arr1, Arr2, False)

            Set Col1 = New Collection
            Call Col1.Add(1)
            Arr1 = Array(1)
            Call Test.Use(Arr1, Col1, True)
            Call Test.Use(Col1, Arr1, True)
            
        Case "Matches_CaseInsensitive"
            Call Test.Use("lower", "lower", True)
            Call Test.Use("lower", "LOWER", True)
            Call Test.Use(Array("HELLO", "hello"), Array("hello", "HELLO"), True)
            
            
        Case "Matches_WithTolerances"
            Call Test.Use(1#, 1#, 0, True)
            Call Test.Use(1#, 1.1, 0.11, True)
            Call Test.Use(1#, 1.11, 0.1, False)
            Call Test.Use(100, 101, 1, True)
            Call Test.Use(100, 102, 1, False)
            
        Case "Matches_WithTolerancesOfDays"
            Call Test.Use(#1/1/2001#, #1/2/2001#, 1, True)
            Call Test.Use(#1/2/2001#, #1/1/2001#, 1, True)
            Call Test.Use(#1/1/2001#, #1/3/2001#, 1, False)
            Call Test.Use(#1/3/2001#, #1/1/2001#, 1, False)
            Call Test.Use(#1/1/2001#, #1/2/2001 12:00:01 AM#, 1, False)
            Call Test.Use(#1/2/2001 12:00:01 AM#, #1/1/2001#, 1, False)
            
        Case "Matches_WithTolerancesOfHours"
            Call Test.Use(#1/1/2001 8:00:00 AM#, #1/1/2001 9:00:00 AM#, 1, True)
            Call Test.Use(#1/1/2001 9:00:00 AM#, #1/1/2001 8:00:00 AM#, 1, True)
            Call Test.Use(#1/1/2001 8:00:00 AM#, #1/1/2001 9:00:01 AM#, 1, False)
            Call Test.Use(#1/1/2001 9:00:01 AM#, #1/1/2001 8:00:00 AM#, 1, False)
            
        Case "Matches_WithTolerancesOfMinutes"
            Call Test.Use(#1/1/2001 8:00:00 AM#, #1/1/2001 8:00:00 AM#, 1, True)
            Call Test.Use(#1/1/2001 8:00:00 AM#, #1/1/2001 8:01:00 AM#, 1, True)
            Call Test.Use(#1/1/2001 8:00:00 AM#, #1/1/2001 8:01:01 AM#, 1, False)
            
        Case "Matches_WithTolerancesOfSeconds"
            Call Test.Use(#1/1/2001 8:00:00 AM#, #1/1/2001 8:00:00 AM#, 1, True)
            Call Test.Use(#1/1/2001 8:00:00 AM#, #1/1/2001 8:00:01 AM#, 1, True)
            Call Test.Use(#1/1/2001 8:00:00 AM#, #1/1/2001 8:00:02 AM#, 1, False)
            Call Test.Use(#1/1/100#, #12/31/9999 11:59:59 PM#, 312413759999#, True)
            
        Case "Matches_WithArraysAsCollections"
            Call Test.Use(Array(1, 2, 3), Array(1, 2, 3), True)
            
            ReDim Multi1(0 To 1, 1 To 2)
            ReDim Multi2(2 To 3, 3 To 4)
            Call Test.Use(Multi1, Multi2, True)
        
            ReDim Multi1(0 To 1, 1 To 2)
            ReDim Multi2(2 To 3, 3 To 4)
            Multi1(0, 1) = 1
            Call Test.Use(Multi1, Multi2, False)
        
            ReDim Multi1(0 To 1, 1 To 2)
            ReDim Multi2(0 To 3)
            Call Test.Use(Multi1, Multi2, True)
            
        Case "Matches_WithExternalComparer"
            Call Test.Use(1, 3, True)
            Call Test.Use(1, 2, False)
            Call Test.Use(2, 4, True)
            
        Case "Matches_WithStrict"
            Call Test.Use(CLng(1), CInt(1), False)
            Call Test.Use(CLng(1), CLng(1), True)
            
            ReDim Arr1(0 To 1) As Long
            ReDim Arr2(0 To 1) As Integer
            Call Test.Use(Arr1, Arr2, False)
            
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
    Call c.IgnoreCase
    
    Dim Actual As Boolean
    Actual = c.Matches(ConstraintActual)
    
    Assert.IsTrue Actual = Expected
End Sub

Public Sub Matches_WithTolerances(ByRef ConstraintExpected As Variant, ByRef ConstraintActual As Variant, ByVal Tolerance As Double, ByVal Expected As Boolean)
    Dim c As EqualConstraint
    Set c = Sim.NewEqualConstraint(ConstraintExpected)
    Call c.Within(Tolerance)
    
    Dim Actual As Boolean
    Actual = c.Matches(ConstraintActual)
    
    Assert.IsTrue Actual = Expected
End Sub

Public Sub Matches_WithTolerancesOfDays(ByVal ConstraintExpected As Date, ByVal ConstraintActual As Date, ByVal Tolerance As Long, ByVal Expected As Boolean)
    Dim c As EqualConstraint
    Set c = Sim.NewEqualConstraint(ConstraintExpected)
    Call c.Within(Tolerance).Days
    
    Dim Actual As Boolean
    Actual = c.Matches(ConstraintActual)
    
    Assert.IsTrue Actual = Expected
End Sub

Public Sub Matches_WithTolerancesOfHours(ByVal ConstraintExpected As Date, ByVal ConstraintActual As Date, ByVal Tolerance As Long, ByVal Expected As Boolean)
    Dim c As EqualConstraint
    Set c = Sim.NewEqualConstraint(ConstraintExpected)
    Call c.Within(Tolerance).Hours
    
    Dim Actual As Boolean
    Actual = c.Matches(ConstraintActual)
    
    Assert.IsTrue Actual = Expected
End Sub

Public Sub Matches_WithTolerancesOfMinutes(ByVal ConstraintExpected As Date, ByVal ConstraintActual As Date, ByVal Tolerance As Long, ByVal Expected As Boolean)
    Dim c As EqualConstraint
    Set c = Sim.NewEqualConstraint(ConstraintExpected)
    Call c.Within(Tolerance).Minutes
    
    Dim Actual As Boolean
    Actual = c.Matches(ConstraintActual)
    
    Assert.IsTrue Actual = Expected
End Sub

Public Sub Matches_WithTolerancesOfSeconds(ByVal ConstraintExpected As Date, ByVal ConstraintActual As Date, ByVal Tolerance As Double, ByVal Expected As Boolean)
    Dim c As EqualConstraint
    Set c = Sim.NewEqualConstraint(ConstraintExpected)
    Call c.Within(Tolerance).Seconds
    
    Dim Actual As Boolean
    Actual = c.Matches(ConstraintActual)
    
    Assert.IsTrue Actual = Expected
End Sub

Public Sub Matches_WithArraysAsCollections(ByRef ConstraintExpected As Variant, ByRef ConstraintActual As Variant, ByVal Expected As Boolean)
    Dim c As EqualConstraint
    Set c = Sim.NewEqualConstraint(ConstraintExpected)
    Call c.AsCollection
    
    Dim Actual As Boolean
    Actual = c.Matches(ConstraintActual)
    
    Assert.IsTrue Actual = Expected
End Sub

Public Sub Matches_WithExternalComparer(ByRef ConstraintExpected As Variant, ByRef ConstraintActual As Variant, ByVal Expected As Boolean)
    Dim c As EqualConstraint
    Set c = Sim.NewEqualConstraint(ConstraintExpected)
    Call c.Using(New OddEvenEqualityComparer)
    
    Dim Actual As Boolean
    Actual = c.Matches(ConstraintActual)
    
    Assert.IsTrue Actual = Expected
End Sub

Public Sub Matches_WithStrict(ByRef ConstraintExpected As Variant, ByRef ConstraintActual As Variant, ByVal Expected As Boolean)
    Dim c As EqualConstraint
    Set c = Sim.NewEqualConstraint(ConstraintExpected)
    Call c.Strict

    Dim Actual As Boolean
    Actual = c.Matches(ConstraintActual)

    Assert.IsTrue Actual = Expected
End Sub

Public Sub Resolve_WhenCalled_ReturnsSelf()
    Dim c As IConstraint
    Set c = Sim.NewEqualConstraint(1)
    
    Dim Actual As IConstraint
    Set Actual = c.Resolve
    
    Assert.IsTrue Actual Is c
End Sub

Public Sub WriteMessageTo_WithTrueFalse_WritesExpectedButWasClause()
    Dim c As IConstraint
    Set c = Sim.NewEqualConstraint(True)
    Call c.Matches(False)
    
    Dim w As New TextMessageWriter
    c.WriteMessageTo w
    
    Dim Actual As String
    Actual = w.ToString
    Const Expected As String = "Expected: True" & vbCrLf & _
                               "But was : False" & vbCrLf
    Assert.IsTrue Actual = Expected
End Sub



Private Function MakeLongArray(ByVal LowerBound As Long, ParamArray Args() As Variant) As Long()
    Dim Result() As Long
    ReDim Result(LowerBound To LowerBound + UBound(Args))
    
    Dim i As Long
    For i = 0 To UBound(Args)
        Result(LowerBound + i) = Args(i)
    Next
    
    MakeLongArray = Result
End Function

Private Function MakeMissing(Optional ByVal Value As Variant) As Variant
    MakeMissing = Value
End Function



