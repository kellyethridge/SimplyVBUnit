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
            Call Test.Using(True, True, True)
            Call Test.Using(True, False, False)
            Call Test.Using(False, False, True)
            Call Test.Using(False, True, False)
            Call Test.Using(True, "True", False)
            Call Test.Using(True, CInt(True), False)
            Call Test.Using("a", "a", True)
            Call Test.Using("A", "a", False)
            Call Test.Using("a", "A", False)
            Call Test.Using("1", 1, False)
            Call Test.Using("True", True, False)
            Call Test.Using(CInt(1), CInt(1), True)
            Call Test.Using(CLng(1), CInt(1), True)
            Call Test.Using(CByte(1), CInt(1), True)
            Call Test.Using(CDbl(1), CInt(1), True)
            Call Test.Using(CSng(1), CInt(1), True)
            Call Test.Using(CCur(1), CInt(1), True)
            Call Test.Using(CInt(1), CLng(1), True)
            Call Test.Using(CDec(1), CLng(1), True)
            Call Test.Using(1, 2, False)
            Call Test.Using(#1/1/2001#, #1/1/2001#, True)
            Call Test.Using(#1/1/2001 8:00:00 AM#, #1/1/2001 8:00:01 AM#, False)
            Call Test.Using(Nothing, Nothing, True)
            Call Test.Using(Nothing, New FakeTest, False)
            Call Test.Using(New FakeTest, Nothing, False)
            Call Test.Using(Nothing, 1, False)
            Call Test.Using(1, Nothing, False)
            Call Test.Using(1999999999, 2E+09!, False)
            Call Test.Using(2E+09!, 1999999999, False)
            Call Test.Using(12345678901234.5, 12345678901234.5@, True)
            Call Test.Using(0.0001, 0.0001@, True)
            Call Test.Using(#1/1/2001#, "1/1/2001", False)
            Call Test.Using(Empty, Empty, True)
            Call Test.Using(Empty, 1, False)
            Call Test.Using(1, Empty, False)
            Call Test.Using(Null, Null, True)
            Call Test.Using(Null, 1, False)
            Call Test.Using(1, Null, False)
            Call Test.Using(MakeMissing, MakeMissing, True)
            Call Test.Using(MakeMissing, 1, False)
            Call Test.Using(1, MakeMissing, False)
            
            Dim Obj As New FakeTest
            Call Test.Using(Obj, Obj, True)
            Call Test.Using(Array(), Array(), True)
            Call Test.Using(Array(1), Array(0), False)
            Call Test.Using(MakeLongArray(1, 1, 2, 3), MakeLongArray(1, 1, 2, 3), True)
            Call Test.Using(MakeLongArray(0, 1, 2, 3), MakeLongArray(1, 1, 2, 3), False)
            Call Test.Using(Array(1, "2"), Array("1", 2), False)
            Call Test.Using(Array(1), Array(1, 2), False)
            Call Test.Using(Array(Obj), Array(Obj), True)
            Call Test.Using(Array(1, 2, 3), Array(1#, 2#, 3#), True)
            
            ReDim Arr1(0 To 1) As Long
            ReDim Arr2(0 To 1) As Integer
            Call Test.Using(Arr1, Arr2, True)
            
            ReDim Multi1(1, 1)
            ReDim Multi2(1, 1)
            Call Test.Using(Multi1, Multi2, True)
            
            ReDim Multi1(1, 1)
            ReDim Multi2(1, 1)
            Multi1(1, 1) = 1
            Call Test.Using(Multi1, Multi2, False)
            
            ReDim Multi1(1, 1)
            ReDim Multi2(1, 2)
            Call Test.Using(Multi1, Multi2, False)
            
            ReDim Multi1(0 To 1, 0 To 1)
            ReDim Multi2(0 To 1, 0 To 1, 0 To 1)
            Call Test.Using(Multi1, Multi2, False)
            
            Dim Col1 As Collection
            Dim Col2 As Collection
            Set Col1 = New Collection
            Col1.Add 1
            Set Col2 = New Collection
            Col2.Add 1
            Call Test.Using(Col1, Col2, True)
            
            Set Col1 = New Collection
            Set Col2 = New Collection
            Col1.Add 1
            Col1.Add 2
            Col2.Add 1
            Call Test.Using(Col1, Col2, False)
            
            Arr1 = Array(Array(1, 2), Array("a", "b"))
            Arr2 = Array(Array(1, 2), Array("a", "b"))
            Call Test.Using(Arr1, Arr2, True)
            
            Arr1 = Array(Array(1, 2), Array("a", "b"))
            Arr2 = Array(Array(1, 2), Array("a", "c"))
            Call Test.Using(Arr1, Arr2, False)

            Set Col1 = New Collection
            Col1.Add 1
            Arr1 = Array(1)
            Call Test.Using(Arr1, Col1, True)
            Call Test.Using(Col1, Arr1, True)
            
        Case "Matches_CaseInsensitive"
            Call Test.Using("lower", "lower", True)
            Call Test.Using("lower", "LOWER", True)
            Call Test.Using(Array("HELLO", "hello"), Array("hello", "HELLO"), True)
            
            
        Case "Matches_WithTolerances"
            Call Test.Using(1#, 1#, 0, True)
            Call Test.Using(1#, 1.1, 0.11, True)
            Call Test.Using(1#, 1.11, 0.1, False)
            Call Test.Using(100, 101, 1, True)
            Call Test.Using(100, 102, 1, False)
            
        Case "Matches_WithTolerancesOfDays"
            Call Test.Using(#1/1/2001#, #1/2/2001#, 1, True)
            Call Test.Using(#1/2/2001#, #1/1/2001#, 1, True)
            Call Test.Using(#1/1/2001#, #1/3/2001#, 1, False)
            Call Test.Using(#1/3/2001#, #1/1/2001#, 1, False)
            Call Test.Using(#1/1/2001#, #1/2/2001 12:00:01 AM#, 1, False)
            Call Test.Using(#1/2/2001 12:00:01 AM#, #1/1/2001#, 1, False)
            
        Case "Matches_WithTolerancesOfHours"
            Call Test.Using(#1/1/2001 8:00:00 AM#, #1/1/2001 9:00:00 AM#, 1, True)
            Call Test.Using(#1/1/2001 9:00:00 AM#, #1/1/2001 8:00:00 AM#, 1, True)
            Call Test.Using(#1/1/2001 8:00:00 AM#, #1/1/2001 9:00:01 AM#, 1, False)
            Call Test.Using(#1/1/2001 9:00:01 AM#, #1/1/2001 8:00:00 AM#, 1, False)
            
        Case "Matches_WithTolerancesOfMinutes"
            Call Test.Using(#1/1/2001 8:00:00 AM#, #1/1/2001 8:00:00 AM#, 1, True)
            Call Test.Using(#1/1/2001 8:00:00 AM#, #1/1/2001 8:01:00 AM#, 1, True)
            Call Test.Using(#1/1/2001 8:00:00 AM#, #1/1/2001 8:01:01 AM#, 1, False)
            
        Case "Matches_WithTolerancesOfSeconds"
            Call Test.Using(#1/1/2001 8:00:00 AM#, #1/1/2001 8:00:00 AM#, 1, True)
            Call Test.Using(#1/1/2001 8:00:00 AM#, #1/1/2001 8:00:01 AM#, 1, True)
            Call Test.Using(#1/1/2001 8:00:00 AM#, #1/1/2001 8:00:02 AM#, 1, False)
            Call Test.Using(#1/1/100#, #12/31/9999 11:59:59 PM#, 312413759999#, True)
            
        Case "Matches_WithArraysAsCollections"
            Call Test.Using(Array(1, 2, 3), Array(1, 2, 3), True)
            
            ReDim Multi1(0 To 1, 1 To 2)
            ReDim Multi2(2 To 3, 3 To 4)
            Call Test.Using(Multi1, Multi2, True)
        
            ReDim Multi1(0 To 1, 1 To 2)
            ReDim Multi2(2 To 3, 3 To 4)
            Multi1(0, 1) = 1
            Call Test.Using(Multi1, Multi2, False)
        
            ReDim Multi1(0 To 1, 1 To 2)
            ReDim Multi2(0 To 3)
            Call Test.Using(Multi1, Multi2, True)
            
        Case "Matches_WithExternalComparer"
            Call Test.Using(1, 3, True)
            Call Test.Using(1, 2, False)
            Call Test.Using(2, 4, True)
            
        Case "Matches_WithStrict"
            Call Test.Using(CLng(1), CInt(1), False)
            Call Test.Using(CLng(1), CLng(1), True)
            
            ReDim Arr1(0 To 1) As Long
            ReDim Arr2(0 To 1) As Integer
            Call Test.Using(Arr1, Arr2, False)
            
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

Public Sub Matches_WithTolerancesOfDays(ByVal ConstraintExpected As Date, ByVal ConstraintActual As Date, ByVal Tolerance As Long, ByVal Expected As Boolean)
    Dim c As EqualConstraint
    Set c = Sim.NewEqualConstraint(ConstraintExpected)
    c.Within(Tolerance).Days
    
    Dim Actual As Boolean
    Actual = c.Matches(ConstraintActual)
    
    Assert.IsTrue Actual = Expected
End Sub

Public Sub Matches_WithTolerancesOfHours(ByVal ConstraintExpected As Date, ByVal ConstraintActual As Date, ByVal Tolerance As Long, ByVal Expected As Boolean)
    Dim c As EqualConstraint
    Set c = Sim.NewEqualConstraint(ConstraintExpected)
    c.Within(Tolerance).Hours
    
    Dim Actual As Boolean
    Actual = c.Matches(ConstraintActual)
    
    Assert.IsTrue Actual = Expected
End Sub

Public Sub Matches_WithTolerancesOfMinutes(ByVal ConstraintExpected As Date, ByVal ConstraintActual As Date, ByVal Tolerance As Long, ByVal Expected As Boolean)
    Dim c As EqualConstraint
    Set c = Sim.NewEqualConstraint(ConstraintExpected)
    c.Within(Tolerance).Minutes
    
    Dim Actual As Boolean
    Actual = c.Matches(ConstraintActual)
    
    Assert.IsTrue Actual = Expected
End Sub

Public Sub Matches_WithTolerancesOfSeconds(ByVal ConstraintExpected As Date, ByVal ConstraintActual As Date, ByVal Tolerance As Double, ByVal Expected As Boolean)
    Dim c As EqualConstraint
    Set c = Sim.NewEqualConstraint(ConstraintExpected)
    c.Within(Tolerance).Seconds
    
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
    c.Matches False
    
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



