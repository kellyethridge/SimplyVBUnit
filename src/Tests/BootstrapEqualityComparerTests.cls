VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BootstrapEqualityComparerTests"
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
' Module: BootstrapEqualityComparerTests
'
Option Explicit
Implements IBootstrapTestClass

Private mCount As Long

Public Function Run() As Long
    Dim Long1() As Long
    Dim Long2() As Long

    mCount = 0
    
    RunTest Equals_WithDefaultSettings_ReturnsExpected(1&, 1&, True)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(1, 1, True)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(1, 1&, True)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(1&, 1, True)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(1@, 1@, True)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(1@, 1&, True)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(1&, 1@, True)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(1&, 1!, True)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(1!, 1&, True)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(1, 1#, True)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(1#, 1, True)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(1#, 1!, True)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(1!, 1#, True)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(1, CDec(1), True)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(CDec(1), 1, True)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(12345678901234.5, 12345678901234.5@, True)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(1, 0, False)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(0, 1, False)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(1, "1", False)
    RunTest Equals_WithDefaultSettings_ReturnsExpected("1", 1, False)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(0, False, False)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(-1, True, False)
    RunTest Equals_WithDefaultSettings_ReturnsExpected("a", "a", True)
    RunTest Equals_WithDefaultSettings_ReturnsExpected("a", "A", False)
    RunTest Equals_WithDefaultSettings_ReturnsExpected("A", "a", False)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(#1/1/2001#, #1/1/2001#, True)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(#1/1/2001#, #1/1/2001 12:00:01 AM#, False)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(#1:00:00 AM#, #1:00:00 AM#, True)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(#1:00:00 AM#, "1:00:00 AM", False)
    RunTest Equals_WithDefaultSettings_ReturnsExpected("1:00:00 AM", #1:00:00 AM#, False)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(CDbl(0), CDate(0), False)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(True, True, True)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(False, False, True)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(True, False, False)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(False, True, False)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(True, "True", False)
    RunTest Equals_WithDefaultSettings_ReturnsExpected("True", True, False)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(Null, Null, True)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(Empty, Empty, True)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(Null, Empty, False)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(Empty, Null, False)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(Nothing, Nothing, True)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(Nothing, 1, False)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(Nothing, Me, False)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(1, Nothing, False)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(Me, Nothing, False)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(Me, Me, True)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(MakeMissing, 1, False)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(1, MakeMissing, False)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(MakeMissing, MakeMissing, True)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(Array(), Array(), True)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(Array(1), Array(1), True)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(Array(), Array(1), False)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(Array(1), Array(), False)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(Array(1), Array(2), False)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(Array(1, 2), Array(2, 1), False)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(NewCollection(), NewCollection(), True)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(NewCollection(1), NewCollection(1), True)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(NewCollection(), NewCollection(1), False)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(NewCollection(1), NewCollection(), False)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(NewCollection(1), NewCollection(2), False)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(NewCollection(1, 2), NewCollection(2, 1), False)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(Array(), NewCollection(), True)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(Array(1), NewCollection(1), True)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(Array(), NewCollection(1), False)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(Array(1), NewCollection(), False)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(Array(1), NewCollection(2), False)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(Array(1, 2), NewCollection(2, 1), False)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(NewCollection(), Array(), True)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(NewCollection(1), Array(1), True)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(NewCollection(), Array(1), False)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(NewCollection(1), Array(), False)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(NewCollection(1), Array(2), False)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(NewCollection(1, 2), Array(2, 1), False)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(NewLongsLb(1, 1), NewLongsLb(1, 1), True)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(NewLongsLb(0, 1), NewLongs(1, 1), False)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(NewLongsLb(1, 1), Array(1), False)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(Array(1), NewLongsLb(1, 1), False)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(NewLongsLb(1, 1), NewCollection(1), True)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(NewCollection(1), NewLongsLb(1, 1), True)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(Array(1, Array(1, 2)), Array(1, Array(1, 2)), True)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(Array(1, Array(1, 2)), Array(1, Array(1, 3)), False)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(Array(1, 2, 3), NewLongs(1, 2, 3), True)
    ReDim Long1(0 To 1, 0 To 1)
    ReDim Long2(0 To 1, 0 To 1, 0 To 1)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(Long1, Long2, False)
    RunTest Equals_WithDefaultSettings_ReturnsExpected(Long2, Long2, True)
    
    RunTest Equals_WithStrictSettings_ReturnsExpected(0, 0, True)
    RunTest Equals_WithStrictSettings_ReturnsExpected(0&, 0, False)
    RunTest Equals_WithStrictSettings_ReturnsExpected(0, 0&, False)
    RunTest Equals_WithStrictSettings_ReturnsExpected(1!, 1!, True)
    RunTest Equals_WithStrictSettings_ReturnsExpected(1!, 1#, False)
    RunTest Equals_WithStrictSettings_ReturnsExpected(1#, 1!, False)
    RunTest Equals_WithStrictSettings_ReturnsExpected(Array(1, 2), Array(1, 2), True)
    RunTest Equals_WithStrictSettings_ReturnsExpected(Array(1, 2), Array(1, 2&), False)
    
    RunTest Equals_WithIgnoreCaseSettings_ReturnsExpected("a", "a", True)
    RunTest Equals_WithIgnoreCaseSettings_ReturnsExpected("a", "A", True)
    RunTest Equals_WithIgnoreCaseSettings_ReturnsExpected("A", "a", True)
    RunTest Equals_WithIgnoreCaseSettings_ReturnsExpected("a", "b", False)
    RunTest Equals_WithIgnoreCaseSettings_ReturnsExpected(Array("a"), Array("A"), True)
    
    RunTest Equals_WithToleranceSettings_ReturnsExpected(1, 1, 0, True)
    RunTest Equals_WithToleranceSettings_ReturnsExpected(1#, 1.1, 0.11, True)
    RunTest Equals_WithToleranceSettings_ReturnsExpected(1#, 1.1!, 0.11, True)
    RunTest Equals_WithToleranceSettings_ReturnsExpected(1#, CDec(1.1), 0.11, True)
    RunTest Equals_WithToleranceSettings_ReturnsExpected(1#, 1.1@, 0.11, True)
    RunTest Equals_WithToleranceSettings_ReturnsExpected(1@, 1.1, 0.11, True)
    RunTest Equals_WithToleranceSettings_ReturnsExpected(1!, 1.1!, 0.11, True)
    RunTest Equals_WithToleranceSettings_ReturnsExpected(1@, CDec(1.1), 0.11, True)
    RunTest Equals_WithToleranceSettings_ReturnsExpected(1@, 1.1@, 0.11, True)
    RunTest Equals_WithToleranceSettings_ReturnsExpected(CDec(1), 1.1, 0.11, True)
    RunTest Equals_WithToleranceSettings_ReturnsExpected(CDec(1), 1.1!, 0.11, True)
    RunTest Equals_WithToleranceSettings_ReturnsExpected(CDec(1), 1.1@, 0.11, True)
    RunTest Equals_WithToleranceSettings_ReturnsExpected(CDec(1), CDec(1.1), 0.11, True)
    RunTest Equals_WithToleranceSettings_ReturnsExpected(1#, 1.11, 0.1, False)
    RunTest Equals_WithToleranceSettings_ReturnsExpected(1#, 1.11!, 0.1, False)
    RunTest Equals_WithToleranceSettings_ReturnsExpected(1#, CDec(1.11), 0.1, False)
    RunTest Equals_WithToleranceSettings_ReturnsExpected(1#, 1.11@, 0.1, False)
    RunTest Equals_WithToleranceSettings_ReturnsExpected(1@, 1.11, 0.1, False)
    RunTest Equals_WithToleranceSettings_ReturnsExpected(1@, 1.11!, 0.1, False)
    RunTest Equals_WithToleranceSettings_ReturnsExpected(1!, CDec(1.11), 0.1, False)
    RunTest Equals_WithToleranceSettings_ReturnsExpected(1@, 1.11@, 0.1, False)
    RunTest Equals_WithToleranceSettings_ReturnsExpected(CDec(1), 1.11, 0.1, False)
    RunTest Equals_WithToleranceSettings_ReturnsExpected(CDec(1), 1.11!, 0.1, False)
    RunTest Equals_WithToleranceSettings_ReturnsExpected(CDec(1), 1.11@, 0.1, False)
    RunTest Equals_WithToleranceSettings_ReturnsExpected(CDec(1), CDec(1.11), 0.1, False)
    RunTest Equals_WithToleranceSettings_ReturnsExpected(100, 101, 1, True)
    RunTest Equals_WithToleranceSettings_ReturnsExpected(100, 102, 1, False)
    RunTest Equals_WithToleranceSettings_ReturnsExpected(#1/1/2001 8:00:00 AM#, #1/1/2001 8:00:02 AM#, 1, False)
    RunTest Equals_WithToleranceSettings_ReturnsExpected(#1/1/2001 8:00:00 AM#, #1/1/2001 8:00:02 AM#, 2, True)

    RunTest Equals_WithTolerancePercentSettings_ReturnsExpected(1000, 900, 10, True)
    RunTest Equals_WithTolerancePercentSettings_ReturnsExpected(1000, 899, 10, False)
    RunTest Equals_WithTolerancePercentSettings_ReturnsExpected(CDec(1000), CDec(900), 10, True)
    RunTest Equals_WithTolerancePercentSettings_ReturnsExpected(CDec(1000), CDec(899), 10, False)
    RunTest Equals_WithTolerancePercentSettings_ReturnsExpected(1000#, 900#, 10, True)
    RunTest Equals_WithTolerancePercentSettings_ReturnsExpected(1000#, 899#, 10, False)
    RunTest Equals_WithTolerancePercentSettings_ReturnsExpected(1000@, 900@, 10, True)
    RunTest Equals_WithTolerancePercentSettings_ReturnsExpected(1000@, 899@, 10, False)
    RunTest Equals_WithTolerancePercentSettings_ReturnsExpected(0, 0, 10, True)
    RunTest Equals_WithTolerancePercentSettings_ReturnsExpected(0, 1, 10, False)
    RunTest Equals_WithTolerancePercentSettings_ReturnsExpected(CDec(0), CDec(0), 10, True)
    RunTest Equals_WithTolerancePercentSettings_ReturnsExpected(CDec(0), CDec(1), 10, False)
    RunTest Equals_WithTolerancePercentSettings_ReturnsExpected(0#, 0#, 10, True)
    RunTest Equals_WithTolerancePercentSettings_ReturnsExpected(0#, 1#, 10, False)
    RunTest Equals_WithTolerancePercentSettings_ReturnsExpected(0@, 0@, 10, True)
    RunTest Equals_WithTolerancePercentSettings_ReturnsExpected(0@, 1@, 10, False)
            
    RunTest Equals_WithToleranceOfTotalDaysSettings_ReturnsExpected(#1/1/2001#, #1/2/2001#, 1, True)
    RunTest Equals_WithToleranceOfTotalDaysSettings_ReturnsExpected(#1/2/2001#, #1/1/2001#, 1, True)
    RunTest Equals_WithToleranceOfTotalDaysSettings_ReturnsExpected(#1/1/2001#, #1/3/2001#, 1, False)
    RunTest Equals_WithToleranceOfTotalDaysSettings_ReturnsExpected(#1/3/2001#, #1/1/2001#, 1, False)
    RunTest Equals_WithToleranceOfTotalDaysSettings_ReturnsExpected(#1/1/2001#, #1/2/2001 12:00:01 AM#, 1, False)
    RunTest Equals_WithToleranceOfTotalDaysSettings_ReturnsExpected(#1/2/2001 12:00:01 AM#, #1/1/2001#, 1, False)

    RunTest Equals_WithToleranceOfTotalHoursSettings_ReturnsExpected(#1/1/2001 8:00:00 AM#, #1/1/2001 9:00:00 AM#, 1, True)
    RunTest Equals_WithToleranceOfTotalHoursSettings_ReturnsExpected(#1/1/2001 9:00:00 AM#, #1/1/2001 8:00:00 AM#, 1, True)
    RunTest Equals_WithToleranceOfTotalHoursSettings_ReturnsExpected(#1/1/2001 8:00:00 AM#, #1/1/2001 9:00:01 AM#, 1, False)
    RunTest Equals_WithToleranceOfTotalHoursSettings_ReturnsExpected(#1/1/2001 9:00:01 AM#, #1/1/2001 8:00:00 AM#, 1, False)

    RunTest Equals_WithToleranceOfTotalMinutesSettings_ReturnsExpected(#1/1/2001 8:00:00 AM#, #1/1/2001 8:00:00 AM#, 1, True)
    RunTest Equals_WithToleranceOfTotalMinutesSettings_ReturnsExpected(#1/1/2001 8:00:00 AM#, #1/1/2001 8:01:00 AM#, 1, True)
    RunTest Equals_WithToleranceOfTotalMinutesSettings_ReturnsExpected(#1/1/2001 8:00:00 AM#, #1/1/2001 8:01:01 AM#, 1, False)
    RunTest Equals_WithToleranceOfTotalMinutesSettings_ReturnsExpected(#1/1/2001 8:01:01 AM#, #1/1/2001 8:00:00 AM#, 1, False)
    
    RunTest Equals_WithToleranceOfTotalSecondsSettings_ReturnsExpected(#1/1/2001 8:00:00 AM#, #1/1/2001 8:00:00 AM#, 1, True)
    RunTest Equals_WithToleranceOfTotalSecondsSettings_ReturnsExpected(#1/1/2001 8:00:00 AM#, #1/1/2001 8:00:01 AM#, 1, True)
    RunTest Equals_WithToleranceOfTotalSecondsSettings_ReturnsExpected(#1/1/2001 8:00:00 AM#, #1/1/2001 8:00:02 AM#, 1, False)
    RunTest Equals_WithToleranceOfTotalSecondsSettings_ReturnsExpected(#1/1/100#, #12/31/9999 11:59:59 PM#, 312413759999#, True)
    
    RunTest Equals_WithAsCollectionSettings_ReturnsExpected(NewLongsLb(0, 1, 2), NewLongsLb(0, 1, 2), True)
    RunTest Equals_WithAsCollectionSettings_ReturnsExpected(NewLongsLb(2, 1, 2), NewLongsLb(0, 1, 2), True)
    RunTest Equals_WithAsCollectionSettings_ReturnsExpected(NewLongsLb(-2, 1, 2), NewLongsLb(0, 1, 2), True)
    ReDim Long1(0 To 1, 1 To 2)
    ReDim Long2(2 To 3, 3 To 4)
    RunTest Equals_WithAsCollectionSettings_ReturnsExpected(Long1, Long2, True)
    ReDim Long1(0 To 1, 1 To 2)
    ReDim Long2(2 To 3, 3 To 4)
    Long1(0, 1) = 1
    RunTest Equals_WithAsCollectionSettings_ReturnsExpected(Long1, Long2, False)
    ReDim Long1(0 To 1, 1 To 2)
    ReDim Long2(0 To 3)
    RunTest Equals_WithAsCollectionSettings_ReturnsExpected(Long1, Long2, True)

    RunTest Equals_WithExternalComparer_ReturnsExpected(1, 3, True)
    RunTest Equals_WithExternalComparer_ReturnsExpected(1, 2, False)
    RunTest Equals_WithExternalComparer_ReturnsExpected(2, 4, True)

    Run = mCount
End Function

Private Sub RunTest(ByRef v As Variant)
    mCount = mCount + 1
End Sub

Private Function MakeMissing(Optional ByVal Value As Variant) As Variant
    MakeMissing = Value
End Function

Private Function Equals_WithDefaultSettings_ReturnsExpected(ByRef X As Variant, ByRef Y As Variant, ByVal Expected As Boolean)
    Dim Comparer As New EqualityComparer
    
    Dim Actual As Boolean
    Actual = Comparer.Equals(X, Y)
    
    Debug.Assert Actual = Expected
End Function

Private Function Equals_WithStrictSettings_ReturnsExpected(ByRef X As Variant, ByRef Y As Variant, ByVal Expected As Boolean)
    Dim Comparer As New EqualityComparer
    Comparer.Strict = True
    
    Dim Actual As Boolean
    Actual = Comparer.Equals(X, Y)
    
    Debug.Assert Actual = Expected
End Function

Private Function Equals_WithIgnoreCaseSettings_ReturnsExpected(ByRef X As Variant, ByRef Y As Variant, ByVal Expected As Boolean)
    Dim Comparer As New EqualityComparer
    Comparer.IgnoreCase = True
    
    Dim Actual As Boolean
    Actual = Comparer.Equals(X, Y)
    
    Debug.Assert Actual = Expected
End Function

Private Function Equals_WithToleranceSettings_ReturnsExpected(ByRef X As Variant, ByRef Y As Variant, ByVal Amount As Double, ByVal Expected As Boolean)
    Dim Comparer As New EqualityComparer
    Set Comparer.Tolerance = Sim.NewTolerance(Amount)
    
    Dim Actual As Boolean
    Actual = Comparer.Equals(X, Y)
    
    Debug.Assert Actual = Expected
End Function

Private Function Equals_WithTolerancePercentSettings_ReturnsExpected(ByRef X As Variant, ByRef Y As Variant, ByVal Amount As Double, ByVal Expected As Boolean)
    Dim Comparer As New EqualityComparer
    Set Comparer.Tolerance = Sim.NewTolerance(Amount).Percent
    
    Dim Actual As Boolean
    Actual = Comparer.Equals(X, Y)
    
    Debug.Assert Actual = Expected
End Function

Private Function Equals_WithToleranceOfTotalDaysSettings_ReturnsExpected(ByRef X As Variant, ByRef Y As Variant, ByVal Amount As Double, ByVal Expected As Boolean)
    Dim Comparer As New EqualityComparer
    Set Comparer.Tolerance = Sim.NewTolerance(Amount).TotalDays
    
    Dim Actual As Boolean
    Actual = Comparer.Equals(X, Y)
    
    Debug.Assert Actual = Expected
End Function

Private Function Equals_WithToleranceOfTotalHoursSettings_ReturnsExpected(ByRef X As Variant, ByRef Y As Variant, ByVal Amount As Double, ByVal Expected As Boolean)
    Dim Comparer As New EqualityComparer
    Set Comparer.Tolerance = Sim.NewTolerance(Amount).TotalHours
    
    Dim Actual As Boolean
    Actual = Comparer.Equals(X, Y)
    
    Debug.Assert Actual = Expected
End Function

Private Function Equals_WithToleranceOfTotalMinutesSettings_ReturnsExpected(ByRef X As Variant, ByRef Y As Variant, ByVal Amount As Double, ByVal Expected As Boolean)
    Dim Comparer As New EqualityComparer
    Set Comparer.Tolerance = Sim.NewTolerance(Amount).TotalMinutes
    
    Dim Actual As Boolean
    Actual = Comparer.Equals(X, Y)
    
    Debug.Assert Actual = Expected
End Function

Private Function Equals_WithToleranceOfTotalSecondsSettings_ReturnsExpected(ByRef X As Variant, ByRef Y As Variant, ByVal Amount As Double, ByVal Expected As Boolean)
    Dim Comparer As New EqualityComparer
    Set Comparer.Tolerance = Sim.NewTolerance(Amount).TotalSeconds
    
    Dim Actual As Boolean
    Actual = Comparer.Equals(X, Y)
    
    Debug.Assert Actual = Expected
End Function

Private Function Equals_WithAsCollectionSettings_ReturnsExpected(ByRef X As Variant, ByRef Y As Variant, ByVal Expected As Boolean)
    Dim Comparer As New EqualityComparer
    Comparer.AsCollection = True
    
    Dim Actual As Boolean
    Actual = Comparer.Equals(X, Y)
    
    Debug.Assert Actual = Expected
End Function

Private Function Equals_WithExternalComparer_ReturnsExpected(ByRef X As Variant, ByRef Y As Variant, ByVal Expected As Boolean)
    Dim Comparer As New EqualityComparer
    Set Comparer.ExternalComparer = New OddEvenEqualityComparer
    
    Dim Actual As Boolean
    Actual = Comparer.Equals(X, Y)
    
    Debug.Assert Actual = Expected
End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IBootstrapTestClass Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IBootstrapTestClass_Run() As Long
    IBootstrapTestClass_Run = Run
End Function
