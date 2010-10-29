VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ValueTallierTests"
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
' Module: ItemTallierTests
'
Option Explicit

Public Sub Constructor_WhenCalled_HasNoTallies()
    Dim Tallier As New ValueTallier
    
    Dim Actual As Long
    Actual = Tallier.TotalTallyCount
    
    Assert.That Actual, Iz.EqualTo(0)
End Sub

Public Sub IncrementTally_WhenCalled_IncrementsTallyForValue()
    Const Value As Long = 1
    Dim Tallier As New ValueTallier
    
    Tallier.IncrementTally Value
    
    Dim Actual As Long
    Actual = Tallier.TotalTallyCount
    Assert.That Actual, Iz.EqualTo(1)
End Sub

Public Sub GetTally_WithNonTalliedValue_ReturnsZero()
    Const Value  As Long = 1
    Dim Tallier As New ValueTallier
    
    Dim Actual As Long
    Actual = Tallier.GetTally(Value)
    
    Assert.That Actual, Iz.EqualTo(0)
End Sub

Public Sub GetTally_WithTalliedValue_ReturnsTallyCount()
    Const Value As Long = 1
    Dim Tallier As New ValueTallier
    Tallier.IncrementTally Value
    Tallier.IncrementTally Value
    
    Dim Actual As Long
    Actual = Tallier.GetTally(Value)
    
    Assert.That Actual, Iz.EqualTo(2)
End Sub

Public Sub IncrementTally_WhenCalled_ReturnsCurrentTally()
    Const Value As Long = 1
    Dim Tallier As New ValueTallier
    
    Dim Actual As Long
    Actual = Tallier.IncrementTally(Value)
    
    Assert.That Actual, Iz.EqualTo(1)
End Sub

Public Sub IncrementTally_WithDifferentValues_DoesNotIncrementOtherValues()
    Const Value1 As Long = 1
    Const Value2 As Integer = 1
    Dim Tallier As New ValueTallier
    Tallier.Strict = True
    Tallier.IncrementTally Value1
    
    Dim Actual As Long
    Actual = Tallier.IncrementTally(Value2)
    
    Assert.That Actual, Iz.EqualTo(1)
End Sub

Public Sub DecrementTally_WithNoTallies_DoesNotDecrementBelowZero()
    Const Value As Long = 1
    Dim Tallier As New ValueTallier
    
    Tallier.DecrementTally Value
    
    Dim Actual As Long
    Actual = Tallier.TotalTallyCount
    Assert.That Actual, Iz.EqualTo(0)
End Sub

Public Sub DecrementTally_WhenCalled_ReturnsCurrentTally()
    Const Value As Long = 1
    Dim Tallier As New ValueTallier
    Tallier.IncrementTally Value
    Tallier.IncrementTally Value
    
    Dim Actual As Long
    Actual = Tallier.DecrementTally(Value)
    
    Assert.That Actual, Iz.EqualTo(1)
End Sub

Public Sub Remove_WithNoTalliesForValue_ReturnsFalse()
    Const Value As Long = 1
    Dim Tallier As New ValueTallier
    
    Dim Actual As Boolean
    Actual = Tallier.Remove(Value)
    
    Assert.IsFalse Actual
End Sub

Public Sub Remove_WithTallyToRemove_ReturnsTrue()
    Const Value As Long = 1
    Dim Tallier As New ValueTallier
    Tallier.IncrementTally Value
    
    Dim Actual As Boolean
    Actual = Tallier.Remove(Value)
    
    Assert.IsTrue Actual
End Sub

Public Sub Remove_WhenCalled_DecrementsValueTally()
    Const Value As Long = 1
    Dim Tallier As New ValueTallier
    Tallier.IncrementTally Value
    Tallier.IncrementTally Value
    
    Tallier.Remove Value
    
    Dim Actual As Long
    Actual = Tallier.GetTally(Value)
    Assert.That Actual, Iz.EqualTo(1)
End Sub

Public Sub IncrementTallies_WhenCalled_IncrementsValueTallies()
    Dim Values As Variant
    Values = Array(1&, 1&, "h", "i", "i", "i")
    Dim Tallier As New ValueTallier
    
    Tallier.IncrementTallies Values
    
    Assert.That Tallier.GetTally(1&), Iz.EqualTo(2)
    Assert.That Tallier.GetTally("h"), Iz.EqualTo(1)
    Assert.That Tallier.GetTally("i"), Iz.EqualTo(3)
End Sub
