VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TallyTests"
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
' Module: TallyTests
'
Option Explicit

Public Sub Constructor_WhenCalled_InitializesTally()
    Dim t As New Tally
    
    Dim Actual As Long
    Actual = t.Count
    
    Assert.AreEqual 0, Actual
End Sub

Public Sub Increment_WhenCalled_IncreasesTally()
    Dim t As New Tally
    
    t.Increment
    
    Dim Actual As Long
    Actual = t.Count
    Assert.AreEqual 1, Actual
End Sub

Public Sub Increment_WhenCalled_ReturnsIncrementedTally()
    Dim t As New Tally
    
    Dim Actual As Long
    Actual = t.Increment
    
    Assert.AreEqual 1, Actual
End Sub

Public Sub Decrement_WhenCalled_DecreasesTally()
    Dim t As New Tally
    t.Increment
    
    t.Decrement
    
    Dim Actual As Long
    Actual = t.Count
    Assert.AreEqual 0, Actual
End Sub

Public Sub Decrement_WhenZero_DoesNotDecreaseTally()
    Dim t As New Tally
    
    t.Decrement
    
    Dim Actual As Long
    Actual = t.Count
    Assert.AreEqual 0, Actual
End Sub

Public Sub Decrement_WhenCalled_ReturnsDecrementedTally()
    Dim t As New Tally
    t.Increment
    t.Increment
    
    Dim Actual As Long
    Actual = t.Decrement
    
    Assert.AreEqual 1, Actual
End Sub
