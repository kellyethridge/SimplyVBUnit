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

Private Obj As ValueTallier

Public Sub Setup()
    Set Obj = New ValueTallier
End Sub

Public Sub Constructor_WhenCalled_HasNoTallies()
    Dim Tallier As New ValueTallier
    
    Dim Actual As Long
    Actual = Tallier.TotalTallyCount
    
    Assert.That Actual, Iz.EqualTo(0)
End Sub

'Public Sub TestIncrementing()
'    Assert.That Obj.IncrementTally("this"), Iz.EqualTo(1)
'    Assert.That Obj.IncrementTally("that"), Iz.EqualTo(1)
'    Assert.That Obj.IncrementTally("this"), Iz.EqualTo(2)
'
'    Assert.That Obj.GetTally("this"), Iz.EqualTo(2)
'    Assert.That Obj.GetTally("that"), Iz.EqualTo(1)
'    Assert.That Obj.TotalTallyCount, Iz.EqualTo(3)
'End Sub
'
'Public Sub TestDecrementing()
'    Dim i As Long
'    For i = 1 To 10
'        Obj.IncrementTally "this"
'        Obj.IncrementTally "that"
'    Next i
'
'    Assert.That Obj.TotalTallyCount, Iz.EqualTo(20)
'
'    For i = 9 To 0 Step -1
'        Assert.That Obj.DecrementTally("this"), Iz.EqualTo(i)
'    Next i
'
'    Assert.That Obj.GetTally("that"), Iz.EqualTo(10)
'    Assert.That Obj.TotalTallyCount, Iz.EqualTo(10)
'End Sub
'
'Public Sub TestRemoval()
'    Dim i As Long
'    For i = 1 To 10
'        Obj.IncrementTally "this"
'        Obj.IncrementTally "that"
'    Next i
'
'    For i = 1 To 10
'        Assert.That Obj.Remove("this"), Iz.True
'    Next i
'    Assert.That Obj.Remove("this"), Iz.False
'    Assert.That Obj.TotalTallyCount, Iz.EqualTo(10)
'
'    Obj.ClearTally "that"
'    Assert.That Obj.TotalTallyCount, Iz.EqualTo(0)
'End Sub
