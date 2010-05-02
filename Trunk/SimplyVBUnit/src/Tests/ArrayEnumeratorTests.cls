VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArrayEnumeratorTests"
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
' Module: ArrayEnumeratorTests
'
Option Explicit

Public Sub Constructor_WithNonArray_RaisesError()
    On Error GoTo errTrap
    
    Call Sim.NewArrayEnumerator(1)
    On Error GoTo 0
    
errTrap:
    Call AssertError(Err, ErrorCode.Argument, "ArrayEnumerator.Init", "Arr must be an Array.")
End Sub

Public Sub MoveNext_WithEmptyArray_ReturnsFalse()
    Dim Arr As Variant
    Arr = Array()
    Dim En As IEnumerator
    Set En = Sim.NewArrayEnumerator(Arr)
    
    Dim Actual As Boolean
    Actual = En.MoveNext
    
    Assert.IsTrue Actual = False
End Sub

Public Sub MoveNext_WithOneItem_ReturnsTrue()
    Dim Arr As Variant
    Arr = Array(1)
    Dim En As IEnumerator
    Set En = Sim.NewArrayEnumerator(Arr)
    
    Dim Actual As Boolean
    Actual = En.MoveNext
    
    Assert.IsTrue Actual = True
End Sub

Public Sub Current_WithOneItem_ReturnsItems()
    Dim En As IEnumerator
    Dim Arr As Variant
    Arr = Array(1)
    Set En = Sim.NewArrayEnumerator(Arr)
    En.MoveNext
    
    Dim Actual As Long
    Actual = En.Current
    
    Assert.IsTrue Actual = 1
End Sub

Public Sub MoveNext_CalledTwiceWithOneElement_ReturnsFalse()
    Dim Arr(0 To 1) As Long
    Dim En As IEnumerator
    Set En = Sim.NewArrayEnumerator(Arr)
    En.MoveNext
    
    Dim Actual As Boolean
    Actual = En.MoveNext
    
    Assert.IsTrue Actual = True
End Sub

Public Sub MoveNext_CalledTwiceWith2DimensionArray_ReturnsTrue()
    Dim Arr(0 To 1, 0 To 0) As Long
    Dim En As IEnumerator
    Set En = Sim.NewArrayEnumerator(Arr)
    
    Assert.IsTrue En.MoveNext = True
    Assert.IsTrue En.MoveNext = True
    Assert.IsTrue En.MoveNext = False
End Sub













