VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StackTests"
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
' Module: StackTests
'
Option Explicit

Public Sub Constructor_WhenCalled_IsEmpty()
    Dim s As New Stack
    
    Dim Actual As Long
    Actual = s.Count
    
    Assert.IsTrue Actual = 0
End Sub

Public Sub Push_WithValue_StoresOnStack()
    Dim s As New Stack
    
    s.Push 1
    
    Dim Actual As Long
    Actual = s.Count
    Assert.IsTrue Actual = 1
End Sub

Public Sub Pop_WithOneItemOnStack_ReturnsItem()
    Dim s As New Stack
    s.Push 1
    
    Dim Actual As Long
    Actual = s.Pop
    
    Assert.IsTrue Actual = 1
End Sub

Public Sub Pop_WithOneItemOnStack_RemovesItemFromStack()
    Dim s As New Stack
    s.Push 1
    
    s.Pop
    
    Dim Actual As Long
    Actual = s.Count
    Assert.IsTrue Actual = 0
End Sub

Public Sub Push_WithTwoValues_StoresBothOnStack()
    Dim s As New Stack
    
    s.Push 1
    s.Push 2
    
    Dim Actual As Long
    Actual = s.Count
    Assert.IsTrue Actual = 2
End Sub

Public Sub Pop_WithTwoValues_ReturnsInOppositeOrderOfAddition()
    Dim s As New Stack
    s.Push 1
    s.Push 2
    
    Dim Actual As Long
    Actual = s.Pop
    
    Assert.IsTrue Actual = 2
End Sub

Public Sub Clear_WhenCalled_RemovesItems()
    Dim s As New Stack
    s.Push 1
    s.Push 2
    
    s.Clear
    
    Dim Actual As Long
    Actual = s.Count
    Assert.IsTrue Actual = 0
End Sub

Public Sub Peek_WhenCalled_ReturnsTopItem()
    Dim s As New Stack
    s.Push 1
    
    Dim Actual As Long
    Actual = s.Peek
    
    Assert.IsTrue Actual = 1
End Sub

Public Sub Peek_WhenCalled_LeavesItemOnStack()
    Dim s As New Stack
    s.Push 1
    
    s.Peek
    
    Dim Actual As Long
    Actual = s.Count
    Assert.IsTrue Actual = 1
End Sub












