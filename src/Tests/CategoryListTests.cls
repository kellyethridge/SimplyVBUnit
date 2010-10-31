VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CategoryListTests"
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
' Module: CategoryListTests
'
Option Explicit

Public Sub Constructor_WhenCalled_ListIsEmpty()
    Dim List As New CategoryList
    
    Dim Actual As Long
    Actual = List.Count
    
    Assert.That Actual, Iz.EqualTo(0)
End Sub

Public Sub Add_WithName_StoresNameInList()
    Dim List As New CategoryList
    
    List.Add "Category"
    
    Dim Actual As Long
    Actual = List.Count
    Assert.That Actual, Iz.EqualTo(1)
End Sub

Public Sub Add_WithName_StoresNameInList2()
    Dim List As New CategoryList
    
    List.Add "Category"
    
    Dim Actual As String
    Actual = List(0)
    Assert.That Actual, Iz.EqualTo("Category")
End Sub

Public Sub Contains_WithNonexistingName_ReturnsFalse()
    Dim List As New CategoryList
    List.Add "Nonexisting"
    
    Dim Actual As Boolean
    Actual = List.Contains("Something")
    
    Assert.IsFalse Actual
End Sub

Public Sub Contains_WithExistingName_ReturnsTrue()
    Dim List As New CategoryList
    List.Add "Existing"
    
    Dim Actual As Boolean
    Actual = List.Contains("Existing")
    
    Assert.IsTrue Actual
End Sub

Public Sub Add_WithExistingName_DoesNothing()
    Dim List As New CategoryList
    
    On Error GoTo errTrap
    List.Add "Name"
    List.Add "Name"
    Exit Sub
    
errTrap:
    Assert.Fail "No error should be raised."
End Sub

Public Sub Item_WithNegativeIndex_RaisesError()
    Dim List As New CategoryList
    
    On Error GoTo errTrap
    Dim Name As String
    Name = List(-1)
    
    On Error GoTo 0
    Assert.Fail "An error should be raised."
    
errTrap:
    AssertError Err, ErrorCode.IndexOutOfRange, "CategoryList.Item", "Index was out of range. Must be non-negative and less than the size of the collection."
End Sub

Public Sub Item_WithIndexPassedEndOfList_RaisesError()
    Dim List As New CategoryList
    List.Add "Name"
    
    On Error GoTo errTrap
    Dim Name As String
    Name = List(1)
    
    On Error GoTo 0
    Assert.Fail "An error should be raised."
    
errTrap:
    AssertError Err, ErrorCode.IndexOutOfRange, "CategoryList.Item", "Index was out of range. Must be non-negative and less than the size of the collection."
End Sub









