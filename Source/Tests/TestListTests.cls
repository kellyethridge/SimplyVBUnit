VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestListTests"
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
' Module: TestListTests
'
Option Explicit

Public Sub Constructor_WhenCalled_CreatesEmptyList()
    Dim List As New TestList

    Dim Actual As Long
    Actual = List.Count

    Debug.Assert Actual = 0
End Sub

Public Sub Add_WithOneItem_AddsItemToList()
    Dim List As New TestList
    Dim Test As New FakeTest

    List.Add Test

    Dim Actual As Long
    Actual = List.Count
    Debug.Assert Actual = 1
End Sub

Public Sub Add_WithOneItem_AddsItemToList2()
    Dim List As New TestList
    Dim Test As New FakeTest

    List.Add Test

    Dim Actual As ITest
    Set Actual = List(0)
    Debug.Assert Actual Is Test
End Sub

Public Sub Add_WithNothing_ThrowsException()
    On Error GoTo errTrap
    Dim List As New TestList

    List.Add Nothing
    Debug.Assert False
    
errTrap:
    Debug.Assert Err.Number = ErrorCode.ArgumentNull
End Sub

Public Sub Item_WithIndexTooLow_ThrowsException()
    On Error GoTo errTrap

    Dim List As New TestList
    List.Add New FakeTest

    Const TOO_LOW As Long = -1
    Dim Test As ITest
    Set Test = List(TOO_LOW)
    Debug.Assert False
    
errTrap:
    Debug.Assert Err.Number = ErrorCode.IndexOutOfRange
End Sub

Public Sub Item_WithIndexTooHigh_ThrowsException()
    On Error GoTo errTrap

    Dim List As New TestList
    List.Add New FakeTest

    Dim TOO_HIGH As Long
    TOO_HIGH = 1
    Dim Test As ITest
    Set Test = List(TOO_HIGH)
    Debug.Assert False
    
errTrap:
    Debug.Assert Err.Number = ErrorCode.IndexOutOfRange
End Sub

Public Sub NewEnum_WhenCalled_ReturnsEnumerator()
    Dim List As New TestList

    Dim Actual As IUnknown
    Set Actual = List.NewEnum

    Assert.IsTrue Not Actual Is Nothing, "NewEnum should not return Nothing."
End Sub

Public Sub TestList_WhenCalledByForEach_Enumerates()
    Dim List As New TestList

    Dim Actual As Boolean
    Actual = IsEnumerable(List)

    Assert.IsTrue Actual = True, "List should be enumerable."
End Sub

Public Sub Constructor_WhenCalled_SetsCapacityToDefaultOf16()
    Dim List As New TestList

    Dim Actual As Long
    Actual = List.Capacity

    Const Expected As Long = 16
    Assert.IsTrue Actual = Expected, "Wrong default capacity."
End Sub

Public Sub Add_FillPastCapacity_IncreasesCapacity()
    Dim List As TestList
    Set List = MakeTestListFilledToCapacity
    Dim OriginalCapacity As Long
    OriginalCapacity = List.Capacity

    List.Add New FakeTest

    Dim Actual As Long
    Actual = List.Capacity
    Assert.IsTrue Actual > OriginalCapacity, "List capacity should increase."
End Sub

Public Sub Equals_2EmptyLists_ReturnTrue()
    Dim List1 As New TestList
    Dim List2 As New TestList

    Dim Actual As Boolean
    Actual = List1.Equals(List2)

    Assert.IsTrue Actual = True, "2 empty lists should be equal."
End Sub

Public Sub Equals_2ListsOfSameLengthWithDifferentContents_ReturnsFalse()
    Dim List1 As New TestList
    List1.Add NewStubTest(Name:="1")

    Dim List2 As New TestList
    List2.Add NewStubTest(Name:="2")

    Dim Actual As Boolean
    Actual = List1.Equals(List2)

    Assert.IsTrue Actual = False, "Lists should not be equal."
End Sub

Public Sub Sort_WhenCalled_SortsTestsByNameAscending()
    Dim List As New TestList
    List.Add NewStubTest(Name:="Z")
    List.Add NewStubTest(Name:="D")
    List.Add NewStubTest(Name:="A")

    List.Sort

    Dim Expected As New TestList
    Expected.Add NewStubTest(Name:="A")
    Expected.Add NewStubTest(Name:="D")
    Expected.Add NewStubTest(Name:="Z")

    Assert.IsTrue List.Equals(Expected), "List should be sorted."
End Sub


Private Function NewStubTest(ByVal Name As String) As ITest
    Dim Stub As New StubTest
    Stub.Name = Name
    Set NewStubTest = Stub
End Function

Private Function MakeTestListFilledToCapacity() As TestList
    Dim List As New TestList
    Dim i As Long
    For i = 1 To List.Capacity
        List.Add New FakeTest
    Next i

    Set MakeTestListFilledToCapacity = List
End Function

Private Function IsEnumerable(ByVal List As TestList) As Boolean
    On Error GoTo errTrap
    Dim Item As ITest
    For Each Item In List
        Exit For
    Next
    IsEnumerable = True
    Exit Function

errTrap:
End Function


