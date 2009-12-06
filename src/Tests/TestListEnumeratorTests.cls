VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestListEnumeratorTests"
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
' Module: TestListEnumeratorTests
'
Option Explicit

Public Sub MoveNext_WithEmptyList_ReturnsFalse()
    Dim List As New TestList
    Dim Enumerator As TestListEnumerator
    Set Enumerator = Sim.NewTestListEnumerator(List)
    
    Dim Actual As Boolean
    Actual = Enumerator.MoveNext
    
    Assert.IsTrue Actual = False, "MoveNext should return false."
End Sub

Public Sub MoveNext_WithListWithOneItem_ReturnsTrue()
    Dim List As New TestList
    List.Add New FakeTest
    Dim Enumerator As TestListEnumerator
    Set Enumerator = Sim.NewTestListEnumerator(List)

    Dim Actual As Boolean
    Actual = Enumerator.MoveNext

    Assert.IsTrue Actual = True, "MoveNext should return true."
End Sub

Public Sub MoveNext_WithListWithOneItem_SetsIndexToCurrentItem()
    Dim Test As New FakeTest
    Dim List As New TestList
    List.Add Test
    Dim Enumerator As TestListEnumerator
    Set Enumerator = Sim.NewTestListEnumerator(List)

    Call Enumerator.MoveNext

    Dim Actual As ITest
    Set Actual = Enumerator.Current
    Assert.IsTrue Actual Is Test, "Wrong item returned."
End Sub

Public Sub MoveNext_Called2TimesWithOneItemList_ReturnsFalse()
    Dim Test As New FakeTest
    Dim List As New TestList
    List.Add Test
    Dim Enumerator As TestListEnumerator
    Set Enumerator = Sim.NewTestListEnumerator(List)
    Enumerator.MoveNext

    Dim Actual As Boolean
    Actual = Enumerator.MoveNext

    Assert.IsTrue Actual = False, "MoveNext should return false."
End Sub

Public Sub Reset_AfterMoveNext_RestartsEnumerator()
    Dim Test As New FakeTest
    Dim List As New TestList
    List.Add Test
    Dim Enumerator As TestListEnumerator
    Set Enumerator = Sim.NewTestListEnumerator(List)
    Call Enumerator.MoveNext

    Call Enumerator.Reset

    Dim Actual As Boolean
    Actual = Enumerator.MoveNext
    Assert.IsTrue Actual = True, "MoveNext should return true."
End Sub
