VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EnumVariantEnumeratorTests"
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
' Module: EnumVariantEnumeratorTests
'
Option Explicit

Public Sub MoveNext_WithEmptyCollection_ReturnsFalse()
    Dim e As IEnumerator
    Set e = Sim.NewEnumVariantEnumerator(New Collection)
    
    Dim Actual As Boolean
    Actual = e.MoveNext
    
    Assert.IsTrue Actual = False
End Sub

Public Sub MoveNext_WithCollectionOfOneItem_ReturnsTrue()
    Dim c As New Collection
    c.Add 1
    
    Dim e As IEnumerator
    Set e = Sim.NewEnumVariantEnumerator(c)
    
    Dim Actual As Boolean
    Actual = e.MoveNext
    
    Assert.IsTrue Actual = True
End Sub

Public Sub Current_WithCollectionOfOneItem_ReturnsItem()
    Dim c As New Collection
    c.Add "one"
    
    Dim e As IEnumerator
    Set e = Sim.NewEnumVariantEnumerator(c)
    e.MoveNext
    
    Dim Actual As Variant
    Actual = e.Current
    
    Assert.IsTrue Actual = "one"
End Sub

Public Sub Constructor_WithNonsupportedType_RaisesError()
    On Error GoTo errTrap
    
    Sim.NewEnumVariantEnumerator Me
    On Error GoTo 0
    Assert.Fail "Should raise an error."
    
errTrap:
    Assert.IsTrue Err.Number = ErrorCode.Argument
End Sub

Public Sub MoveNext_WithTestListOfTwoTests_ReturnsTrueTwice()
    Dim List As New TestList
    List.Add New FakeTest
    List.Add New FakeTest
    
    Dim e As IEnumerator
    Set e = Sim.NewEnumVariantEnumerator(List)
    
    Assert.IsTrue e.MoveNext = True
    Assert.IsTrue e.MoveNext = True
    Assert.IsTrue e.MoveNext = False
End Sub
