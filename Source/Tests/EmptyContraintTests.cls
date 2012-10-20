VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EmptyContraintTests"
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
' Module: EmptyContraintTests
'
Option Explicit

Public Sub Matches_WithEmptyString_ReturnsTrue()
    Dim c As New EmptyConstraint
    
    Dim Actual As Boolean
    Actual = c.Matches("")
    
    Assert.IsTrue Actual
End Sub

Public Sub Matches_WithNonEmptyString_ReturnsFalse()
    Dim c As New EmptyConstraint
    
    Dim Actual As Boolean
    Actual = c.Matches("a")
    
    Assert.IsFalse Actual
End Sub

Public Sub Matches_WithEmptyCollection_ReturnsTrue()
    Dim c As New EmptyConstraint
    
    Dim Actual As Boolean
    Actual = c.Matches(New Collection)
    
    Assert.IsTrue Actual
End Sub

Public Sub Matches_WithNonEmptyCollection_ReturnsFalse()
    Dim c As New EmptyConstraint
    Dim NonEmpty As New Collection
    NonEmpty.Add 1
    
    Dim Actual As Boolean
    Actual = c.Matches(NonEmpty)
    
    Assert.IsFalse Actual
End Sub

Public Sub Matches_WithNullArray_ReturnsTrue()
    Dim c As New EmptyConstraint
    Dim NullArray() As Long
    
    Dim Actual As Boolean
    Actual = c.Matches(NullArray)
    
    Assert.IsTrue Actual
End Sub

Public Sub Matches_WithNonEmptyArray_ReturnsFalse()
    Dim c As New EmptyConstraint
    Dim NonEmpty() As Long
    ReDim NonEmpty(0)
    
    Dim Actual As Boolean
    Actual = c.Matches(NonEmpty)
    
    Assert.IsFalse Actual
End Sub

Public Sub Matches_WithZeroLengthArray_ReturnsTrue()
    Dim c As New EmptyConstraint
    Dim EmptyArray As Variant
    EmptyArray = Array()
    
    Dim Actual As Boolean
    Actual = c.Matches(EmptyArray)
    
    Assert.IsTrue Actual
End Sub

Public Sub Matches_WithEmptyCustomEnumerable_ReturnsTrue()
    Dim c As New EmptyConstraint
    Dim Custom As New CollectionProxy
    
    Dim Actual As Boolean
    Actual = c.Matches(Custom)
    
    Assert.IsTrue Actual
End Sub

Public Sub Matches_WithNonEmptyCustomEnumerable_ReturnsFalse()
    Dim c As New EmptyConstraint
    Dim Custom As New CollectionProxy
    Custom.Add 1
    
    Dim Actual As Boolean
    Actual = c.Matches(Custom)
    
    Assert.IsFalse Actual
End Sub

Public Sub Matched_WithNonStringOrEnumerable_ThrowsException()
    Assert.Throws 5
    
    Dim c As New EmptyConstraint
    c.Matches 1
End Sub
