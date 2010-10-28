VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SameAsConstraintTests"
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
' Module: SameAsConstraintTests
'
Option Explicit

Public Sub Matches_WithSameInstance_ReturnsTrue()
    Dim Instance As New FakeEmpty
    Dim c As SameAsConstraint
    Set c = Sim.NewSameAsConstraint(Instance)
    
    Dim Actual As Boolean
    Actual = c.Matches(Instance)
    
    Assert.IsTrue Actual
End Sub

Public Sub Matches_WithDifferenceInstances_ReturnsFalse()
    Dim Instance1 As New FakeEmpty
    Dim Instance2 As New FakeEmpty
    Dim c As SameAsConstraint
    Set c = Sim.NewSameAsConstraint(Instance1)
    
    Dim Actual As Boolean
    Actual = c.Matches(Instance2)
    
    Assert.IsFalse Actual
End Sub
