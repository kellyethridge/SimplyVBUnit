VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AndOperatorTests"
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
' Module: AndOperatorTests
'
Option Explicit


Public Sub Resolve_WithTwoConstraints_ReturnsAndOperatorOfTwoConstraint()
    Dim LeftSide    As New StubConstraint
    Dim RightSide   As New StubConstraint
    Dim Constraints As New ConstraintStack
    Constraints.Push LeftSide
    Constraints.Push RightSide
    
    Dim Op As New AndOperator
    Op.Reduce Constraints
    
    Dim c As AndConstraint
    Set c = Constraints.Pop
    Assert.IsTrue c.LeftSide Is LeftSide
    Assert.IsTrue c.RightSide Is RightSide
End Sub
