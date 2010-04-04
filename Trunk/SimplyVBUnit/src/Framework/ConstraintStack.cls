VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConstraintStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
' Module: ConstraintStack
'
Option Explicit

Private mConstraints As New Stack

Public Sub Push(ByVal Constraint As IConstraint)
    mConstraints.Push Constraint
End Sub

Public Function Pop() As IConstraint
    Set Pop = mConstraints.Pop
End Function

Public Property Get Count() As Long
    Count = mConstraints.Count
End Property

Public Property Get Peek() As IConstraint
    Set Peek = mConstraints.Peek
End Property

Public Sub Clear()
    mConstraints.Clear
End Sub
