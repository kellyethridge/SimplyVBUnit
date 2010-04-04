VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConstraintBuilder"
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
' Module: ConstraintBuilder
'
Option Explicit

Private mConstraints    As New ConstraintStack
Private mOperators      As New OperatorStack


Public Function Resolve() As IConstraint
    ReduceOperators
    Set Resolve = Release
End Function

Public Sub AddConstraint(ByVal Constraint As IConstraint)
    mConstraints.Push Constraint
    ReduceOperators
End Sub

Public Sub AddOperator(ByVal Operator As IOperator)
    mOperators.Push Operator
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ReduceOperators()
    Do While HasOperators
        ResolveNextOperator
    Loop
End Sub

Private Function HasOperators() As Boolean
    HasOperators = (mOperators.Count > 0)
End Function

Private Function HasConstraints() As Boolean
    HasConstraints = (mConstraints.Count > 0)
End Function

Private Sub ResolveNextOperator()
    Dim Operator As IOperator
    Set Operator = mOperators.Pop
    
    Operator.Reduce mConstraints
End Sub

Private Function Release() As IConstraint
    Dim Result As IConstraint
    
    If HasConstraints Then
        Set Result = mConstraints.Pop
    End If
    
    mConstraints.Clear
    mOperators.Clear
    
    Set Release = Result
End Function
