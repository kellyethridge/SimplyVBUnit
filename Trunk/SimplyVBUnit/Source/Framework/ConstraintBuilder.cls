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
'The MIT License (MIT)
'Copyright (c) 2012 Kelly Ethridge
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights to
'use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of
'the Software, and to permit persons to whom the Software is furnished to do so,
'subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
'INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
'PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
'FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
'OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
'DEALINGS IN THE SOFTWARE.
'
'
' Module: ConstraintBuilder
'
Option Explicit

Private mConstraints    As New ConstraintStack
Private mOperators      As New OperatorStack
Private mFoundLP        As Boolean


Public Function Resolve() As IConstraint
    ReduceOperators
    Set Resolve = Release
End Function

Public Sub AddConstraint(ByVal Constraint As IConstraint)
    mConstraints.Push Constraint
    ReduceOperators
End Sub

Public Sub AddOperator(ByVal Operator As IOperator)
    If TypeOf Operator Is RightParenOperator Then
        ReduceOperators
    Else
        mOperators.Push Operator
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ReduceOperators()
    mFoundLP = False
    
    Do While HasOperators
        ReduceNextOperator
    Loop
End Sub

Private Function HasOperators() As Boolean
    HasOperators = (mOperators.Count > 0) And Not mFoundLP
End Function

Private Function HasConstraints() As Boolean
    HasConstraints = (mConstraints.Count > 0)
End Function

Private Sub ReduceNextOperator()
    Dim Operator As IOperator
    Set Operator = mOperators.Pop
    
    If TypeOf Operator Is LeftParenOperator Then
        mFoundLP = True
    Else
        Operator.Reduce mConstraints
    End If
End Sub

Private Function Release() As IConstraint
    Dim Result As IConstraint
    
    If HasConstraints Then
        Set Result = mConstraints.Peek
        mConstraints.Clear
        mOperators.Clear
    End If
    
    Set Release = Result
End Function
