VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExpressionBuilder"
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
' Module: ExpressionBuilder
'
Option Explicit
Implements IBasicExpression
Implements IConstraintExpression

Private mBuilder As New ConstraintBuilder


Public Function Resolve() As IConstraint
    Call mBuilder.Resolve
    Set Resolve = mBuilder.FinalConstraint
    Set mBuilder = Nothing
End Function

Public Function OrElse() As IConstraintExpression
    Call mBuilder.AddOperator(New OrOperator)
    Set OrElse = Me
End Function

Public Function AndAlso() As IConstraintExpression
    Call mBuilder.AddOperator(New AndOperator)
    Set AndAlso = Me
End Function


Private Function AddConstraint(ByVal c As IConstraint) As IConstraint
    Call mBuilder.AddConstraint(c)
    Set AddConstraint = c
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IBasicExpression Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IBasicExpression_AndAlso() As SimplyVBUnitType.IConstraintExpression
    Set IBasicExpression_AndAlso = AndAlso
End Function

Private Function IBasicExpression_OrElse() As SimplyVBUnitType.IConstraintExpression
    Set IBasicExpression_OrElse = OrElse
End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IConstraintExpression Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IConstraintExpression_AndAlso() As SimplyVBUnitType.IConstraintExpression
    Set IConstraintExpression_AndAlso = AndAlso
End Function

Private Function IConstraintExpression_EqualTo(Expected As Variant) As SimplyVBUnitType.IEqualConstraint
    Set IConstraintExpression_EqualTo = AddConstraint(Sim.NewEqualConstraint2(Expected, Me))
End Function

Private Function IConstraintExpression_OrElse() As SimplyVBUnitType.IConstraintExpression
    Set IConstraintExpression_OrElse = OrElse
End Function
