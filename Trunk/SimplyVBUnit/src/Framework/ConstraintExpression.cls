VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConstraintExpression"
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
' Module: ConstraintExpression
'
Option Explicit
Implements IBasicExpression
Implements IConstraintExpression

Private mBuilder As New ConstraintBuilder


Public Function Resolve() As IConstraint
    Set Resolve = mBuilder.Resolve
    Set mBuilder = Nothing
End Function

Public Function OrElse() As IConstraintExpression
    Set OrElse = AddOperator(New OrOperator)
End Function

Public Function AndAlso() As IConstraintExpression
    Set AndAlso = AddOperator(New AndOperator)
End Function

Public Function Not_() As IConstraintExpression
    Set Not_ = AddOperator(New NotOperator)
End Function

Public Function EqualTo(ByRef Expected As Variant) As EqualConstraint
    Set EqualTo = AddConstraint(Sim.NewEqualConstraint(Expected, Me))
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AddConstraint(ByVal Constraint As IConstraint) As IConstraint
    mBuilder.AddConstraint Constraint
    Set AddConstraint = Constraint
End Function

Private Function AddOperator(ByVal Operator As IOperator) As IConstraintExpression
    mBuilder.AddOperator Operator
    Set AddOperator = Me
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
Private Function IConstraintExpression_EqualTo(Expected As Variant) As SimplyVBUnitType.IEqualConstraint
    Set IConstraintExpression_EqualTo = EqualTo(Expected)
End Function

Private Function IConstraintExpression_False() As SimplyVBUnitType.IBasicExpression
    Set IConstraintExpression_False = EqualTo(False)
End Function

Private Function IConstraintExpression_Missing() As SimplyVBUnitType.IBasicExpression
    Set IConstraintExpression_Missing = EqualTo(MissingVariant)
End Function

Private Function IConstraintExpression_Not() As SimplyVBUnitType.IConstraintExpression
    Set IConstraintExpression_Not = Not_
End Function

Private Function IConstraintExpression_Nothing() As SimplyVBUnitType.IBasicExpression
    Set IConstraintExpression_Nothing = EqualTo(Nothing)
End Function

Private Function IConstraintExpression_True() As SimplyVBUnitType.IBasicExpression
    Set IConstraintExpression_True = EqualTo(True)
End Function

Private Function IConstraintExpression_VBEmpty() As SimplyVBUnitType.IBasicExpression
    Set IConstraintExpression_VBEmpty = EqualTo(Empty)
End Function

Private Function IConstraintExpression_VBNull() As SimplyVBUnitType.IBasicExpression
    Set IConstraintExpression_VBNull = EqualTo(Null)
End Function
