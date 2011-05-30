VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EndsWithConstraint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
' Module: EndsWithConstraint
'
Option Explicit
Implements IConstraint
Implements ITextConstraint
Implements IBasicExpression
Implements IResolvable

Private mExpression As ConstraintExpression
Private mExpected   As String
Private mActual     As Variant
Private mIgnoreCase As Boolean


Public Function Matches(ByRef Actual As Variant) As Boolean
    Dim Result As Boolean
    If VarType(Actual) = vbString Then
        Dim EndOfActual As String
        EndOfActual = Right$(Actual, Len(mExpected))
        
        Result = EqualStrings(EndOfActual, mExpected, mIgnoreCase)
    End If
    
    If Not Result Then
        VariantCopyInd mActual, Actual
    End If
    
    Matches = Result
End Function


Public Function IgnoreCase() As EndsWithConstraint
    mIgnoreCase = True
    Set IgnoreCase = Me
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef Expected As String, ByVal Expression As ConstraintExpression)
    mExpected = Expected
    Set mExpression = Expression
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IBasicExpression Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IBasicExpression_AndAlso() As SimplyVBUnitType.IConstraintExpression
    Set IBasicExpression_AndAlso = mExpression.AndAlso
End Function

Private Function IBasicExpression_OrElse() As SimplyVBUnitType.IConstraintExpression
    Set IBasicExpression_OrElse = mExpression.OrElse
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IConstraint Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IConstraint_Matches(Actual As Variant) As Boolean
    IConstraint_Matches = Matches(Actual)
End Function

Private Sub IConstraint_WriteActualValueTo(ByVal Writer As TextMessageWriter)
    Writer.WriteActualValue mActual
End Sub

Private Sub IConstraint_WriteDescriptionTo(ByVal Writer As TextMessageWriter)
    Writer.WritePredicate "ends with"
    Writer.WriteExpectedValue mExpected
    
    If mIgnoreCase Then
        Writer.WriteModifier "ignoring case"
    End If
End Sub

Private Sub IConstraint_WriteMessageTo(ByVal Writer As TextMessageWriter)
    Writer.DisplayConstraintDifferences Me
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IResolvable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IResolvable_Resolve() As IConstraint
    Set IResolvable_Resolve = Resolve(Me, mExpression)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ITextConstraint Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ITextConstraint_AndAlso() As SimplyVBUnitType.IConstraintExpression
    Set ITextConstraint_AndAlso = mExpression.AndAlso
End Function

Private Function ITextConstraint_IgnoreCase() As SimplyVBUnitType.IBasicExpression
    Set ITextConstraint_IgnoreCase = IgnoreCase
End Function

Private Function ITextConstraint_OrElse() As SimplyVBUnitType.IConstraintExpression
    Set ITextConstraint_OrElse = mExpression.OrElse
End Function
