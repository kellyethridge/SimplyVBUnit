VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ResolvableConstraint"
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
' Module: ResolvableConstraint
'
Option Explicit
Implements IResolvable
Implements IBasicExpression
Implements IConstraintExpression
Implements IConstraint

Private mExpression As ConstraintExpression
Private mConstraint As IConstraint


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Constraint As IConstraint, ByVal Expression As ConstraintExpression)
    Set mConstraint = Constraint
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
    IConstraint_Matches = mConstraint.Matches(Actual)
End Function

Private Sub IConstraint_WriteActualValueTo(ByVal Writer As TextMessageWriter)
    mConstraint.WriteActualValueTo Writer
End Sub

Private Sub IConstraint_WriteDescriptionTo(ByVal Writer As TextMessageWriter)
    mConstraint.WriteDescriptionTo Writer
End Sub

Private Sub IConstraint_WriteMessageTo(ByVal Writer As TextMessageWriter)
    mConstraint.WriteMessageTo Writer
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IConstraintExpression Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IConstraintExpression_Empty() As SimplyVBUnitType.IBasicExpression
    Set IConstraintExpression_Empty = mExpression.Empty_
End Function

Private Function IConstraintExpression_EqualTo(Expected As Variant) As SimplyVBUnitType.IEqualityExpression
    Set IConstraintExpression_EqualTo = mExpression.EqualTo(Expected)
End Function

Private Function IConstraintExpression_EquivalentTo(Expected As Variant) As SimplyVBUnitType.IEqualityExpression
    Set IConstraintExpression_EquivalentTo = mExpression.EquivalentTo(Expected)
End Function

Private Function IConstraintExpression_False() As SimplyVBUnitType.IBasicExpression
    Set IConstraintExpression_False = mExpression.False_
End Function

Private Function IConstraintExpression_GreaterThan(Expected As Variant) As SimplyVBUnitType.IBasicExpression
    Set IConstraintExpression_GreaterThan = mExpression.GreaterThan(Expected)
End Function

Private Function IConstraintExpression_GreaterThanOrEqualTo(Expected As Variant) As SimplyVBUnitType.IBasicExpression
    Set IConstraintExpression_GreaterThanOrEqualTo = mExpression.GreaterThanOrEqualTo(Expected)
End Function

Private Function IConstraintExpression_InRange(FromValue As Variant, ToValue As Variant) As IComparisonConstraint
    Set IConstraintExpression_InRange = mExpression.InRange(FromValue, ToValue)
End Function

Private Function IConstraintExpression_LessThan(Expected As Variant) As SimplyVBUnitType.IBasicExpression
    Set IConstraintExpression_LessThan = mExpression.LessThan(Expected)
End Function

Private Function IConstraintExpression_LessThanOrEqualTo(Expected As Variant) As SimplyVBUnitType.IBasicExpression
    Set IConstraintExpression_LessThanOrEqualTo = mExpression.LessThanOrEqualTo(Expected)
End Function

Private Function IConstraintExpression_Missing() As SimplyVBUnitType.IBasicExpression
    Set IConstraintExpression_Missing = mExpression.Missing
End Function

Private Function IConstraintExpression_Not() As SimplyVBUnitType.IConstraintExpression
    Set IConstraintExpression_Not = mExpression.Not_
End Function

Private Function IConstraintExpression_Nothing() As SimplyVBUnitType.IBasicExpression
    Set IConstraintExpression_Nothing = mExpression.Nothing_
End Function

Private Function IConstraintExpression_Ordered() As SimplyVBUnitType.IOrderedConstraint
    Set IConstraintExpression_Ordered = mExpression.Ordered
End Function

Private Function IConstraintExpression_SameAs(ByVal Expected As Object) As SimplyVBUnitType.IBasicExpression
    Set IConstraintExpression_SameAs = mExpression.SameAs(Expected)
End Function

Private Function IConstraintExpression_SamePath(Expected As String) As SimplyVBUnitType.IPathConstraint
    Set IConstraintExpression_SamePath = mExpression.SamePath(Expected)
End Function

Private Function IConstraintExpression_StringContaining(Expected As String) As SimplyVBUnitType.ITextConstraint
    Set IConstraintExpression_StringContaining = mExpression.StringContaining(Expected)
End Function

Private Function IConstraintExpression_StringEnding(Expected As String) As SimplyVBUnitType.ITextConstraint
    Set IConstraintExpression_StringEnding = mExpression.StringEnding(Expected)
End Function

Private Function IConstraintExpression_StringMatching(Pattern As String) As SimplyVBUnitType.ITextConstraint
    Set IConstraintExpression_StringMatching = mExpression.StringMatching(Pattern)
End Function

Private Function IConstraintExpression_StringStarting(Expected As String) As SimplyVBUnitType.ITextConstraint
    Set IConstraintExpression_StringStarting = mExpression.StringStarting(Expected)
End Function

Private Function IConstraintExpression_SubsetOf(Expected As Variant) As SimplyVBUnitType.IEqualityExpression
    Set IConstraintExpression_SubsetOf = mExpression.SubsetOf(Expected)
End Function

Private Function IConstraintExpression_True() As SimplyVBUnitType.IBasicExpression
    Set IConstraintExpression_True = mExpression.True_
End Function

Private Function IConstraintExpression_Unique() As SimplyVBUnitType.IEqualityExpression
    Set IConstraintExpression_Unique = mExpression.Unique
End Function

Private Function IConstraintExpression_VBEmpty() As SimplyVBUnitType.IBasicExpression
    Set IConstraintExpression_VBEmpty = mExpression.vbEmpty
End Function

Private Function IConstraintExpression_VBNull() As SimplyVBUnitType.IBasicExpression
    Set IConstraintExpression_VBNull = mExpression.vbNull
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IResolvable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IResolvable_Resolve() As IConstraint
    Set IResolvable_Resolve = mExpression.Resolve
End Function
