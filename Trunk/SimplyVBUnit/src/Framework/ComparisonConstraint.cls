VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ComparisonConstraint"
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
' Module: ComparisonConstraint
'
Option Explicit
Implements IBasicExpression
Implements IComparisonConstraint
Implements IResolvable
Implements IConstraint

Private mExpression     As ConstraintExpression
Private mExpected       As Variant
Private mActual         As Variant
Private mCanBeLess      As Boolean
Private mCanBeGreater   As Boolean
Private mCanBeEqual     As Boolean
Private mPredicate      As String
Private mFactory        As New ComparerFactory


Public Function DateInterval(ByVal Interval As DateInterval) As ComparisonConstraint
    mFactory.Interval = Interval
    Set DateInterval = Me
End Function

Public Function IgnoreCase() As ComparisonConstraint
    mFactory.IgnoreCase = True
    Set IgnoreCase = Me
End Function

Public Function Using(ByVal Comparer As IComparer) As ComparisonConstraint
    Set mFactory.ExternalComparer = Comparer
    Set Using = Me
End Function

Public Function Matches(ByRef Actual As Variant) As Boolean
    VariantCopyInd mActual, Actual
    
    Dim Comparer As IComparer
    Set Comparer = mFactory.Create(mExpected)
    
    Dim Result As Boolean
    Select Case Comparer.Compare(mActual, mExpected)
        Case Is > 0:    Result = mCanBeGreater
        Case Is < 0:    Result = mCanBeLess
        Case Else:      Result = mCanBeEqual
    End Select
    
    Matches = Result
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef Expected As Variant, ByVal CanBeLess As Boolean, ByVal CanBeGreater As Boolean, ByVal CanBeEqual As Boolean, ByRef Predicate As String, ByVal Expression As ConstraintExpression)
    VariantCopyInd mExpected, Expected
    mCanBeLess = CanBeLess
    mCanBeGreater = CanBeGreater
    mCanBeEqual = CanBeEqual
    mPredicate = Predicate
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
'   IComparisonConstraint Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IComparisonConstraint_AndAlso() As SimplyVBUnitType.IConstraintExpression
    Set IComparisonConstraint_AndAlso = mExpression.AndAlso
End Function

Private Function IComparisonConstraint_DateInterval(ByVal Interval As SimplyVBUnitType.DateInterval) As SimplyVBUnitType.IBasicExpression
    Set IComparisonConstraint_DateInterval = DateInterval(Interval)
End Function

Private Function IComparisonConstraint_IgnoreCase() As SimplyVBUnitType.IBasicExpression
    Set IComparisonConstraint_IgnoreCase = IgnoreCase
End Function

Private Function IComparisonConstraint_OrElse() As SimplyVBUnitType.IConstraintExpression
    Set IComparisonConstraint_OrElse = mExpression.OrElse
End Function

Private Function IComparisonConstraint_Using(ByVal Comparer As SimplyVBUnitType.IComparer) As SimplyVBUnitType.IBasicExpression
    Set IComparisonConstraint_Using = Using(Comparer)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IConstraint Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IConstraint_Matches(Actual As Variant) As Boolean
    IConstraint_Matches = Matches(Actual)
End Function

Private Sub IConstraint_WriteActualValueTo(ByVal Writer As TextMessageWriter)
    Writer.WriteValue mActual
End Sub

Private Sub IConstraint_WriteDescriptionTo(ByVal Writer As TextMessageWriter)
    Writer.WritePredicate mPredicate
    Writer.WriteValue mExpected
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
