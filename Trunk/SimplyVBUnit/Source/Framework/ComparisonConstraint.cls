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
    Dim Comparer As IComparer
    Set Comparer = mFactory.Create(mExpected)
    
    Dim Result As Boolean
    Select Case Comparer.Compare(Actual, mExpected)
        Case Is > 0:    Result = mCanBeGreater
        Case Is < 0:    Result = mCanBeLess
        Case Else:      Result = mCanBeEqual
    End Select
        
    If Not Result Then
        VariantCopyInd mActual, Actual
    End If
    
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
