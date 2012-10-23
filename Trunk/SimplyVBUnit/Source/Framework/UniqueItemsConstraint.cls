VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UniqueItemsConstraint"
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
' Module: SComponent.UniqueItemsConstraint
'
Option Explicit
Implements IConstraint
Implements IResolvable
Implements IEqualityExpression

Private mExpression As ConstraintExpression
Private mActual     As Variant
Private mComparer   As New EqualityComparer


Public Function Matches(ByRef Actual As Variant) As Boolean
    If Not IsEnumerable(Actual) Then _
        Err.Raise ErrorCode.Argument, "UniqueItemsConstraint.Matches", "The actual value must be enumerable."
    
    Matches = IsUnique(Actual)
End Function

Public Function Strict() As UniqueItemsConstraint
    mComparer.Strict = True
    Set Strict = Me
End Function

Public Function IgnoreCase() As UniqueItemsConstraint
    mComparer.IgnoreCase = True
    Set IgnoreCase = Me
End Function

Public Function Using(ByVal Comparer As IEqualityComparer) As UniqueItemsConstraint
    Set mComparer.ExternalComparer = Comparer
    Set Using = Me
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Expression As ConstraintExpression)
    Set mExpression = Expression
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IsUnique(ByRef Actual As Variant) As Boolean
    mComparer.AsCollection = True
    
    Dim AccountedFor        As New Collection
    Dim ActualValue         As Variant
    Dim AccountedForValue   As Variant
    For Each ActualValue In Actual
        For Each AccountedForValue In AccountedFor
            If mComparer.Equals(ActualValue, AccountedForValue) Then
                VariantCopyInd mActual, Actual
                Exit Function
            End If
        Next
        
        AccountedFor.Add ActualValue
    Next
    
    IsUnique = True
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
    Writer.WritePredicate "all items unique"
    
    If mComparer.IgnoreCase Then
        Writer.WriteModifier "ignoring case"
    End If
End Sub

Private Sub IConstraint_WriteMessageTo(ByVal Writer As TextMessageWriter)
    Writer.DisplayConstraintDifferences Me
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IEqualityExpression Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IEqualityExpression_AndAlso() As SimplyVBUnitType.IConstraintExpression
    Set IEqualityExpression_AndAlso = mExpression.AndAlso
End Function

Private Function IEqualityExpression_AsCollection() As SimplyVBUnitType.IEqualityExpression
    
End Function

Private Function IEqualityExpression_IgnoreCase() As SimplyVBUnitType.IEqualityExpression
    Set IEqualityExpression_IgnoreCase = IgnoreCase
End Function

Private Function IEqualityExpression_NoClip() As SimplyVBUnitType.IEqualityExpression

End Function

Private Function IEqualityExpression_OrElse() As SimplyVBUnitType.IConstraintExpression
    Set IEqualityExpression = mExpression.OrElse
End Function

Private Function IEqualityExpression_Strict() As SimplyVBUnitType.IEqualityExpression
    Set IEqualityExpression_Strict = Strict
End Function

Private Function IEqualityExpression_Using(ByVal Comparer As SimplyVBUnitType.IEqualityComparer) As SimplyVBUnitType.IEqualityExpression
    Set IEqualityExpression_Using = Using(Comparer)
End Function

Private Function IEqualityExpression_Within(ByVal Tolerance As Variant) As SimplyVBUnitType.IToleranceUnits
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IResolvable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IResolvable_Resolve() As IConstraint
    Set IResolvable_Resolve = Resolve(Me, mExpression)
End Function

