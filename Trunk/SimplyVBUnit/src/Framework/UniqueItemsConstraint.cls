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
' Module: UniqueConstraint
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
    VariantCopyInd mActual, Actual
    mComparer.AsCollection = True
    
    Dim AccountedFor        As New Collection
    Dim ActualValue         As Variant
    Dim AccountedForValue   As Variant
    For Each ActualValue In Actual
        For Each AccountedForValue In AccountedFor
            If mComparer.Equals(ActualValue, AccountedForValue) Then
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
    If mExpression Is Nothing Then
        Set IResolvable_Resolve = Me
    Else
        Set IResolvable_Resolve = mExpression.Resolve
    End If
End Function

