VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EquivalentConstraint"
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
' Module: EquivalentToConstraint
'
Option Explicit
Implements IConstraint
Implements IBasicExpression
Implements IEqualityExpression
Implements IResolvable


Private mExpression As ConstraintExpression
Private mExpected   As Variant
Private mActual     As Variant
Private mComparer   As New EqualityComparer


Public Function Strict() As IEqualityExpression
    mComparer.Strict = True
    Set Strict = Me
End Function

Public Function IgnoreCase() As IEqualityExpression
    mComparer.IgnoreCase = True
    Set IgnoreCase = Me
End Function

Public Function Using(ByVal Comparer As IEqualityComparer) As IEqualityExpression
    Set mComparer.ExternalComparer = Comparer
    Set Using = Me
End Function

Public Function Matches(ByRef Actual As Variant) As Boolean
    If Not IsEnumerable(Actual) Then _
        Error.EnumerationNotSupported Me, "Matches", "Actual"
    
    VariantCopyInd mActual, Actual
    
    Dim Tallier As New ValueTallier
    Tallier.Initialize mExpected, mComparer
    
    Dim Value As Variant
    For Each Value In Actual
        If Not Tallier.TryRemove(Value) Then Exit Function
    Next
    
    Matches = Tallier.IsEmpty
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef Expected As Variant, ByVal Expression As ConstraintExpression)
    If Not IsEnumerable(Expected) Then _
        Error.EnumerationNotSupported Me, "Init", "Expected"
    
    VariantCopyInd mExpected, Expected
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
    Writer.WriteValue mActual
End Sub

Private Sub IConstraint_WriteDescriptionTo(ByVal Writer As TextMessageWriter)
    Writer.WritePredicate "equivalent to"
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
