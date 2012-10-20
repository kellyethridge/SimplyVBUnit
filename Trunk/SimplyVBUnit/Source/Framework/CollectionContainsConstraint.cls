VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CollectionContainsCstr"
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
' Module: CollectionContainsConstraint
'
Option Explicit
Implements IResolvable
Implements IConstraint
Implements ICollectionContainsConstraint
Implements IBasicExpression

Private mExpression As ConstraintExpression
Private mExpected   As Variant
Private mActual     As Variant
Private mComparer   As New EqualityComparer


Public Function Matches(ByRef Actual As Variant) As Boolean
    If Not IsEnumerable(Actual) Then _
        Error.EnumerationNotSupported Me, "Matches", "Actual"
        
    Dim Result  As Boolean
    Dim Item    As Variant
    For Each Item In Actual
        If mComparer.Equals(mExpected, Item) Then
            Result = True
            Exit For
        End If
    Next
    
    If Not Result Then
        VariantCopyInd mActual, Actual
    End If
    
    Matches = Result
End Function

Public Function IgnoreCase() As CollectionContainsCstr
    mComparer.IgnoreCase = True
    Set IgnoreCase = Me
End Function

Public Function Using(ByVal Comparer As IComparer) As CollectionContainsCstr
    Set mComparer.ExternalComparer = Comparer
    Set Using = Me
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef Expected As Variant, ByVal Expression As ConstraintExpression)
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
'   ICollectionContainsConstraint Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ICollectionContainsConstraint_AndAlso() As SimplyVBUnitType.IConstraintExpression
    Set ICollectionContainsConstraint_AndAlso = mExpression.AndAlso
End Function

Private Function ICollectionContainsConstraint_IgnoreCase() As SimplyVBUnitType.IBasicExpression
    Set ICollectionContainsConstraint_IgnoreCase = IgnoreCase
End Function

Private Function ICollectionContainsConstraint_OrElse() As SimplyVBUnitType.IConstraintExpression
    Set ICollectionContainsConstraint_OrElse = mExpression.OrElse
End Function

Private Function ICollectionContainsConstraint_Using(ByVal Comparer As SimplyVBUnitType.IComparer) As SimplyVBUnitType.IBasicExpression
    Set ICollectionContainsConstraint_Using = Using(Comparer)
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
    Writer.WritePredicate "collection containing"
    Writer.WriteExpectedValue mExpected
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
