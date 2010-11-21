VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderedConstraint"
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
' Module: OrderedConstraint
'
Option Explicit
Implements IResolvable
Implements IConstraint
Implements IBasicExpression
Implements IOrderedConstraint


Private Const ORDER_ASC     As Long = 1
Private Const ORDER_DESC    As Long = -1

Private mExpression As ConstraintExpression
Private mSign       As Long
Private mActual     As Variant
Private mFactory    As New ComparerFactory


Public Function DateInterval(ByVal Interval As DateInterval) As OrderedConstraint
    mFactory.Interval = Interval
    Set DateInterval = Me
End Function

Public Function IgnoreCase() As OrderedConstraint
    mFactory.IgnoreCase = True
    Set IgnoreCase = Me
End Function

Public Function Using(ByVal Comparer As IComparer) As OrderedConstraint
    Set mFactory.ExternalComparer = Comparer
    Set Using = Me
End Function

Public Function Descending() As OrderedConstraint
    mSign = ORDER_DESC
    Set Descending = Me
End Function

Public Function Matches(ByRef Actual As Variant) As Boolean
    Dim Enumerator As IEnumerator
    Set Enumerator = GetEnumerator(Actual)
    
    Dim First       As Variant
    Dim Second      As Variant
    Dim Comparer    As IComparer
    
    If Enumerator.MoveNext Then
        VariantCopyInd First, Enumerator.Current
        
        Set Comparer = mFactory.Create(First)
        
        Do While Enumerator.MoveNext
            VariantCopyInd Second, Enumerator.Current
            
            If (Comparer.Compare(First, Second) * mSign) > 0 Then
                VariantCopyInd mActual, Actual
                Exit Function
            End If
            
            VariantCopyInd First, Second
        Loop
    End If
    
    Matches = True
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mSign = ORDER_ASC
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
    Writer.WriteText "Collection or array ordered"
    
    If mSign = ORDER_DESC Then
        Writer.WriteModifier "descending"
    End If
    
    If mFactory.IgnoreCase Then
        Writer.WriteModifier "ignoring case"
    End If
End Sub

Private Sub IConstraint_WriteMessageTo(ByVal Writer As TextMessageWriter)
    Writer.DisplayConstraintDifferences Me
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IOrderedConstraint Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IOrderedConstraint_AndAlso() As SimplyVBUnitType.IConstraintExpression
    Set IOrderedConstraint_AndAlso = mExpression.AndAlso
End Function

Private Function IOrderedConstraint_DateInterval(ByVal Interval As SimplyVBUnitType.DateInterval) As SimplyVBUnitType.IBasicExpression
    Set IOrderedConstraint_DateInterval = DateInterval(Interval)
End Function

Private Function IOrderedConstraint_Descending() As SimplyVBUnitType.IOrderedConstraint
    Set IOrderedConstraint_Descending = Descending
End Function

Private Function IOrderedConstraint_IgnoreCase() As SimplyVBUnitType.IOrderedConstraint
    Set IOrderedConstraint_IgnoreCase = IgnoreCase
End Function

Private Function IOrderedConstraint_OrElse() As SimplyVBUnitType.IConstraintExpression
    Set IOrderedConstraint_OrElse = mExpression.OrElse
End Function

Private Function IOrderedConstraint_Using(ByVal Comparer As SimplyVBUnitType.IComparer) As SimplyVBUnitType.IOrderedConstraint
    Set IOrderedConstraint_Using = Using(Comparer)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IResolvable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IResolvable_Resolve() As IConstraint
    Set IResolvable_Resolve = Resolve(Me, mExpression)
End Function

