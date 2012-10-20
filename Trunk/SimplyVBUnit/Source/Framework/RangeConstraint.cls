VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RangeConstraint"
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
' Module: RangeConstraint
'
Option Explicit
Implements IResolvable
Implements IConstraint
Implements IComparisonConstraint

Private mExpression As ConstraintExpression
Private mActual     As Variant
Private mFrom       As Variant
Private mTo         As Variant
Private mFactory    As New ComparerFactory



Public Function IgnoreCase() As RangeConstraint
    mFactory.IgnoreCase = True
    Set IgnoreCase = Me
End Function

Public Function DateInterval(ByVal Interval As DateInterval) As RangeConstraint
    mFactory.Interval = Interval
    Set DateInterval = Me
End Function

Public Function Using(ByVal Comparer As IComparer) As RangeConstraint
    Set mFactory.ExternalComparer = Comparer
    Set Using = Me
End Function

''
' Checks that the value "Actual" is within the specified range of values.
'
' @param Actual The value to be checked.
' @return Returns True if the value passes the constraint, False otherwise.
'
Public Function Matches(ByRef Actual As Variant) As Boolean
    Dim Comparer As IComparer
    Set Comparer = mFactory.Create(Actual)
    
    Dim Result As Boolean
    If Comparer.Compare(Actual, mFrom) >= 0 Then
        Result = (Comparer.Compare(Actual, mTo) <= 0)
    End If
    
    If Result Then
        Matches = True
    Else
        VariantCopyInd mActual, Actual
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal FromValue As Variant, ByVal ToValue As Variant, ByVal Expression As ConstraintExpression)
    VariantCopyInd mFrom, FromValue
    VariantCopyInd mTo, ToValue
    Set mExpression = Expression
End Sub


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
    Writer.WriteText "in range ({0},{1})", mFrom, mTo
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

