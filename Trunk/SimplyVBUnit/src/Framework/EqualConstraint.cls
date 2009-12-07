VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EqualConstraint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Copyright 2009 Kelly Ethridge
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
' Module: EqualConstraint
'
Option Explicit
Implements IConstraint
Implements IBasicExpression
Implements IEqualConstraint
Implements IToleranceUnits


Private mExpected   As Variant
Private mActual     As Variant
Private mComparer   As New EqualityComparer
Private mExpression As ExpressionBuilder


''
' Tells the constraint to perform string comparisons as case-insensitive.
'
' @return Returns a reference to itself to make it easier to chain commands together.
'
Public Function IgnoreCase() As EqualConstraint
    Call mComparer.IgnoreCase
    Set IgnoreCase = Me
End Function

Public Function Within(ByVal Tolerance As Variant) As IToleranceUnits
    mComparer.Within Tolerance
    Set Within = Me
End Function

Public Function AsCollection() As EqualConstraint
    Call mComparer.AsCollection
    Set AsCollection = Me
End Function

Public Function Using(ByVal Comparer As IEqualityComparer) As EqualConstraint
    Call mComparer.Using(Comparer)
    Set Using = Me
End Function

Public Function Strict() As EqualConstraint
    mComparer.Strict
    Set Strict = Me
End Function

''
' Compares the actual value to the expected value, returning the result.
'
' @param Actual The value to be compared to the expected value.
' @return The result of the comparison.
'
Public Function Matches(ByRef Actual As Variant) As Boolean
    Call VariantCopyInd(mActual, Actual)
    Matches = mComparer.EqualVariants(mExpected, mActual)
End Function

''
' Writes a message to a text writer describing how the constraint failed.
'
' @param Writer The text write that the message will be written to.
'
Public Sub WriteMessageTo(ByVal Writer As TextMessageWriter)
    Call Writer.DisplayConstraintDifference(Me)
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef Expected As Variant)
    Call VariantCopyInd(mExpected, Expected)
End Sub

Friend Sub Init2(ByRef Expected As Variant, ByVal Expression As ExpressionBuilder)
    Call Init(Expected)
    Set mExpression = Expression
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Days() As EqualConstraint
    mComparer.ToleranceType = InDays
    Set Days = Me
End Function

Private Function Hours() As EqualConstraint
    mComparer.ToleranceType = InHours
    Set Hours = Me
End Function

Private Function Minutes() As EqualConstraint
    mComparer.ToleranceType = InMinutes
    Set Minutes = Me
End Function

Private Function Seconds() As EqualConstraint
    mComparer.ToleranceType = InSeconds
    Set Seconds = Me
End Function


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

Private Function IConstraint_Resolve() As IConstraint
    If mExpression Is Nothing Then
        Set IConstraint_Resolve = Me
    Else
        Set IConstraint_Resolve = mExpression.Resolve
    End If
End Function

Private Sub IConstraint_WriteActualValueTo(ByVal Writer As TextMessageWriter)
    Call Writer.WriteActualValue(mActual)
End Sub

Private Sub IConstraint_WriteDescriptionTo(ByVal Writer As TextMessageWriter)
    Call Writer.WriteExpectedValue(mExpected)
End Sub

Private Sub IConstraint_WriteMessageTo(ByVal Writer As TextMessageWriter)
    Call WriteMessageTo(Writer)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IEqualConstraint
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IEqualConstraint_AndAlso() As SimplyVBUnitType.IConstraintExpression
    Set IEqualConstraint_AndAlso = mExpression.AndAlso
End Function

Private Function IEqualConstraint_AsCollection() As SimplyVBUnitType.IEqualConstraint
    Set IEqualConstraint_AsCollection = AsCollection
End Function

Private Function IEqualConstraint_IgnoreCase() As SimplyVBUnitType.IEqualConstraint
    Set IEqualConstraint_IgnoreCase = Me.IgnoreCase
End Function

Private Function IEqualConstraint_OrElse() As SimplyVBUnitType.IConstraintExpression
    Set IEqualConstraint_OrElse = mExpression.OrElse
End Function

Private Function IEqualConstraint_Strict() As SimplyVBUnitType.IEqualConstraint
    Set IEqualConstraint_Strict = Strict
End Function

Private Function IEqualConstraint_Using(ByVal Comparer As SimplyVBUnitType.IEqualityComparer) As SimplyVBUnitType.IEqualConstraint
    Set IEqualConstraint_Using = Using(Comparer)
End Function

Private Function IEqualConstraint_Within(ByVal Tolerance As Variant) As SimplyVBUnitType.IToleranceUnits
    Set IEqualConstraint_Within = Within(Tolerance)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IToleranceUnits Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IToleranceUnits_Days() As SimplyVBUnitType.IEqualConstraint
    Set IToleranceUnits_Days = Days
End Function

Private Function IToleranceUnits_Hours() As SimplyVBUnitType.IEqualConstraint
    Set IToleranceUnits_Hours = Hours
End Function

Private Function IToleranceUnits_Minutes() As SimplyVBUnitType.IEqualConstraint
    Set IToleranceUnits_Minutes = Minutes
End Function

Private Function IToleranceUnits_Seconds() As SimplyVBUnitType.IEqualConstraint
    Set IToleranceUnits_Seconds = Seconds
End Function
