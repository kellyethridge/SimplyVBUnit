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
Implements IUsingExpression

Private mActual     As Variant
Private mFrom       As Variant
Private mTo         As Variant
Private mComparer   As IComparer
Private mExpression As ConstraintExpression


''
' Checks that the value "Actual" is within the specified range of values.
'
' @param Actual The value to be checked.
' @return Returns True if the value passes the constraint, False otherwise.
'
Public Function Matches(ByRef Actual As Variant) As Boolean
    Call VariantCopyInd(mActual, Actual)
    
    If mComparer Is Nothing Then
        Set mComparer = DefaultComparer.Instance
    End If
    
    Dim Result As Boolean
    If mComparer.Compare(mActual, mFrom, Nothing) >= 0 Then
        Result = (mComparer.Compare(mActual, mTo, Nothing) <= 0)
    End If
    
    Matches = Result
End Function

Public Function Using(ByVal Comparer As IComparer) As RangeConstraint
    Set mComparer = Comparer
    Set Using = Me
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal FromValue As Variant, ByVal ToValue As Variant, ByVal Expression As ConstraintExpression)
    Call VariantCopyInd(mFrom, FromValue)
    Call VariantCopyInd(mTo, ToValue)
    Set mExpression = Expression
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IConstraint Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IConstraint_Matches(Actual As Variant) As Boolean
    IConstraint_Matches = Matches(Actual)
End Function

Private Sub IConstraint_WriteActualValueTo(ByVal Writer As TextMessageWriter)
    Call Writer.WriteValue(mActual)
End Sub

Private Sub IConstraint_WriteDescriptionTo(ByVal Writer As TextMessageWriter)
    Call Writer.WriteText("in range ({0},{1})", mFrom, mTo)
End Sub

Private Sub IConstraint_WriteMessageTo(ByVal Writer As TextMessageWriter)
    Call Writer.DisplayConstraintDifferences(Me)
End Sub


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


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IUsingExpression Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IUsingExpression_AndAlso() As SimplyVBUnitType.IConstraintExpression
    Set IUsingExpression_AndAlso = mExpression.AndAlso
End Function

Private Function IUsingExpression_OrElse() As SimplyVBUnitType.IConstraintExpression
    Set IUsingExpression_OrElse = mExpression.OrElse
End Function

Private Function IUsingExpression_Using(ByVal Comparer As SimplyVBUnitType.IComparer) As SimplyVBUnitType.IBasicExpression
    Set IUsingExpression_Using = Using(Comparer)
End Function
