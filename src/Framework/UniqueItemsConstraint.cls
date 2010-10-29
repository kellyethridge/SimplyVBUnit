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
Implements IBasicExpression
Implements IUniqueItemsConstraint

Private mExpression As ConstraintExpression
Private mActual     As Variant
Private mStrict     As Boolean


Public Function Matches(ByRef Actual As Variant) As Boolean
    If Not IsEnumerable(Actual) Then _
        Err.Raise ErrorCode.Argument, "UniqueItemsConstraint.Matches", "The actual value must be enumerable."
    
    Matches = IsUnique(Actual)
End Function

Public Function Strict() As IBasicExpression
    mStrict = True
    Set Strict = Me
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
    
    Dim Tallier As New ValueTallier
    Tallier.Strict = mStrict
    
    Dim Element As Variant
    For Each Element In Actual
        If Tallier.IncrementTally(Element) > 1 Then Exit Function
    Next Element
    
    IsUnique = True
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

Private Sub IConstraint_WriteActualValueTo(ByVal Writer As TextMessageWriter)

End Sub

Private Sub IConstraint_WriteDescriptionTo(ByVal Writer As TextMessageWriter)

End Sub

Private Sub IConstraint_WriteMessageTo(ByVal Writer As TextMessageWriter)

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
'   IUniqeItemsConstraint Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IUniqueItemsConstraint_AndAlso() As SimplyVBUnitType.IConstraintExpression
    Set IUniqueItemsConstraint_AndAlso = mExpression.AndAlso
End Function

Private Function IUniqueItemsConstraint_OrElse() As SimplyVBUnitType.IConstraintExpression
    Set IUniqueItemsConstraint_OrElse = mExpression.OrElse
End Function

Private Function IUniqueItemsConstraint_Strict() As SimplyVBUnitType.IBasicExpression
    Set IUniqueItemsConstraint_Strict = Strict
End Function
