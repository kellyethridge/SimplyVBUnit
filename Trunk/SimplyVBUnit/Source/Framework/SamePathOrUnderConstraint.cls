VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SamePathOrUnderConstraint"
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
' Module: SComponent.SamePathOrUnderConstraint
'
Option Explicit
Implements IResolvable
Implements IConstraint
Implements IPathConstraint
Implements IBasicExpression


Private mExpression As ConstraintExpression
Private mExpected   As String
Private mActual     As Variant
Private mIgnoreCase As Boolean



Public Function RespectCase() As SamePathOrUnderConstraint
    mIgnoreCase = False
    Set RespectCase = Me
End Function

Public Function Matches(ByRef Actual As Variant) As Boolean
    Dim Result As Boolean
    If VarType(Actual) = vbString Then
        Dim ActualPath      As String
        Dim ExpectedPath    As String
        ActualPath = CanonicalizePath(CStr(Actual))
        ExpectedPath = CanonicalizePath(mExpected)
        
        Select Case Len(ActualPath)
            Case Is = Len(ExpectedPath)
                Result = EqualStrings(ActualPath, ExpectedPath, mIgnoreCase)
                
            Case Is > Len(ExpectedPath)
                If EqualStrings(Left$(ActualPath, Len(ExpectedPath)), ExpectedPath, mIgnoreCase) Then
                    Result = (Mid$(ActualPath, Len(ExpectedPath) + 1, 1) = "\")
                End If
        End Select
    End If
    
    If Not Result Then
        VariantCopyInd mActual, Actual
    End If
    
    Matches = Result
End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef Expected As String, ByVal Expression As ConstraintExpression)
    mExpected = Expected
    Set mExpression = Expression
    mIgnoreCase = True
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
    Writer.WriteActualValue mActual
End Sub

Private Sub IConstraint_WriteDescriptionTo(ByVal Writer As TextMessageWriter)
    Writer.WritePredicate "Path or under matching"
    Writer.WriteExpectedValue mExpected
End Sub

Private Sub IConstraint_WriteMessageTo(ByVal Writer As TextMessageWriter)
    Writer.DisplayConstraintDifferences Me
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IPathConstraint Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IPathConstraint_AndAlso() As SimplyVBUnitType.IConstraintExpression
    Set IPathConstraint_AndAlso = mExpression.AndAlso
End Function

Private Function IPathConstraint_OrElse() As SimplyVBUnitType.IConstraintExpression
    Set IPathConstraint_OrElse = mExpression.OrElse
End Function

Private Function IPathConstraint_RespectCase() As SimplyVBUnitType.IBasicExpression
    Set IPathConstraint_RespectCase = RespectCase
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IResolvable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IResolvable_Resolve() As IConstraint
    Set IResolvable_Resolve = Resolve(Me, mExpression)
End Function
