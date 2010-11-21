VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SamePathConstraint"
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
' Module: SamePathConstraint
'
Option Explicit
Implements IConstraint
Implements IResolvable
Implements IBasicExpression
Implements IPathConstraint


Private mExpected   As String
Private mActual     As Variant
Private mExpression As ConstraintExpression
Private mIgnoreCase As Boolean


Public Function Matches(ByRef Actual As Variant) As Boolean
    Dim Result As Boolean
    If VarType(Actual) = vbString Then
        Result = EqualStrings(Canonicalize(CStr(Actual)), Canonicalize(mExpected), mIgnoreCase)
    End If
    
    If Not Result Then
        VariantCopyInd mActual, Actual
    End If
    
    Matches = Result
End Function

Public Function RespectCase() As SamePathConstraint
    mIgnoreCase = False
    Set RespectCase = Me
End Function


Friend Sub Init(ByRef Expected As String, ByVal Expression As ConstraintExpression)
    mExpected = Expected
    Set mExpression = Expression
    mIgnoreCase = True
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Canonicalize(ByRef Path As String) As String
    Dim Source() As String
    Source = Split(Path, "\")
    Dim Target() As String
    ReDim Target(0 To UBound(Source))
    
    Dim Index   As Long
    Dim i       As Long
    For i = 0 To UBound(Source)
        Select Case Source(i)
            Case ".."
                If Index > 0 Then
                    Index = Index - 1
                End If
                
            Case "."
                ' skip
                
            Case Else
                Target(Index) = Source(i)
                Index = Index + 1
        End Select
    Next
    
    Index = Index - 1
    If Len(Target(Index)) = 0 Then
        Index = Index - 1
    End If
    
    ReDim Preserve Target(0 To Index)
    
    Canonicalize = Join(Target, "\")
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
    Writer.WriteActualValue mActual
End Sub

Private Sub IConstraint_WriteDescriptionTo(ByVal Writer As TextMessageWriter)
    Writer.WritePredicate "Path matching"
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

