Attribute VB_Name = "modConstraint"
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
' Module: SimplyVBUnit.modConstraint
'
' Provides helper functions for constraints.
'
Option Explicit

Private mMissingVariant As Variant


Public Property Get MissingVariant() As Variant
    If IsEmpty(mMissingVariant) Then
        InitMissingVariant
    End If
    
    MissingVariant = mMissingVariant
End Property

Public Function CanonicalizePath(ByRef Path As String) As String
    Const DIRECTORY_SEPARATOR As String = "\"
    
    Dim Source() As String
    Source = Split(Path, DIRECTORY_SEPARATOR)
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
    
    CanonicalizePath = Join(Target, DIRECTORY_SEPARATOR)
End Function

Public Function Resolve(ByVal Constraint As IConstraint, ByVal Expression As ConstraintExpression) As IConstraint
    If Expression Is Nothing Then
        Set Resolve = Constraint
    Else
        Set Resolve = Expression.Resolve
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitMissingVariant(Optional ByRef MissingValue As Variant)
    mMissingVariant = MissingValue
End Sub
