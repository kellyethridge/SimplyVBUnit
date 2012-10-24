VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EmptyConstraint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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
' Module: EmptyConstraint
'
Option Explicit
Implements IConstraint

Private mActual As Variant


Public Function Matches(ByRef Actual As Variant) As Boolean
    VariantCopyInd mActual, Actual
    
    Dim Result As Boolean
    If IsArray(Actual) Then
        Result = IsEmptyArray(Actual)
    ElseIf IsCollection(Actual) Then
        Result = IsEmptyCollection(Actual)
    ElseIf IsEnumerable(Actual) Then
        Result = IsEmptyEnumerable(Actual)
    ElseIf VarType(Actual) = vbString Then
        Result = IsEmptyString(Actual)
    Else
        Error.ResourceArgument Me, "Matches", "Actual", Arg_MustBeStringOrEnumerable
    End If
    
    Matches = Result
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IsEmptyArray(ByRef Actual As Variant) As Boolean
    Dim pSA As Long
    pSA = GetArrayPointer(Actual)
    
    Dim Result As Boolean
    Result = True
    
    Dim Rank As Long
    Rank = SafeArrayGetDim(pSA)
    
    Dim i As Long
    For i = 1 To Rank
        If UBound(Actual, i) >= LBound(Actual, i) Then
            Result = False
            Exit For
        End If
    Next i
    
    IsEmptyArray = Result
End Function

Private Function IsEmptyCollection(ByVal Actual As Collection) As Boolean
    IsEmptyCollection = (Actual.Count = 0)
End Function

Private Function IsEmptyEnumerable(ByVal Actual As Object) As Boolean
    Dim Enumerator As IEnumerator
    Set Enumerator = Sim.NewEnumVariantEnumerator(Actual)
    IsEmptyEnumerable = (Enumerator.MoveNext = False)
End Function

Private Function IsEmptyString(ByRef Actual As Variant) As Boolean
    IsEmptyString = (Len(Actual) = 0)
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
    Writer.WriteText "<empty>"
End Sub

Private Sub IConstraint_WriteMessageTo(ByVal Writer As TextMessageWriter)
    Writer.DisplayConstraintDifferences Me
End Sub
