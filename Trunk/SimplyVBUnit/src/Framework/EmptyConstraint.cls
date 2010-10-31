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
        Err.Raise ErrorCode.Argument, "EmptyConstraint.Matches", "Actual must be a String or support enumeration."
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
