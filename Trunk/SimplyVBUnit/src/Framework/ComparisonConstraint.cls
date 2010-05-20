VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ComparisonConstraint"
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
' Module: ComparisonConstraint
'
Option Explicit
Implements IResolvable
Implements IConstraint

Private mExpected       As Variant
Private mActual         As Variant
Private mCanBeLess      As Boolean
Private mCanBeGreater   As Boolean
Private mCanBeEqual     As Boolean
Private mPredicate      As String
Private mComparer       As IComparer
Private mCompareMethod  As VbCompareMethod


Public Function IgnoreCase() As ComparisonConstraint
    mCompareMethod = vbTextCompare
    Set IgnoreCase = Me
End Function

Public Function Using(ByVal Comparer As IComparer) As ComparisonConstraint
    Set mComparer = Comparer
    Set Using = Me
End Function


Public Function Matches(ByRef Actual As Variant) As Boolean
    VariantCopyInd mActual, Actual
    
    Dim RetVal As Boolean
    If mComparer Is Nothing Then
        RetVal = NormalCompare
    Else
        RetVal = ComparerCompare
    End If
    
    Matches = RetVal
End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef Expected As Variant, ByVal CanBeLess As Boolean, ByVal CanBeGreater As Boolean, ByVal CanBeEqual As Boolean, ByRef Predicate As String)
    VariantCopyInd mExpected, Expected
    mCanBeLess = CanBeLess
    mCanBeGreater = CanBeGreater
    mCanBeEqual = CanBeEqual
    mPredicate = Predicate
    mCompareMethod = vbBinaryCompare
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ComparerCompare() As Boolean
    Dim RetVal As Boolean
    
    Select Case mComparer.Compare(mExpected, mActual)
        Case Is < 0:    RetVal = mCanBeGreater
        Case Is > 0:    RetVal = mCanBeLess
        Case Else:      RetVal = mCanBeEqual
    End Select
    
    ComparerCompare = RetVal
End Function

Private Function NormalCompare() As Boolean
    Select Case VarType(mExpected)
        Case vbString:  RetVal = CompareStrings
        Case vbDate:    RetVal = CompareDates
        Case Else:      RetVal = CompareNumbers
    End Select
End Function

Private Function CompareDates() As Boolean
    Dim Result As Boolean

    If VarType(mActual) = vbDate Then
        Select Case DateDiff("s", mExpected, mActual)
            Case Is > 0:    Result = mCanBeGreater
            Case Is < 0:    Result = mCanBeLess
            Case Else:      Result = mCanBeEqual
        End Select
    End If
    
    CompareDates = Result
End Function

Private Function CompareStrings() As Boolean
    Dim Result As Boolean

    If VarType(mActual) = vbString Then
        Select Case StrComp(mExpected, mActual, mCompareMethod)
            Case Is > 0:    Result = mCanBeLess
            Case Is < 0:    Result = mCanBeGreater
            Case Else:      Result = mCanBeEqual
        End Select
    End If
    
    CompareStrings = Result
End Function

Private Function CompareNumbers() As Boolean
    Dim Result As Boolean

    If IsNumber(mExpected) And IsNumber(mActual) Then
        If mExpected < mActual Then
            Result = mCanBeLess
        ElseIf mExpected > mActual Then
            Result = mCanBeGreater
        Else
            Result = mCanBeEqual
        End If
    End If
    
    CompareNumbers = Result
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
    Writer.WritePredicate mPredicate
    Writer.WriteValue mExpected
End Sub

Private Sub IConstraint_WriteMessageTo(ByVal Writer As TextMessageWriter)
    Writer.DisplayConstraintDifferences Me
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IResolvable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IResolvable_Resolve() As IConstraint

End Function
