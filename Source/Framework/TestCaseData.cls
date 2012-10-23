VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestCaseData"
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
' Module: SComponent.TestCaseData
'
Option Explicit
Private Const LOWER_BOUND As Long = 0

Private mArguments()    As Variant
Private mErrorInfo      As ErrorInfo
Private mExpected       As IConstraint


Public Property Get Expected() As IConstraint
    Set Expected = mExpected
End Property

Public Property Set Expected(ByVal RHS As IConstraint)
    Set mExpected = RHS
End Property

Public Property Get ErrorInfo() As ErrorInfo
    Set ErrorInfo = mErrorInfo
End Property

Public Property Set ErrorInfo(ByVal RHS As ErrorInfo)
    Set mErrorInfo = RHS
End Property

Public Property Get Arguments() As Variant()
    Arguments = mArguments
End Property

Public Sub CopyArguments(ByRef Arguments() As Variant)
    ResizeArray UBound(Arguments)
    MakeReverseCopy Arguments
End Sub

Public Property Get HasArguments() As Boolean
    On Error GoTo errTrap
    HasArguments = UBound(mArguments) >= 0
errTrap:
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ResizeArray(ByVal UpperBound As Long)
    ReDim mArguments(LOWER_BOUND To UpperBound)
End Sub

Private Sub MakeReverseCopy(ByRef Arguments() As Variant)
    Dim i As Long
    For i = LOWER_BOUND To UBound(Arguments)
        VariantCopyInd mArguments(i), Arguments(UBound(Arguments) - i)
    Next
End Sub
