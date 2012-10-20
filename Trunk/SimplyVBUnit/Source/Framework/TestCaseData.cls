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
' Module: TestCaseData
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
