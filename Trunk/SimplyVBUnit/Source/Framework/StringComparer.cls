VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringComparer"
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
' Module: StringComparer
'
Option Explicit
Implements IComparer

Private mCompareMethod As VbCompareMethod


Public Property Get IgnoreCase() As Boolean
    IgnoreCase = (mCompareMethod = vbTextCompare)
End Property

Public Property Let IgnoreCase(ByVal RHS As Boolean)
    If RHS Then
        mCompareMethod = vbTextCompare
    Else
        mCompareMethod = vbBinaryCompare
    End If
End Property

Public Function Compare(ByRef X As Variant, ByRef Y As Variant) As Long
    Compare = StrComp(X, Y, mCompareMethod)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    IgnoreCase = False
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IComparer Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IComparer_Compare(X As Variant, Y As Variant) As Long
    IComparer_Compare = Compare(X, Y)
End Function
