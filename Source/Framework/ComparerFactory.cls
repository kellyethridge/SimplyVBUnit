VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ComparerFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
' Module: SComponent.ComparerFactory
'
Option Explicit

Private mIgnoreCase         As Boolean
Private mInterval           As DateInterval
Private mExternalComparer   As IComparer

Public Property Get IgnoreCase() As Boolean
    IgnoreCase = mIgnoreCase
End Property

Public Property Let IgnoreCase(ByVal RHS As Boolean)
    mIgnoreCase = RHS
End Property

Public Property Get Interval() As DateInterval
    Interval = mInterval
End Property

Public Property Let Interval(ByVal RHS As DateInterval)
    mInterval = RHS
End Property

Public Property Get ExternalComparer() As IComparer
    Set ExternalComparer = mExternalComparer
End Property

Public Property Set ExternalComparer(ByVal RHS As IComparer)
    Set mExternalComparer = RHS
End Property

Public Function Create(ByRef ForValue As Variant) As IComparer
    Dim Result As IComparer
    
    If Not mExternalComparer Is Nothing Then
        Set Result = ExternalComparer
    ElseIf VarType(ForValue) = vbString Then
        Set Result = CreateStringComparer
    ElseIf VarType(ForValue) = vbDate Then
        Set Result = CreateDateComparer
    Else
        Set Result = New NumberComparer
    End If
    
    Set Create = Result
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CreateStringComparer() As IComparer
    Dim Comparer As New StringComparer
    Comparer.IgnoreCase = mIgnoreCase
    Set CreateStringComparer = Comparer
End Function

Private Function CreateDateComparer() As IComparer
    Dim Comparer As New DateComparer
    Comparer.Interval = mInterval
    Set CreateDateComparer = Comparer
End Function
