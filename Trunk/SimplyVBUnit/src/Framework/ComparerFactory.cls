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
' Module: ComparerFactory
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
