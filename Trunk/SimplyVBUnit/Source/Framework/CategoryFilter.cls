VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CategoryFilter"
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
' Module: CategoryFilter
'
Option Explicit
Implements ITestFilter

Private mBase       As New TestFilterBase
Private mCategories As New Collection


Public Sub Add(ByRef Name As String)
    If Len(Name) = 0 Then _
        Error.CannotBeEmpty Me, "Add", "Name"
        
    mCategories.Add Name
End Sub

Public Function Match(ByVal Test As ITest) As Boolean
    If Test.HasCategories Then
        Dim Category As Variant
        For Each Category In mCategories
            If Test.Categories.Contains(CStr(Category)) Then
                Match = True
                Exit Function
            End If
        Next
    End If
End Function

Public Function Pass(ByVal Test As ITest) As Boolean
    If Test.HasCategories Then
        Pass = mBase.Pass(Me, Test)
    Else
        Pass = True
    End If
End Function

Public Property Get Count() As Long
    Count = mCategories.Count
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ITestFilter Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ITestFilter_Match(ByVal Test As ITest) As Boolean
    ITestFilter_Match = Match(Test)
End Function

Private Function ITestFilter_Pass(ByVal Test As ITest) As Boolean
    ITestFilter_Pass = Pass(Test)
End Function
