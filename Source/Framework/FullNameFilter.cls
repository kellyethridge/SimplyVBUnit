VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FullNameFilter"
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
' Module: FullNameFilter
'
Option Explicit
Implements ITestFilter

Private mBase   As New TestFilterBase
Private mNames  As New Collection
Private mLikes  As New Collection

Public Sub Add(ByRef Name As String)
    On Error GoTo errTrap
    
    If Right$(Name, 1) = "*" Then
        mLikes.Add Name
    Else
        mNames.Add Name, Name
    End If
errTrap:
End Sub

Public Function Match(ByVal Test As ITest) As Boolean
    On Error GoTo errTrap
    
    Dim NameToCheck As String
    NameToCheck = Test.FullName
    
    Dim Name As String
    Name = mNames(NameToCheck)
    Match = True
    Exit Function
    
errTrap:
    Dim NameExp As Variant
    For Each NameExp In mLikes
        If NameToCheck Like NameExp Then
            Match = True
            Exit Function
        End If
    Next
End Function

Public Function Pass(ByVal Test As ITest) As Boolean
    Pass = mBase.Pass(Me, Test)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ITestFilter Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ITestFilter_Match(ByVal Test As ITest) As Boolean
    ITestFilter_Match = Match(Test)
End Function

Private Function ITestFilter_Pass(ByVal Test As ITest) As Boolean
    ITestFilter_Pass = Pass(Test)
End Function
