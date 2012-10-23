VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ErrorInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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
' Module: SComponent.ErrorInfo
'
Option Explicit

Private mNumber         As Long
Private mSource         As String
Private mDescription    As String
Private mLastDllError   As Long


Public Property Get Number() As Long
    Number = mNumber
End Property

Public Property Get Source() As String
    Source = mSource
End Property

Public Property Get Description() As String
    Description = mDescription
End Property

Public Property Get LastDllError() As Long
    LastDllError = mLastDllError
End Property

Public Function Equals(ByVal Other As ErrorInfo, Optional ByVal IgnoreCase As Boolean) As Boolean
    If Other Is Nothing Then Exit Function
    
    Dim Result As Boolean
    Result = (Other.Number = Me.Number)
    Result = Result And EqualStringsIfBothExist(Me.Source, Other.Source, IgnoreCase)
    Result = Result And EqualStringsIfBothExist(Me.Description, Other.Description, IgnoreCase)
    Result = Result And (Other.LastDllError = Me.LastDllError)
    
    Equals = Result
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Number As Long, ByRef Source As String, ByRef Description As String, ByVal LastDllError As Long)
    mNumber = Number
    mSource = Source
    mDescription = Description
    mLastDllError = LastDllError
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function EqualStringsIfBothExist(ByRef String1 As String, ByRef String2 As String, ByVal IgnoreCase As Boolean) As Boolean
    If ShouldCompare(String1, String2) Then
        If Not EqualStrings(String1, String2, IgnoreCase) Then
            Exit Function
        End If
    End If
    
    EqualStringsIfBothExist = True
End Function

Private Function ShouldCompare(ByRef Text1 As String, ByRef Text2 As String) As Boolean
    If Len(Text1) > 0 Then
        ShouldCompare = (Len(Text2) > 0)
    End If
End Function



