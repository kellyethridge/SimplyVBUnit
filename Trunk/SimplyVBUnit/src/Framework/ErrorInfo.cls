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
' Module: ErrorInfo
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
Friend Sub Init(ByVal Number As Long, ByVal Source As String, ByVal Description As String, ByVal LastDllError As Long)
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



