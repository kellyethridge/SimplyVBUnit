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

Public Property Get Number() As Long
    Number = mNumber
End Property

Public Property Get Source() As String
    Source = mSource
End Property

Public Property Get Description() As String
    Description = mDescription
End Property

Public Function Equals(ByVal Other As ErrorInfo) As Boolean
    If Other Is Nothing Then Exit Function
    If Other.Number <> mNumber Then Exit Function
    If Other.Source <> mSource Then Exit Function
    If Other.Description <> mDescription Then Exit Function
    
    Equals = True
End Function

Friend Sub Init(ByVal Number As Long, ByRef Source As String, ByRef Description As String)
    mNumber = Number
    mSource = Source
    mDescription = Description
End Sub
