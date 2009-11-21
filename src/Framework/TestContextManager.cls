VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestContextManager"
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
' Module: TestContextManager
'
Option Explicit

Private mCurrentContext As TestContext
Private mContexts       As New Stack


Public Property Get CurrentContext() As TestContext
    Set CurrentContext = mCurrentContext
End Property

Public Sub BeginContext(ByVal Context As TestContext)
    Call StoreCurrentContext
    Call InitializeNewContext(Context)
End Sub

Private Sub StoreCurrentContext()
    Call mContexts.Push(mCurrentContext)
End Sub

Private Sub InitializeNewContext(ByVal Context As TestContext)
    If Not mCurrentContext Is Nothing Then
        Set Context.Error = mCurrentContext.Error
        Set Context.Log = mCurrentContext.Log
        Set Context.Out = mCurrentContext.Out
    End If

    Set mCurrentContext = Context
End Sub

Public Sub EndContext()
    Set mCurrentContext = mContexts.Pop
End Sub

