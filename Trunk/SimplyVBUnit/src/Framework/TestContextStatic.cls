VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestContextStatic"
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
' Module: TestContextStatic
'
Option Explicit

Private mContexts   As New TestContextStack
Private mCurrent    As Long


Public Property Get Out() As ITextWriter
    If HasContext Then
        Set Out = Current.Out
    End If
End Property

Public Property Set Out(ByVal RHS As ITextWriter)
    If HasContext Then
        Set Current.Out = RHS
    End If
End Property

Public Property Get Error() As ITextWriter
    If HasContext Then
        Set Error = Current.Error
    End If
End Property

Public Property Set Error(ByVal RHS As ITextWriter)
    If HasContext Then
        Set Current.Error = RHS
    End If
End Property

Public Property Get Log() As ITextWriter
    If HasContext Then
        Set Log = Current.Log
    End If
End Property

Public Property Set Log(ByVal RHS As ITextWriter)
    If HasContext Then
        Set Current.Log = RHS
    End If
End Property

Private Property Get Current() As TestContext
    Set Current = TestUtils.GetObjectFromPointer(mCurrent)
End Property

Private Function HasContext() As Boolean
    HasContext = (mCurrent <> vbNullPtr)
End Function

Friend Sub BeginContext(ByVal Context As TestContext)
    Call CopyWritersTo(Context)
    Call StoreAndReplaceCurrentContext(Context)
End Sub

Private Sub CopyWritersTo(ByVal Context As TestContext)
    If HasContext Then
        Call Context.CopyWritersFrom(Current)
    End If
End Sub

Private Sub StoreAndReplaceCurrentContext(ByVal Context As TestContext)
    Call mContexts.Push(Current)
    Call SetCurrent(Context)
End Sub

Friend Sub EndContext()
    Call SetCurrent(mContexts.Pop)
End Sub

Private Sub SetCurrent(ByVal Context As TestContext)
    mCurrent = ObjPtr(Context)
End Sub
