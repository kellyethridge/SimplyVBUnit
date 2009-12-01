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


Public Property Get Out() As ITextWriter
    Set Out = TestContextManager.CurrentContext.Out
End Property

Public Property Set Out(ByVal RHS As ITextWriter)
    Set TestContextManager.CurrentContext.Out = RHS
End Property

Public Property Get Error() As ITextWriter
    Set Error = TestContextManager.CurrentContext.Error
End Property

Public Property Set Error(ByVal RHS As ITextWriter)
    Set TestContextManager.CurrentContext.Error = RHS
End Property

Public Property Get Log() As ITextWriter
    Set Log = TestContextManager.CurrentContext.Log
End Property

Public Property Set Log(ByVal RHS As ITextWriter)
    Set TestContextManager.CurrentContext.Log = RHS
End Property

Friend Property Get Throws() As ThrowsConstraint
    Set Throws = TestContextManager.CurrentContext.Throws
End Property

Friend Property Set Throws(ByVal RHS As ThrowsConstraint)
    Set TestContextManager.CurrentContext.Throws = RHS
End Property

Friend Property Get AssertCount() As Long
    AssertCount = TestContextManager.CurrentContext.AssertCount
End Property

Friend Sub IncrementAssertCount()
    Call TestContextManager.CurrentContext.IncrementAssertCount
End Sub
