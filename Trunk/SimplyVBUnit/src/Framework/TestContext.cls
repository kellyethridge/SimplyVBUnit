VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestContext"
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
' Module: TestContext
'
Option Explicit

Private mLog            As ITextWriter
Private mError          As ITextWriter
Private mOut            As ITextWriter
Private mThrows         As IConstraint
Private mAssertCount    As Long


Public Property Get Out() As ITextWriter
    Set Out = mOut
End Property

Public Property Set Out(ByVal RHS As ITextWriter)
    Set mOut = RHS
End Property

Public Property Get Error() As ITextWriter
    Set Error = mError
End Property

Public Property Set Error(ByVal RHS As ITextWriter)
    Set mError = RHS
End Property

Public Property Get Log() As ITextWriter
    Set Log = mLog
End Property

Public Property Set Log(ByVal RHS As ITextWriter)
    Set mLog = RHS
End Property

Public Property Get Throws() As IConstraint
Attribute Throws.VB_MemberFlags = "40"
    Set Throws = mThrows
End Property

Public Property Get AssertCount() As Long
Attribute AssertCount.VB_MemberFlags = "40"
    AssertCount = mAssertCount
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Property Set Throws(ByVal RHS As IConstraint)
    Set mThrows = RHS
End Property

Friend Sub IncrementAssertCount()
    mAssertCount = mAssertCount + 1
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    TestContextManager.BeginContext Me
End Sub

Private Sub Class_Terminate()
    TestContextManager.EndContext
End Sub
