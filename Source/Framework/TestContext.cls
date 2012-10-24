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
    Set mLog = NullTextWriter.Instance
    Set mError = NullTextWriter.Instance
    Set mOut = NullTextWriter.Instance
    TestContextManager.BeginContext Me
End Sub

Private Sub Class_Terminate()
    TestContextManager.EndContext
End Sub
