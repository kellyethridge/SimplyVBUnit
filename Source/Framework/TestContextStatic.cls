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
' Module: SComponent.TestContextStatic
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


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
    TestContextManager.CurrentContext.IncrementAssertCount
End Sub
