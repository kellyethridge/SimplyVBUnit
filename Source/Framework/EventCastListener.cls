VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EventCastListener"
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
' Module: SComponent.EventCastListener
'
Option Explicit
Implements IEventListener

Public Event RunException(ByVal Info As ErrorInfo)
Public Event RunFinished(ByVal Rules As TestResult)
Public Event RunStarted(ByVal Name As String, ByVal TestCount As Long)
Public Event TestCaseFinished(ByVal Result As TestResult)
Public Event TestCaseStarted(ByVal Test As TestCase)
Public Event TestOutput(ByVal Output As TestOutput)
Public Event TestSuiteFinished(ByVal Result As TestResult)
Public Event TestSuiteStarted(ByVal Suite As TestSuite)
Public Event UnhandledError(ByVal Info As ErrorInfo)


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IEventListener Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub IEventListener_RunException(ByVal Info As ErrorInfo)
    RaiseEvent RunException(Info)
End Sub

Private Sub IEventListener_RunFinished(ByVal Result As TestResult)
    RaiseEvent RunFinished(Result)
End Sub

Private Sub IEventListener_RunStarted(ByVal Name As String, ByVal TestCount As Long)
    RaiseEvent RunStarted(Name, TestCount)
End Sub

Private Sub IEventListener_TestCaseFinished(ByVal Result As TestResult)
    RaiseEvent TestCaseFinished(Result)
End Sub

Private Sub IEventListener_TestCaseStarted(ByVal Test As TestCase)
    RaiseEvent TestCaseStarted(Test)
End Sub

Private Sub IEventListener_TestOutput(ByVal Output As TestOutput)
    RaiseEvent TestOutput(Output)
End Sub

Private Sub IEventListener_TestSuiteFinished(ByVal Result As TestResult)
    RaiseEvent TestSuiteFinished(Result)
End Sub

Private Sub IEventListener_TestSuiteStarted(ByVal Suite As TestSuite)
    RaiseEvent TestSuiteStarted(Suite)
End Sub

Private Sub IEventListener_UnhandledError(ByVal Info As ErrorInfo)
    RaiseEvent UnhandledError(Info)
End Sub
