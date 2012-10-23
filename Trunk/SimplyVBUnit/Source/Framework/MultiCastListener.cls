VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MultiCastListener"
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
' Module: SComponent.MultiCastListener
'
Option Explicit
Implements IEventListener

Private mListeners As New Collection

Public Sub Add(ByVal Listener As IEventListener)
    If Listener Is Nothing Then _
        Error.CannotBeNothing Me, "Add", "Listener"
    
    mListeners.Add Listener
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IEventListener Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub IEventListener_RunException(ByVal Info As ErrorInfo)
    Dim Listener As IEventListener
    For Each Listener In mListeners
        Listener.RunException Info
    Next
End Sub

Private Sub IEventListener_RunFinished(ByVal Result As TestResult)
    Dim Listener As IEventListener
    For Each Listener In mListeners
        Listener.RunFinished Result
    Next
End Sub

Private Sub IEventListener_RunStarted(ByVal Name As String, ByVal TestCount As Long)
    Dim Listener As IEventListener
    For Each Listener In mListeners
        Listener.RunStarted Name, TestCount
    Next
End Sub

Private Sub IEventListener_TestCaseFinished(ByVal Result As TestResult)
    Dim Listener As IEventListener
    For Each Listener In mListeners
        Listener.TestCaseFinished Result
    Next
End Sub

Private Sub IEventListener_TestCaseStarted(ByVal Test As TestCase)
    Dim Listener As IEventListener
    For Each Listener In mListeners
        Listener.TestCaseStarted Test
    Next
End Sub

Private Sub IEventListener_TestOutput(ByVal Output As TestOutput)
    Dim Listener As IEventListener
    For Each Listener In mListeners
        Listener.TestOutput Output
    Next
End Sub

Private Sub IEventListener_TestSuiteFinished(ByVal Result As TestResult)
    Dim Listener As IEventListener
    For Each Listener In mListeners
        Listener.TestSuiteFinished Result
    Next
End Sub

Private Sub IEventListener_TestSuiteStarted(ByVal Suite As TestSuite)
    Dim Listener As IEventListener
    For Each Listener In mListeners
        Listener.TestSuiteStarted Suite
    Next
End Sub

Private Sub IEventListener_UnhandledError(ByVal Info As ErrorInfo)
    Dim Listener As IEventListener
    For Each Listener In mListeners
        Listener.UnhandledError Info
    Next
End Sub
