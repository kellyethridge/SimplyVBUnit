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
' Copyright 2011 Kelly Ethridge
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
' Module: MultiCastListener
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
