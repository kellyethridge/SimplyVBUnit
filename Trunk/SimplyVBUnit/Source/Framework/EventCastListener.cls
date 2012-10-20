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
' Module: EventCastListener
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
