VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NullListener"
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
' Module: NullListener
'
Option Explicit
Implements IEventListener

Private Sub IEventListener_RunException(ByVal Info As ErrorInfo)

End Sub

Private Sub IEventListener_RunFinished(ByVal Result As TestResult)

End Sub

Private Sub IEventListener_RunStarted(ByVal Name As String, ByVal TestCount As Long)

End Sub

Private Sub IEventListener_TestCaseFinished(ByVal Result As TestResult)

End Sub

Private Sub IEventListener_TestCaseStarted(ByVal Test As TestCase)

End Sub

Private Sub IEventListener_TestFixtureFinished(ByVal Result As TestResult)

End Sub

Private Sub IEventListener_TestFixtureStarted(ByVal Fixture As TestFixture)

End Sub

Private Sub IEventListener_TestOutput(ByVal Output As TestOutput)

End Sub

Private Sub IEventListener_TestSuiteFinished(ByVal Result As TestResult)

End Sub

Private Sub IEventListener_TestSuiteStarted(ByVal Suite As TestSuite)

End Sub

Private Sub IEventListener_UnhandledError(ByVal Info As ErrorInfo)

End Sub
