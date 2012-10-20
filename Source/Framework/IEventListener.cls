VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IEventListener"
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
' Module: IEventListener
'

''
' Provides an interface used to listen to events within the testing framework.
'
' @remarks Implement this interface to allow the framework to talk back to the class
' and give the class the ability to respond to certain events.
'
' @see EventCastListener
' @see MultiCastListener
'
Option Explicit

''
' This method is called when an <i>ITestRunner</i> is started.
'
' @param Name The name of the first test in the run (usually a <i>TestSuite</i>).
' @param TestCount The total number of tests to be run. TestSuites and TestFixtures do not count.
'
Public Sub RunStarted(ByVal Name As String, ByVal TestCount As Long): End Sub

''
' This method is called when an <i>ITestRunner</i> has finished.
'
' @param Result The final result of the tests that were run.
'
Public Sub RunFinished(ByVal Result As TestResult): End Sub

''
' This method is called when an error occurs outside of a test.
'
' @param ErrInfo Information about the error that occurred.
'
Public Sub RunException(ByVal Info As ErrorInfo): End Sub

''
' This method is called when a <i>TestSuite</i> is started.
'
' @param Suite The <i>TestSuite</i> that is starting.
'
Public Sub TestSuiteStarted(ByVal Suite As TestSuite): End Sub

''
' This method is called when a <i>TestSuite</i> has finished running.
'
' @param Result The final result for the test suite.
'
Public Sub TestSuiteFinished(ByVal Result As TestResult): End Sub

''
' This method is called when a test case is started running.
'
' @param Test The test case to be run.
'
Public Sub TestCaseStarted(ByVal Test As TestCase): End Sub

''
' This method is called when a test case has finished running.
'
' @param Result The final result for the test case.
'
Public Sub TestCaseFinished(ByVal Result As TestResult): End Sub

''
' This method is called when user output is being sent.
'
' @param Output The output the user wishes to send.
'
Public Sub TestOutput(ByVal Output As TestOutput): End Sub

''
' This method is called when an unhandled error occurs during a test.
'
' @param Info Information about the error.
'
Public Sub UnhandledError(ByVal Info As ErrorInfo): End Sub



