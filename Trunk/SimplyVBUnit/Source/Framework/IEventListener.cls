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
' Module: IEventListener
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



