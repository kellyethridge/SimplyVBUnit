VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StatusBarController"
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
' Module: SComponent.StatusBarController
'
Option Explicit

Private mBar        As StatusBar
Private mFailures   As Long
Private mTestsRun   As Long

Private WithEvents mListener As EventCastListener
Attribute mListener.VB_VarHelpID = -1


Friend Sub Init(ByVal Bar As StatusBar, ByVal Listener As EventCastListener)
    Set mBar = Bar
    Set mListener = Listener
End Sub

Friend Sub Reset(ByVal TestCount As Long)
    mBar.Panels("Status").Text = "Running"
    mBar.Panels("TestCases").Text = "Test Cases: " & TestCount
    mBar.Panels("Failures").Text = "Failures: 0"
    mBar.Panels("TestsRun").Text = "Tests Run: 0"
    mBar.Panels("Time").Text = "Time: 0ms"
End Sub

Private Sub mListener_RunFinished(ByVal Result As TestResult)
    mBar.Panels("Status").Text = "Tests Completed"
    mBar.Panels("TestsRun").Text = "Tests Run: " & mTestsRun
    mBar.Panels("Failures").Text = "Failures: " & mFailures
    mBar.Panels("Time").Text = Result.Time & "ms"
End Sub

Private Sub mListener_RunStarted(ByVal Name As String, ByVal TestCount As Long)
    mFailures = 0
    mTestsRun = 0
End Sub

Private Sub mListener_TestCaseFinished(ByVal Result As TestResult)
    If Result.IsFailure Then
        mFailures = mFailures + 1
    End If
    
    mTestsRun = mTestsRun + 1
End Sub
