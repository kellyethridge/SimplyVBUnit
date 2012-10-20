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
'    CopyRight (c) 2008 Kelly Ethridge
'
'    This file is part of SimplyVBUnitUI.
'
'    SimplyVBUnitUI is free software; you can redistribute it and/or modify
'    it under the terms of the GNU Library General Public License as published by
'    the Free Software Foundation; either version 2.1 of the License, or
'    (at your option) any later version.
'
'    SimplyVBUnitUI is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Library General Public License for more details.
'
'    You should have received a copy of the GNU Library General Public License
'    along with Foobar; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'    Module: StatusBarController
'

Option Explicit

Private mBar As StatusBar
Private WithEvents mListener As EventCastListener
Attribute mListener.VB_VarHelpID = -1
Private mFailures As Long
Private mTestsRun As Long


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

Private Sub mListener_RunFinished(ByVal Result As SimplyVBUnit.TestResult)
    mBar.Panels("Status").Text = "Tests Completed"
    mBar.Panels("TestsRun").Text = "Tests Run: " & mTestsRun
    mBar.Panels("Failures").Text = "Failures: " & mFailures
    mBar.Panels("Time").Text = Result.Time & "ms"
End Sub

Private Sub mListener_RunStarted(ByVal Name As String, ByVal TestCount As Long)
    mFailures = 0
    mTestsRun = 0
End Sub

Private Sub mListener_TestCaseFinished(ByVal Result As SimplyVBUnit.TestResult)
    If Result.IsFailure Then
        mFailures = mFailures + 1
    End If
    
    mTestsRun = mTestsRun + 1
End Sub
