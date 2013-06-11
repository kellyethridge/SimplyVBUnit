VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestProgressController"
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
' Module: SComponent.TestProgressController
'
Option Explicit

Private Const WM_USER           As Long = &H400
Private Const PBM_SETBARCOLOR   As Long = WM_USER + 9

Private mBar        As ProgressBar
Private mDisplay    As Label
Private mFailed     As Boolean
Private mConfig     As UIConfiguration
Private mUpdateFreq As Long
Private mTestCount  As Long

Private WithEvents mListener As EventCastListener
Attribute mListener.VB_VarHelpID = -1


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Bar As ProgressBar, ByVal Display As Label, ByVal Listener As EventCastListener)
    Set mBar = Bar
    Set mDisplay = Display
    Set mListener = Listener
End Sub

Friend Property Set Config(ByVal RHS As UIConfiguration)
    Set mConfig = RHS
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   mListener Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mListener_RunFinished(ByVal Result As TestResult)
    mDisplay.Caption = ""
    mBar.Value = mBar.Max
End Sub

Private Sub mListener_RunStarted(ByVal Name As String, ByVal TestCount As Long)
    mUpdateFreq = mConfig.DoEventsFrequency
    mTestCount = 0
    mBar.Max = IIf(TestCount = 0, 1, TestCount)
    mBar.Min = 0
    mBar.Value = 0
    mFailed = False
    mDisplay.Caption = ""
    SendMessage mBar.hWnd, PBM_SETBARCOLOR, 0, ByVal vbGreen
End Sub

Private Sub mListener_TestCaseFinished(ByVal Result As TestResult)
    mTestCount = mTestCount + 1
    If mTestCount = mUpdateFreq And mBar.Value < mBar.Max Then
        mBar.Value = mBar.Value + mUpdateFreq
        mTestCount = 0
    End If
    
    If Not mFailed Then
        If Result.IsFailure Or Result.IsError Then
            SendMessage mBar.hWnd, PBM_SETBARCOLOR, 0, ByVal vbRed
            mFailed = True
        ElseIf Result.IsIgnored Then
            SendMessage mBar.hWnd, PBM_SETBARCOLOR, 0, ByVal vbYellow
        End If
    End If
End Sub

Private Sub mListener_TestCaseStarted(ByVal Test As TestCase)
    mDisplay.Caption = Test.FullName
    mDisplay.Refresh
End Sub
