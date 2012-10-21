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
'    Module: TestProgressController
'

Option Explicit

Private Const WM_USER           As Long = &H400
Private Const PBM_SETBARCOLOR   As Long = WM_USER + 9

Private mBar        As ProgressBar
Private mDisplay    As Label
Private mFailed     As Boolean

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


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   mListener Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mListener_RunFinished(ByVal Result As SimplyVBUnit.TestResult)
    mDisplay.Caption = ""
End Sub

Private Sub mListener_RunStarted(ByVal Name As String, ByVal TestCount As Long)
    mBar.Max = IIf(TestCount = 0, 1, TestCount)
    mBar.Min = 0
    mBar.Value = 0
    mFailed = False
    mDisplay.Caption = ""
    Call SendMessage(mBar.hWnd, PBM_SETBARCOLOR, 0, ByVal vbGreen)
End Sub

Private Sub mListener_TestCaseFinished(ByVal Result As SimplyVBUnit.TestResult)
    If mBar.Value < mBar.Max Then
        mBar.Value = mBar.Value + 1
    End If
    
    If Result.IsFailure Or Result.IsError Then
        Call SendMessage(mBar.hWnd, PBM_SETBARCOLOR, 0, ByVal vbRed)
        mFailed = True
    ElseIf Result.IsIgnored Then
        Call SendMessage(mBar.hWnd, PBM_SETBARCOLOR, 0, ByVal vbYellow)
    End If
End Sub

Private Sub mListener_TestCaseStarted(ByVal Test As SimplyVBUnit.TestCase)
    mDisplay.Caption = Test.FullName
    mDisplay.Refresh
End Sub
