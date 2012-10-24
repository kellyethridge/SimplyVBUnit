VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ResultsTabController"
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
' Module: SComponent.ResultsTabController
'
Option Explicit

Private mOutput As UIListBox
Private mNotRun As TreeView
Private mOut    As TextBox
Private mLog    As TextBox
Private mError  As TextBox
Private mOutOK  As Boolean
Private mLogOK  As Boolean
Private mErrOK  As Boolean

Private WithEvents mListener As EventCastListener
Attribute mListener.VB_VarHelpID = -1


Public Sub SetOutputSupport(ByVal Config As UIConfiguration)
    mOutOK = Config.OutputToTextConsole
    mLogOK = Config.OutputToLogConsole
    mErrOK = Config.OutputToErrorConsole
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Output As UIListBox, ByVal NotRun As TreeView, ByVal Out As TextBox, ByVal Log As TextBox, ByVal Error As TextBox, ByVal Listener As EventCastListener)
    Set mOutput = Output
    Set mNotRun = NotRun
    Set mOut = Out
    Set mLog = Log
    Set mError = Error
    Set mListener = Listener
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BuildTestCaseName(ByVal Result As TestResult) As String
    BuildTestCaseName = Result.Test.Parent.Name & "." & Result.Test.Name
End Function

Private Function StripCrLf(ByRef Value As String) As String
    Do While Right$(Value, 2) = vbCrLf
        Value = Left$(Value, Len(Value) - 2)
    Loop
    
    StripCrLf = Value
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   mListener Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mListener_RunStarted(ByVal Name As String, ByVal TestCount As Long)
    Call mNotRun.Nodes.Clear
    Call mOutput.Clear
    mOut.Text = ""
    mLog.Text = ""
    mError.Text = ""
End Sub

Private Sub mListener_TestCaseFinished(ByVal Result As TestResult)
    If Result.IsFailure Then
        Call mOutput.AddItem(BuildTestCaseName(Result) & " failed: " & StripCrLf(Result.Message))
    ElseIf Result.IsIgnored Then
        Dim Node As Node
        Set Node = mNotRun.Nodes.Add(, , , Result.Test.FullName)
        Call mNotRun.Nodes.Add(Node, tvwChild, , "Reason: " & StripCrLf(Result.Message))
    ElseIf Result.IsError Then
        Call mOutput.AddItem(BuildTestCaseName(Result) & " " & StripCrLf(Result.Message))
    End If
End Sub

Private Sub mListener_TestOutput(ByVal Output As TestOutput)
    Dim Box As TextBox
    
    If (Output.OutputType = TestOutputType.StandardOutput) And mOutOK Then
        Set Box = mOut
    ElseIf (Output.OutputType = TestOutputType.LogOutput) And mLogOK Then
        Set Box = mLog
    ElseIf (Output.OutputType = TestOutputType.ErrorOutput) And mErrOK Then
        Set Box = mError
    End If
    
    If Not Box Is Nothing Then
        Box.SelStart = Len(Box.Text)
        Box.SelText = Output.Text
    End If
End Sub

