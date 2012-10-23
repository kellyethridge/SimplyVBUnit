VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestRunner"
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
' Module: SComponent.TestRunner
'
Option Explicit
Implements ITestRunner

Private mTest           As ITest
Private mAllowCancel    As Boolean
Private mIsCancelled    As Boolean


Public Property Get Test() As ITest
    Set Test = mTest
End Property

Public Function Run(Optional ByVal Listener As IEventListener, Optional ByVal Filter As ITestFilter) As TestResult
    If Listener Is Nothing Then
        Set Listener = NullListener.Instance
    End If
    If Filter Is Nothing Then
        Set Filter = TestFilter.EmptyFilter
    End If
        
    With New TestContext
        On Error GoTo errTrap
        mTest.SetRunner Me
        Listener.RunStarted mTest.Name, mTest.CountTestCases(Filter)
        
        Dim Result As TestResult
        Set Result = mTest.Run(Listener, Filter)
        
        Listener.RunFinished Result
        mTest.SetRunner NullTestRunner.Instance
        Set Run = Result
        Exit Function
        
errTrap:
        Listener.RunException ErrorInfo.FromErr(Err)
        mTest.SetRunner NullTestRunner.Instance
        
        Set Run = Sim.NewTestResult(mTest)
        Run.Error Err.Number & ", " & Err.Description, ParentSite
    End With
End Function

Public Property Get IsCancelled() As Boolean
    If mAllowCancel Then
        DoEvents
        IsCancelled = mIsCancelled
    End If
End Property

Public Sub Cancel()
    mIsCancelled = mAllowCancel
End Sub

Public Property Get AllowCancel() As Boolean
    AllowCancel = mAllowCancel
End Property

Public Property Let AllowCancel(ByVal RHS As Boolean)
    mAllowCancel = RHS
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Test As ITest)
    If Test Is Nothing Then _
        Error.CannotBeNothing Me, "Init", "Test"
    
    Set mTest = Test
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ITestRunner Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ITestRunner_Cancel()
    Cancel
End Sub

Private Property Get ITestRunner_IsCancelled() As Boolean
    ITestRunner_IsCancelled = IsCancelled
End Property

Private Function ITestRunner_Run(Optional ByVal Listener As IEventListener, Optional ByVal Filter As ITestFilter) As TestResult
    Set ITestRunner_Run = Run(Listener, Filter)
End Function
