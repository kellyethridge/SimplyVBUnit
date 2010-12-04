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
Attribute VB_Creatable = True
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
' Module: TestRunner
'
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
' Module: TestRunner
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
