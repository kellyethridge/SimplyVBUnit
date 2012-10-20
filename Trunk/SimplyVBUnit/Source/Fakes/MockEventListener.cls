VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MockEventListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IEventListener

Public Calls        As New CallTrace
Public ErrorInfo    As ErrorInfo
Public Output       As TestOutput

Private Sub CallMade(ByVal MethodName As String)
    Calls.Add MethodName
    Err.Clear
End Sub

Private Sub IEventListener_RunException(ByVal Info As ErrorInfo)
    CallMade "RunException"
End Sub

Private Sub IEventListener_RunFinished(ByVal Result As TestResult)
    CallMade "RunFinished"
End Sub

Private Sub IEventListener_RunStarted(ByVal Name As String, ByVal TestCount As Long)
    CallMade "RunStarted"
End Sub

Private Sub IEventListener_TestCaseFinished(ByVal Result As TestResult)
    CallMade "TestCaseFinished"
End Sub

Private Sub IEventListener_TestCaseStarted(ByVal Test As TestCase)
    CallMade "TestCaseStarted"
End Sub

Private Sub IEventListener_TestOutput(ByVal Output As TestOutput)
    CallMade "TestOutput"
    Set Me.Output = Output
End Sub

Private Sub IEventListener_TestSuiteFinished(ByVal Result As TestResult)
    CallMade "TestSuiteFinished"
End Sub

Private Sub IEventListener_TestSuiteStarted(ByVal Suite As TestSuite)
    CallMade "TestSuiteStarted"
End Sub

Private Sub IEventListener_UnhandledError(ByVal Info As ErrorInfo)
    Set ErrorInfo = Info
    CallMade "UnhandledError"
End Sub
