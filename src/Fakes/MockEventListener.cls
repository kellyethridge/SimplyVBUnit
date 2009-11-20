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

Private Sub CallMade(ByVal MethodName As String)
    Call Calls.Add(MethodName)
    Call Err.Clear
End Sub

Private Sub IEventListener_RunException(ByVal Info As SVBUnit.ErrorInfo)
    Call CallMade("RunException")
End Sub

Private Sub IEventListener_RunFinished(ByVal Result As SVBUnit.TestResult)
    Call CallMade("RunFinished")
End Sub

Private Sub IEventListener_RunStarted(ByVal Name As String, ByVal TestCount As Long)
    Call CallMade("RunStarted")
End Sub

Private Sub IEventListener_TestCaseFinished(ByVal Result As SVBUnit.TestResult)
    Call CallMade("TestCaseFinished")
End Sub

Private Sub IEventListener_TestCaseStarted(ByVal Test As SVBUnit.TestCase)
    Call CallMade("TestCaseStarted")
End Sub

Private Sub IEventListener_TestFixtureFinished(ByVal Result As SVBUnit.TestResult)
    Call CallMade("TestFixtureFinished")
End Sub

Private Sub IEventListener_TestFixtureStarted(ByVal Fixture As SVBUnit.TestFixture)
    Call CallMade("TestFixtureStarted")
End Sub

Private Sub IEventListener_TestOutput(ByVal Output As SVBUnit.TestOutput)
    Call CallMade("TestOutput")
End Sub

Private Sub IEventListener_TestSuiteFinished(ByVal Result As SVBUnit.TestResult)
    Call CallMade("TestSuiteFinished")
End Sub

Private Sub IEventListener_TestSuiteStarted(ByVal Suite As SVBUnit.TestSuite)
    Call CallMade("TestSuiteStarted")
End Sub

Private Sub IEventListener_UnhandledError(ByVal Info As SVBUnit.ErrorInfo)
    Set ErrorInfo = Info
    Call CallMade("UnhandledError")
End Sub
