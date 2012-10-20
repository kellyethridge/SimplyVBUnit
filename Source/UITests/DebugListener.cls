VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DebugListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IEventListener

Private mSuccesses As Long
Private mFailures  As Long
Private mErrors    As Long

Private Sub IEventListener_RunException(ByVal Info As SimplyVBUnit.ErrorInfo)

End Sub

Private Sub IEventListener_RunFinished(ByVal Result As SimplyVBUnit.TestResult)
    Dim sb As New StringBuilder
    sb.Append "Passed: " & mSuccesses & vbCrLf
    sb.Append "Failed: " & mFailures & vbCrLf
    sb.Append "Errors: " & mErrors & vbCrLf
    sb.Append "Total : " & mSuccesses + mFailures + mErrors & vbCrLf
    sb.Append "Time  : " & Result.Time & vbCrLf
    
    Debug.Print sb.ToString
    Debug.Print
End Sub

Private Sub IEventListener_RunStarted(ByVal Name As String, ByVal TestCount As Long)
    mSuccesses = 0
    mFailures = 0
    mErrors = 0
    Debug.Print String$(20, "-") & " Run Starting " & String$(20, "-")
End Sub

Private Sub IEventListener_TestCaseFinished(ByVal Result As SimplyVBUnit.TestResult)
    Dim sb As New StringBuilder
    
    If Result.IsSuccess Then
        mSuccesses = mSuccesses + 1
    ElseIf Result.IsError Then
        mErrors = mErrors + 1
        Call sb.AppendFormat("Test '{0}' errored: {1}{2}", Result.Test.Name, Result.Message, vbCrLf)
        Debug.Print sb.ToString
    Else
        mFailures = mFailures + 1
        Call sb.AppendFormat("Test '{0}' failed: {1}{2}", Result.Test.FullName, Result.Message, vbCrLf)
        Debug.Print sb.ToString
    End If
End Sub

Private Sub IEventListener_TestCaseStarted(ByVal Test As SimplyVBUnit.TestCase)

End Sub

Private Sub IEventListener_TestFixtureFinished(ByVal Result As SimplyVBUnit.TestResult)

End Sub

Private Sub IEventListener_TestFixtureStarted(ByVal Fixture As SimplyVBUnit.TestFixture)

End Sub

Private Sub IEventListener_TestOutput(ByVal Output As SimplyVBUnit.TestOutput)

End Sub

Private Sub IEventListener_TestSuiteFinished(ByVal Result As SimplyVBUnit.TestResult)

End Sub

Private Sub IEventListener_TestSuiteStarted(ByVal Suite As SimplyVBUnit.TestSuite)

End Sub

Private Sub IEventListener_UnhandledError(ByVal Info As SimplyVBUnit.ErrorInfo)

End Sub
