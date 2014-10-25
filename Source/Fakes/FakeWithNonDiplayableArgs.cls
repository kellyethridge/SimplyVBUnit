VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FakeWithNonDiplayableArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements ITestCaseSource

Public Argument As String

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBUnit.TestCaseCollector)
    Select Case Test.MethodName
        Case "TestWithArgs"
            Test.Use Argument
    End Select
End Sub


Public Sub TestWithArgs(ByVal s As String)

End Sub
