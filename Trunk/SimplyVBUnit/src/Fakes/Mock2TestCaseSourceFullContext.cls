VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Mock2TestCaseSourceFullContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements ITestCaseSource

Public Calls As New CallTrace


Private Sub ITestCaseSource_GetTestCases(ByVal Test As SVBUnit.TestCaseDataBuilder)
    Select Case Test.MethodName
        Case "Test1"
            Call Test.Using(1)
    
        Case "Test2"
            Call Test.Using(1)
    End Select
End Sub

Public Sub Test1(ByVal Value As Variant)
    Call Err.Clear
    Call Calls.Add("Test1")
End Sub

Public Sub Test2(ByVal Value As Variant)
    Call Err.Clear
    Call Calls.Add("Test2")
End Sub

Public Sub Setup()
    Call Err.Clear
    Call Calls.Add("Setup")
End Sub

Public Sub Teardown()
    Call Err.Clear
    Call Calls.Add("Teardown")
End Sub

Public Sub FixtureSetup()
    Call Err.Clear
    Call Calls.Add("FixtureSetup")
End Sub

Public Sub FixtureTeardown()
    Call Err.Clear
    Call Calls.Add("FixtureTeardown")
End Sub