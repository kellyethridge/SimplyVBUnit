VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MockITestFixture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements ITestFixture

Public Calls As New CallTrace

Private Sub ITestFixture_GetTestCases(ByVal Tests As TestFixtureBuilder)
    Tests.Add "Test1"
    With Tests.Add("Test3")
        .Use 2
        .Use 3
    End With
    With Tests.Add("Test4")
        .Use(2).Expect 2
        .Use(3).Expect 3
    End With
End Sub

Public Sub FixtureSetup()
    Calls.Add "FixtureSetup"
End Sub

Public Sub Setup()
    Calls.Add "Setup"
End Sub

Public Sub Teardown()
    Calls.Add "Teardown"
End Sub

Public Sub Test1()
    Calls.Add "Test1"
End Sub

Public Sub Test2()
    Calls.Add "Test2"
End Sub

Public Sub Test3(ByVal Value As Long)
    Calls.Add "Test3." & Value
End Sub

Public Function Test4(ByVal Value As Long) As Long
    Calls.Add "Test4." & Value
    Test4 = Value
End Function


