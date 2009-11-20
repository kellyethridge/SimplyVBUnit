VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Stub2TestsWithFullAsserts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Test1AssertCount             As Long
Public Test2AssertCount             As Long
Public SetupAssertCount             As Long
Public TeardownAssertCount          As Long
Public FixtureSetupAssertCount      As Long
Public FixtureTeardownAssertCount   As Long


Public Sub Setup()
    DoAsserts Me.SetupAssertCount
End Sub

Public Sub Teardown()
    DoAsserts Me.TeardownAssertCount
End Sub

Public Sub FixtureSetup()
    DoAsserts Me.FixtureSetupAssertCount
End Sub

Public Sub FixtureTeardown()
    DoAsserts Me.FixtureTeardownAssertCount
End Sub

Public Sub Test1()
    DoAsserts Me.Test1AssertCount
End Sub

Public Sub Test2()
    DoAsserts Me.Test2AssertCount
End Sub


Private Sub DoAsserts(ByVal Count As Long)
    Dim i As Long
    For i = 1 To Count
        Assert.IsTrue True
    Next
End Sub
