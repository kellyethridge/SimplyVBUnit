VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Stub1TestWithFullContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Errors As New CallError

Public Sub Setup()
    Errors.Raise "Setup"
End Sub

Public Sub Teardown()
    Errors.Raise "Teardown"
End Sub

Public Sub FixtureSetup()
    Errors.Raise "FixtureSetup"
End Sub

Public Sub FixtureTeardown()
    Errors.Raise "FixtureTeardown"
End Sub

Public Sub Test1()
    Errors.Raise "Test1"
End Sub
