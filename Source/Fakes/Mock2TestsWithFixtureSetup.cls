VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Mock2TestsWithFixtureSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Calls    As New CallTrace
Public Errors   As New CallError

Public Sub FixtureSetup()
    Called "FixtureSetup"
End Sub

Public Sub Test1()
    Called "Test1"
End Sub

Public Sub Test2()
    Called "Test2"
End Sub

Private Sub Called(ByVal MethodName As String)
    Err.Clear
    Calls.Add MethodName
    Errors.Raise MethodName
End Sub
