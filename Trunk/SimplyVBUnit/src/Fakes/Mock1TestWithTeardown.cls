VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Mock1TestWithTeardown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Errors   As New CallError
Public Calls    As New CallTrace


Public Sub Teardown()
    Call Called("Teardown")
End Sub

Public Sub Test1()
    Call Called("Test1")
End Sub


Private Sub Called(ByVal MethodName As String)
    Call Err.Clear
    Call Calls.Add(MethodName)
    Call Errors.Raise(MethodName)
End Sub
