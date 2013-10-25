VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemoryLeakTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Assert_WhenCallFinishes_TerminateShouldBeCalled()
    Dim Actual As CallTrace
    
    On Error GoTo errTrap
    Assert.That CreateMonitor(Actual), Iz.Nothing
    
errTrap:
    Err.Clear
    Dim Expected As New CallTrace
    Expected.Add "Class_Initialize"
    Expected.Add "Class_Terminate"
    Assert.IsTrue Actual.Equals(Expected), "Assert did not release value."
End Sub


Private Function CreateMonitor(ByRef Trace As CallTrace) As ClassEventMonitor
    Set CreateMonitor = New ClassEventMonitor
    Set Trace = CreateMonitor.Trace
End Function

