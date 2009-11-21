VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Mock3TestsWithWriters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Test1Writer As ITextWriter
Public Test2Writer As ITextWriter
Public Test3Writer As ITextWriter

Public Sub Test1()
    Set Test1Writer = New FakeTextWriter
    Set TestContext.Out = Test1Writer
End Sub

Public Sub Test2()
    Set Test2Writer = TestContext.Out
End Sub

Public Sub Test3()
    Set Test3Writer = TestContext.Out
End Sub
