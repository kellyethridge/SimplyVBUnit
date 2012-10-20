VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FakeTestContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements TestContext

Private mError  As ITextWriter
Private mLog    As ITextWriter
Private mOut    As ITextWriter


Private Property Get TestContext_AssertCount() As Long

End Property

Private Property Set TestContext_Error(ByVal RHS As ITextWriter)
    Set mError = RHS
End Property

Private Property Get TestContext_Error() As ITextWriter
    Set TestContext_Error = mError
End Property

Private Property Set TestContext_Log(ByVal RHS As ITextWriter)
    Set mLog = RHS
End Property

Private Property Get TestContext_Log() As ITextWriter
    Set TestContext_Log = mLog
End Property

Private Property Set TestContext_Out(ByVal RHS As ITextWriter)
    Set mOut = RHS
End Property

Private Property Get TestContext_Out() As ITextWriter
    Set TestContext_Out = mOut
End Property

Private Property Get TestContext_Throws() As IConstraint

End Property
