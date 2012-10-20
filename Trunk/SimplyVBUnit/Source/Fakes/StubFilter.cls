VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StubFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements ITestFilter

Public Matches  As Boolean
Public Passes   As Boolean

Private Function ITestFilter_Match(ByVal Test As SVBUnit.ITest) As Boolean
    ITestFilter_Match = Matches
End Function

Private Function ITestFilter_Pass(ByVal Test As SVBUnit.ITest) As Boolean
    ITestFilter_Pass = Passes
End Function
