VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FakeTypes"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type FakeStructure
    Value1 As Long
    Value2 As String
End Type

Public Function NewFakeStructure(ByVal Value1 As Long, ByVal Value2 As String) As FakeStructure
    NewFakeStructure.Value1 = Value1
    NewFakeStructure.Value2 = Value2
End Function
