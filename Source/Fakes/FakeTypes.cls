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

Public Type SimpleType
    Value1 As Long
    Value2 As String
End Type

Public Type ComplexType
    Value1 As Long
    Simple1 As SimpleType
End Type

Public Function NewSimpleType(ByVal Value1 As Long, ByVal Value2 As String) As SimpleType
    NewSimpleType.Value1 = Value1
    NewSimpleType.Value2 = Value2
End Function

Public Function NewComplexType(ByVal Value1 As Long, ByRef Simple1 As SimpleType) As ComplexType
    NewComplexType.Value1 = Value1
    NewComplexType.Simple1 = Simple1
End Function
