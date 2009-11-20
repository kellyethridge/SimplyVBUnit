VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Mock1TestCaseSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements ITestCaseSource

Public Calls As New CallTrace
Public Value1 As Variant
Public Value2 As Variant
Public Value3 As Variant

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SVBUnit.TestCaseDataBuilder)
    Calls.Add "GetTestCases." & Test.MethodName
    Test.Using Value1, Value2, Value3
End Sub


Public Sub Test1(ByVal Arg1 As Variant, ByVal Arg2 As Variant, ByVal Arg3 As Variant)
    Calls.Add "Test1." & Arg1 & "." & Arg2 & "." & Arg3
End Sub
