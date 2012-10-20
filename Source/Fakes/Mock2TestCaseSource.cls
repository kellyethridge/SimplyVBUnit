VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Mock2TestCaseSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements ITestCaseSource

Public Calls As New CallTrace
Public Value1 As Variant
Public Value2 As Variant


Private Sub ITestCaseSource_GetTestCases(ByVal Test As TestCaseCollector)
    Calls.Add "GetTestCases." & Test.MethodName
    
    Test.Use Value1
    Test.Use Value2
End Sub

Public Sub Test1(ByVal Value As Variant)
    Err.Clear
    Calls.Add "Test1"
End Sub

