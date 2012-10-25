VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Mock1TestCaseSourceExpect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements ITestCaseSource

Public ParameterValue   As Variant
Public Constraint       As New MockConstraint
Public Calls            As New CallTrace


Private Sub ITestCaseSource_GetTestCases(ByVal Test As TestCaseCollector)
    Select Case Test.MethodName
        Case "Test1"
            Test.Use(ParameterValue).Expect Constraint
        
        Case "Test2"
            Call Test.Use(1).Expect(Iz.EqualTo(1))
    End Select
End Sub

Public Function Test1(ByVal Value As Variant) As Variant
    Calls.Add "Test1." & CStr(Value)
    VariantCopyInd Test1, Value
End Function

Public Function Test2(ByVal Value As Variant) As Variant
    Test2 = Value
End Function

