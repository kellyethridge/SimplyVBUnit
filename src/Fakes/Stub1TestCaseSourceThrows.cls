VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Stub1TestCaseSourceThrows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements ITestCaseSource

Public Errors       As New CallError
Public Number       As Long
Public Source       As String
Public Description  As String


Private Sub ITestCaseSource_GetTestCases(ByVal Test As SVBUnit.TestCaseCollector)
    Select Case Test.MethodName
        Case "Test1"
            Call Test.Use(1).Throws(Number, Source, Description)
        
    End Select
End Sub

Public Sub Test1(ByVal Value As Variant)
    Errors.Raise "Test1"
End Sub
