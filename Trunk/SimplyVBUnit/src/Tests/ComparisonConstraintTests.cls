VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ComparisonConstraintTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITestCaseSource

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SVBUnit.TestCaseDataBuilder)
    Select Case Test.MethodName
        Case "TestLessThan"
            Call Test.Use(2, 1).Expect(True)
            Call Test.Use(1, 2).Expect(False)
            Call Test.Use(1.1, 1).Expect(True)
            Call Test.Use(CDec(1.5), 1.4).Expect(True)
            Call Test.Use(1.5, CDec(1.6)).Expect(False)
            Call Test.Use(2@, 2.5@).Expect(False)
            Call Test.Use(2.9@, 2.899@).Expect(True)
            Call Test.Use("a", "A").Expect(True)
            Call Test.Use("A", "a").Expect(False)
            Call Test.Use("b", "b").Expect(False)
            Call Test.Use(1, 1).Expect(False)
            
        Case "TestGreaterThan"
            Call Test.Use(2, 1).Expect(False)
            Call Test.Use(1, 2).Expect(True)
            Call Test.Use(1.1, 1).Expect(False)
            Call Test.Use(CDec(1.5), 1.4).Expect(False)
            Call Test.Use(1.5, CDec(1.6)).Expect(True)
            Call Test.Use(2@, 2.5@).Expect(True)
            Call Test.Use(2.9@, 2.899@).Expect(False)
            Call Test.Use("a", "A").Expect(False)
            Call Test.Use("A", "a").Expect(True)
            Call Test.Use("b", "b").Expect(False)
            Call Test.Use(1, 1).Expect(False)
            
        Case "TestLessThanOrEqualTo"
            Call Test.Use(2, 1).Expect(True)
            Call Test.Use(1, 2).Expect(False)
            Call Test.Use(1.1, 1).Expect(True)
            Call Test.Use(CDec(1.5), 1.4).Expect(True)
            Call Test.Use(1.5, CDec(1.6)).Expect(False)
            Call Test.Use(2@, 2.5@).Expect(False)
            Call Test.Use(2.9@, 2.899@).Expect(True)
            Call Test.Use("a", "A").Expect(True)
            Call Test.Use("A", "a").Expect(False)
            Call Test.Use("b", "b").Expect(True)
            Call Test.Use(1, 1).Expect(True)
        
        Case "TestGreaterThanOrEqualTo"
            Call Test.Use(2, 1).Expect(False)
            Call Test.Use(1, 2).Expect(True)
            Call Test.Use(1.1, 1).Expect(False)
            Call Test.Use(CDec(1.5), 1.4).Expect(False)
            Call Test.Use(1.5, CDec(1.6)).Expect(True)
            Call Test.Use(2@, 2.5@).Expect(True)
            Call Test.Use(2.9@, 2.899@).Expect(False)
            Call Test.Use("a", "A").Expect(False)
            Call Test.Use("A", "a").Expect(True)
            Call Test.Use("b", "b").Expect(True)
            Call Test.Use(1, 1).Expect(True)
        
'        Case "TestRange"
'            Call Test.Use(1, 5, 1).Expect(True)
'            Call Test.Use(1, 5, 0).Expect(False)
'            Call Test.Use(1, 5, 5).Expect(True)
'            Call Test.Use(1, 5, 6).Expect(False)
'            Call Test.Use(1.5, 2.5, 2).Expect(True)
'            Call Test.Use(1.5, 2.5, 2.6).Expect(False)
'            Call Test.Use(1.5, 2.5, 1.4).Expect(False)
            
    End Select
End Sub

Public Function TestLessThan(ByRef Expected As Variant, ByRef Actual As Variant) As Boolean
    Dim c As IConstraint
    Set c = Iz.LessThan(Expected)
    TestLessThan = c.Matches(Actual)
End Function

Public Function TestGreaterThan(ByRef Expected As Variant, ByRef Actual As Variant) As Boolean
    Dim c As IConstraint
    Set c = Iz.GreaterThan(Expected)
    TestGreaterThan = c.Matches(Actual)
End Function

Public Function TestLessThanOrEqualTo(ByRef Expected As Variant, ByRef Actual As Variant) As Boolean
    Dim c As IConstraint
    Set c = Iz.LessThanOrEqualTo(Expected)
    TestLessThanOrEqualTo = c.Matches(Actual)
End Function

Public Function TestGreaterThanOrEqualTo(ByRef Expected As Variant, ByRef Actual As Variant) As Boolean
    Dim c As IConstraint
    Set c = Iz.GreaterThanOrEqualTo(Expected)
    TestGreaterThanOrEqualTo = c.Matches(Actual)
End Function

'Public Function TestRange(ByRef FromValue As Variant, ByRef ToValue As Variant, ByRef Actual As Variant) As Boolean
'    TestRange = Iz.InRange(FromValue, ToValue).Matches(Actual)
'End Function

