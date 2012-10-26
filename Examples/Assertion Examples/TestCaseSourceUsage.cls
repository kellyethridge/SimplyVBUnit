VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestCaseSourceUsage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestRowData
'
' This class demonstrates the new ability to provide multiple sets of data to be passed into
' a single test method. This gives the ability perform a wide variety of tests without having
' to write an entire test case for each scenerio.
'
Option Explicit
Implements ITestCaseSource

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    ' Determine the method that SimplyVBUnit is requesting test case data for.
    ' Each set of test case data added for that method will result in a new unit test
    ' to be executed with the specified parameters.
    Select Case Test.MethodName
    
        ' creates 3 sets of values to be passed to the specified method name. Each set
        ' of values creates a new unit test for SimplyVBUnit to execute. An unlimited
        ' number of sets can be creates for each method.
        Case "CompareDoubleValueWithoutReturn"
            Call Test.Use(2, 4)
            Call Test.Use(5, 10)
            Call Test.Use(20, 40)
        
        Case "CompareDoubleValueWithReturn"
            ' These test cases set an expectation of a result from the specified method.
            ' The expectation can be any value. If the value implements the IConstraint
            ' interface, then the expected constraint is used to test the return value.
            ' All other values are simply tested for equality by using the Iz.EqualTo constraint.
            Call Test.Use(2).Expect(4)
            Call Test.Use(5).Expect(10)
            Call Test.Use(20).Expect(Iz.EqualTo(40))
            Call Test.Use(20).Expect(Iz.InRange(30, 50))
            
        Case "DivideValue"
            Const DivideByZeroError As Long = 11
            Call Test.Use(0).Throws(DivideByZeroError)
    End Select
End Sub

' this will be called several times with different paremeter values.
'
' This test simply accepts two values, and tests that multiplying the original value by 2
' will equal to the expected value. The test is performed entirely within the method body.
'
Public Sub CompareDoubleValueWithoutReturn(ByVal original As Long, ByVal Expected As Long)
    Assert.That original * 2, Iz.EqualTo(Expected)
End Sub

' This test accepts a value and returns the value * 2. The test is not performed within
' the method body. The framework uses an expectation value to compare the result of this
' method to. The expectation is set at the same time as the original value is defined in
' the code above.
'
Public Function CompareDoubleValueWithReturn(ByVal Value As Long) As Long
    CompareDoubleValueWithReturn = Value * 2
End Function


Public Function DivideValue(ByVal Denominator As Long) As Long
    DivideValue = 10 / Denominator
End Function
