VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestSuiteTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Copyright 2009 Kelly Ethridge
'
' Licensed under the Apache License, Version 2.0 (the "License");
' you may not use this file except in compliance with the License.
' You may obtain a copy of the License at
'
'     http://www.apache.org/licenses/LICENSE-2.0
'
' Unless required by applicable law or agreed to in writing, software
' distributed under the License is distributed on an "AS IS" BASIS,
' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
' See the License for the specific language governing permissions and
' limitations under the License.
'
' Module: TestSuiteTests
'
Option Explicit

Public Sub Constructor_WhenCalledWithDefaults_SetsNameToDefault()
    Dim Suite As New TestSuite
    
    Dim Actual As String
    Actual = Suite.Name
    
    Debug.Assert Actual = "TestSuite"
End Sub

Public Sub Run_WithSuccessfulTest_ReturnsSuccessfulResult()
    Dim Result As TestResult
    Dim Suite As New TestSuite
    Call Suite.Add(New Fake1Test)
    
    Set Result = Suite.Run
    
    Dim Actual As Boolean
    Actual = Result.IsSuccess
    Debug.Assert Actual = True
End Sub

Public Sub Run_WithFailingTest_ReturnsFailingResult()
    On Error GoTo errTrap
    
    Dim Result As TestResult
    Dim Suite As New TestSuite
    Dim TestClass As New Stub1Test
    Call TestClass.Errors.Add("Test1", AssertCode.FailureCode)
    Call Suite.Add(TestClass)
    
    Set Result = Suite.Run
    
    Dim Actual As Boolean
    Actual = Result.IsFailure
    Debug.Assert Actual = True
    Exit Sub
    
errTrap:
    Debug.Assert False
End Sub

Public Sub Add_WithITest_SetsParent()
    Dim Test As ITest
    Set Test = New FakeTest
    Dim Suite As New TestSuite
    
    Call Suite.Add(Test)
    
    Dim Actual As ITest
    Set Actual = Test.Parent
    Debug.Assert Actual Is Suite
End Sub

Public Sub Add_WithTestClass_SetsParent()
    Dim TestClass As New Fake1Test
    Dim Suite As New TestSuite
    
    Call Suite.Add(TestClass)
    
    Dim Actual As ITest
    Set Actual = Suite.Tests(0).Parent
    Debug.Assert Actual Is Suite
End Sub

Public Sub Termination_WhenCalled_SetsChildParentToNothing()
    Dim Parent As New TestSuite
    Dim Child As New TestSuite
    Call Parent.Add(Child)
    
    Set Parent = Nothing
    
    Dim Actual As ITest
    Set Actual = Child.Parent
    Assert.IsTrue Actual Is Nothing, "Parent should be Nothing."
End Sub

Public Sub Constructor_WhenCalledWithoutName_StoresDefaultName()
    Dim Suite As New TestSuite
    
    Dim Actual As String
    Actual = Suite.Name
    
    Assert.IsTrue Actual = "TestSuite", "Wrong suite name."
End Sub

Public Sub FullName_WhenCalledWithoutParent_ReturnsNameOnly()
    Dim Suite As New TestSuite
    
    Dim Actual As String
    Actual = Suite.FullName
    
    Assert.IsTrue Actual = "TestSuite", "Wrong full name."
End Sub

Public Sub Constructor_WithName_StoresName()
    Dim Suite As TestSuite
    
    Set Suite = Sim.NewTestSuite("Name")
    
    Dim Actual As String
    Actual = Suite.Name
    Assert.IsTrue Actual = "Name", "Wrong suite name."
End Sub

Public Sub FullName_WhenCalledWithParent_IncludesParentName()
    Dim Parent As New TestSuite
    Dim Child As New TestSuite
    Call Parent.Add(Child)
    
    Dim Actual As String
    Actual = Child.FullName
    
    Assert.IsTrue Actual = "TestSuite.TestSuite", "Wrong full name."
End Sub

Public Sub Add_WithNothing_RaisesError()
    On Error GoTo errTrap
    Dim Suite As New TestSuite
    
    Call Suite.Add(Nothing)
    On Error GoTo 0
    Assert.Fail "An error should be raised."
    
errTrap:
    Call AssertError(Err, ErrorCode.ArgumentNull, "TestSuite.Add", "TestClass cannot be Nothing.")
End Sub

Public Sub TestClass_WhenCalled_ReturnsNothing()
    Dim Suite As New TestSuite
    
    Dim Actual As Object
    Set Actual = Suite.TestClass
    
    Assert.IsTrue Actual Is Nothing, "Should be Nothing."
End Sub

Public Sub Sort_WhenCalled_SortsTests()
    Dim Suite As New TestSuite
    Call Suite.Add(NewStubTest("DTest"))
    Call Suite.Add(NewStubTest("CTest"))
    Call Suite.Add(NewStubTest("BTest"))
    Call Suite.Add(NewStubTest("ATest"))
    
    Call Suite.Sort
    
    Dim Actual As TestList
    Set Actual = Suite.Tests
    Dim Expected As New TestList
    Call Expected.Add(NewStubTest("ATest"))
    Call Expected.Add(NewStubTest("BTest"))
    Call Expected.Add(NewStubTest("CTest"))
    Call Expected.Add(NewStubTest("DTest"))
    Assert.IsTrue Actual.Equals(Expected), "Tests not sorted."
End Sub

Public Sub Sort_WhenCalled_ForwardsSortToChildTests()
    Dim Suite As New TestSuite
    Call Suite.Add(New Fake4TestsInDescendingOrder)
    
    Call Suite.Sort
    
    Dim Fixture As ITest
    Set Fixture = Suite.Tests(0)
    Dim Actual As TestList
    Set Actual = Fixture.Tests
    Dim Expected As New TestList
    Call Expected.Add(NewStubTest("ATest"))
    Call Expected.Add(NewStubTest("BTest"))
    Call Expected.Add(NewStubTest("CTest"))
    Call Expected.Add(NewStubTest("DTest"))
    Assert.IsTrue Actual.Equals(Expected), "Tests not sorted."
End Sub

Public Sub Run_WithEventListener_CallsListenerMethods()
    Dim Mock As New MockEventListener
    Dim Suite As New TestSuite
    
    Call Suite.Run(Mock)
    
    Call AssertCalls(Mock.Calls, Constants.TESTSUITESTARTED_LISTENERMETHOD, _
                                 Constants.TESTSUITEFINISHED_LISTENERMETHOD)
End Sub

Public Sub Run_WithRunStateSetToNotRunnable_DoesNotRunTests()
    Dim Mock As New Mock1Test
    Dim Suite As New TestSuite
    Call Suite.Add(Mock)
    Suite.RunState = NotRunnableRunState
    
    Call Suite.Run
    
    Call AssertNoCalls(Mock.Calls)
End Sub

Public Sub Run_WithRunStateSetToNotRunnable_ReturnsResultSetToNotRunnable()
    Dim Suite As New TestSuite
    Call Suite.Add(New FakeTest)
    Suite.RunState = NotRunnableRunState
    
    Dim Result As TestResult
    Set Result = Suite.Run
    
    Dim Actual As ResultState
    Actual = Result.ResultState
    Assert.IsTrue Actual = NotRunnableState
End Sub

Public Sub Run_WithNonPassingFilter_DoesNotRun()
    Dim Stub As New StubFilter
    Stub.Matches = False
    Stub.Passes = False
    Dim Mock As New Mock2TestsWithFullContext
    Dim Suite As New TestSuite
    Call Suite.Add(Mock)
    
    Call Suite.Run(Nothing, Stub)
    
    Call AssertNoCalls(Mock.Calls)
End Sub

Public Sub Run_WithNonPassingFilter_ReturnsInconclusiveResult()
    Dim Stub As New StubFilter
    Stub.Matches = False
    Stub.Passes = False
    Dim Mock As New Mock1Test
    Dim Suite As New TestSuite
    Call Suite.Add(Mock)

    Dim Result As TestResult
    Set Result = Suite.Run(Nothing, Stub)

    Dim Actual As ResultState
    Actual = Result.ResultState
    Assert.IsTrue Actual = InconclusiveState
End Sub

Public Sub Run_WithNonPassingFilterAndNonRunnableState_ReturnsNonRunnableResult()
    Dim Stub As New StubFilter
    Stub.Matches = False
    Stub.Passes = False
    Dim Suite As New TestSuite
    Call Suite.Add(New Fake1Test)
    Suite.RunState = NotRunnableRunState

    Dim Result As TestResult
    Set Result = Suite.Run(Nothing, Stub)

    Dim Actual As ResultState
    Actual = Result.ResultState
    Assert.IsTrue Actual = NotRunnableState
End Sub

Public Sub Run_WithRunStateSetToNotRunnable_CallsEventListenerMethods()
    Dim Mock As New MockEventListener
    Dim Suite As New TestSuite
    Call Suite.Add(New Fake1Test)
    Suite.RunState = NotRunnableRunState

    Call Suite.Run(Mock)

    Call AssertCalls(Mock.Calls, Constants.TESTSUITESTARTED_LISTENERMETHOD, _
                                 Constants.TESTSUITEFINISHED_LISTENERMETHOD)
End Sub

Public Sub CountTestCases_WithPassingFilter_ReturnsFour()
    Dim Suite As New TestSuite
    Call Suite.Add(New Fake2Tests)
    Call Suite.Add(New Fake2Tests)
    
    Dim Actual As Long
    Actual = Suite.CountTestCases(TestFilter.EmptyFilter)
    
    Const Expected As Long = 4
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub CountTestCases_WithNonpassingFilter_ReturnsZero()
    Dim Stub As New StubFilter
    Stub.Passes = False
    Dim Suite As New TestSuite
    Call Suite.Add(New Fake2Tests)
    Call Suite.Add(New Fake2Tests)
    
    Dim Actual As Long
    Actual = Suite.CountTestCases(Stub)
    
    Const Expected As Long = 0
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub CountTestCases_WithPartialPassingFilter_ReturnsTwo()
    Dim Suite As New TestSuite
    Call Suite.Add(New Fake2Tests)
    Call Suite.Add(New Fake2Tests)
    Dim Filter As New FullNameFilter
    Call Filter.Add(Suite.Tests(0).Tests(0).FullName)
    
    Dim Actual As Long
    Actual = Suite.CountTestCases(Filter)
    
    Const Expected As Long = 2
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Categories_WhenDefinedInTestName_AddsToList()
    Dim Suite As TestSuite
    Set Suite = Sim.NewTestSuite("DB__Suite")
    
    Dim Actual As Boolean
    Actual = Suite.Categories.Contains("DB")
    
    Assert.IsTrue Actual
End Sub

Public Sub Name_WithCategoryDefined_RemovesCategoryFromName()
    Dim Suite As TestSuite
    Set Suite = Sim.NewTestSuite("DB__Suite")

    Dim Actual As String
    Actual = Suite.Name

    Assert.That Actual, Iz.EqualTo("Suite")
End Sub

Public Sub Name_WithMultipleCategoriesDefined_AddsToList()
    Dim Suite As TestSuite
    Set Suite = Sim.NewTestSuite("DB_General__Suite")

    Dim Actual1 As Boolean
    Dim Actual2 As Boolean
    Actual1 = Suite.Categories.Contains("DB")
    Actual2 = Suite.Categories.Contains("General")

    Assert.IsTrue Actual1
    Assert.IsTrue Actual2
End Sub


Private Function NewStubTest(ByVal Name As String) As ITest
    Dim Stub As New StubTest
    Stub.Name = Name
    Set NewStubTest = Stub
End Function








