VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestFixtureTests"
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
' Module: TestFixtureTests
'
Option Explicit

Public Sub Run_UsingSuccessfulTest_ReturnsSuccessTestResult()
    On Error GoTo errTrap
    
    Dim TestClass As New Fake1Test
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(TestClass)
    
    Dim Result As TestResult
    Set Result = Fixture.Run
    Debug.Assert Result.IsSuccess = True
    Exit Sub
    
errTrap:
    Debug.Assert False
End Sub

Public Sub Run_UsingFailingTest_ReturnsFailingTestResult()
    On Error GoTo errTrap
    
    Dim TestClass As New Stub1Test
    Call TestClass.Errors.Add(Constants.TEST1_METHOD, AssertCode.FailureCode)
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(TestClass)
    
    Dim Result As TestResult
    Set Result = Fixture.Run
    
    Debug.Assert Result.IsSuccess = False
    Exit Sub
    
errTrap:
    Debug.Assert False ' no error should be raised.
End Sub

Public Sub Run_WhenCalled_ReturnsTestResultContainingTestFixture()
    On Error GoTo errTrap
    
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(New FakeEmpty)
    Dim Result As TestResult
    Set Result = Fixture.Run
    
    Dim Actual As ITest
    Set Actual = Result.Test
    Debug.Assert Actual Is Fixture
    Exit Sub
    
errTrap:
    Debug.Assert False
End Sub

Public Sub Constructor_WithTestClass_StoresTestCases()
    On Error GoTo errTrap
    Dim TestClass As New Fake2Tests
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(TestClass)
    
    Dim Actual As Long
    Actual = Fixture.Tests.Count
    
    Debug.Assert Actual = 2
    Exit Sub
    
errTrap:
    Debug.Assert False
End Sub

Public Sub Constructor_WithTestClass_StoresTestCase()
    On Error GoTo errTrap
    Dim TestClass As New Fake1Test
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(TestClass)
    
    Dim Actual As String
    Actual = Fixture.Tests(0).Name
    Debug.Assert Actual = Constants.TEST1_METHOD
    Exit Sub
    
errTrap:
    Debug.Assert False
End Sub

Public Sub Run_WithTestCases_AddsTestCaseResultToFixtureResult()
    On Error GoTo errTrap
    Dim TestClass As New Fake1Test
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(TestClass)
    
    Dim Result As TestResult
    Set Result = Fixture.Run
    
    Dim Actual As Long
    Actual = Result.Results.Count
    Debug.Assert Actual = 1
    Exit Sub
    
errTrap:
    Debug.Assert False
End Sub

Public Sub Run_WithTestCases_AddsTestCaseResultToTestFixtureResult2()
    On Error GoTo errTrap
    Dim TestClass As New Fake1Test
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(TestClass)
    
    Dim Result As TestResult
    Set Result = Fixture.Run.Results(0)
    
    Dim Actual As Boolean
    Actual = Result.IsSuccess
    Debug.Assert Actual = True
    Exit Sub
    
errTrap:
    Debug.Assert False
End Sub

Public Sub Name_WhenCalled_ReturnsNameOfTestClass()
    Dim TestClass As New Fake1Test
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(TestClass)
    
    Dim Actual As String
    Actual = Fixture.Name
    
    Debug.Assert Actual = "Fake1Test"
End Sub

Public Sub Constructor_WhenCalled_SetsTestCaseParentToFixture()
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(New Fake1Test)
    
    Dim Actual As ITest
    Set Actual = Fixture.Tests(0).Parent
    
    Debug.Assert Actual Is Fixture
End Sub

Public Sub Termination_WhenCalled_SetsTestCaseParentToNothing()
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(New Fake1Test)
    
    Dim Child As ITest
    Set Child = Fixture.Tests(0)
    Set Fixture = Nothing
    
    Dim Actual As ITest
    Set Actual = Child.Parent
    
    Assert.IsTrue Actual Is Nothing, "Parent should be set to nothing."
End Sub

Public Sub FullName_WithParent_ReturnsParentNameAndSelfName()
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(New Fake1Test)
    Dim Suite As New TestSuite
    Call Suite.Add(Fixture)
    
    Dim Actual As String
    Actual = Fixture.FullName
    
    Assert.IsTrue Actual = "TestSuite.Fake1Test", "Wrong full name."
End Sub

Public Sub FullName_WhenCalledWithoutParent_ReturnsOnlyName()
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(New Fake1Test)
    
    Dim Actual As String
    Actual = Fixture.FullName
    
    Assert.IsTrue Actual = "Fake1Test", "Wrong fixture name."
End Sub

Public Sub Constructor_WithNothing_RaisesError()
    On Error GoTo errTrap
    
    Sim.NewTestFixture Nothing
    On Error GoTo 0
    Assert.Fail "An error should be raised."
    
errTrap:
    Dim Actual As ErrorInfo
    Set Actual = ErrorInfo.FromErr(Err)
    Dim Expected  As ErrorInfo
    Set Expected = Sim.NewErrorInfo(ErrorCode.ArgumentNull, "TestFixture.Init", "TestClass cannot be Nothing.")
    Assert.IsTrue Actual.Equals(Expected)
End Sub

Public Sub Constructor_UsingClassWithSetup_DoesNotIncludeSetupAsTest()
    Dim TestClass As New Mock2TestsWithSetup
    Dim Fixture As TestFixture
    
    Set Fixture = Sim.NewTestFixture(TestClass)
    
    Dim Actual As Long
    Actual = Fixture.TestCount
    Assert.IsTrue Actual = 2, "Wrong number of tests."
End Sub

Public Sub Run_UsingClassWithSetup_RunsSetupOnceBeforeEachTest()
    Dim TestClass As New Mock2TestsWithSetup
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(TestClass)
    
    Call Fixture.Run
    
    Dim Actual As CallTrace
    Set Actual = TestClass.Calls
    Dim Expected As New CallTrace
    Call Expected.Add(Constants.SETUP_METHOD)
    Call Expected.Add(Constants.TEST1_METHOD)
    Call Expected.Add(Constants.SETUP_METHOD)
    Call Expected.Add(Constants.TEST2_METHOD)
    Assert.IsTrue Actual.Equals(Expected), "Setup should be called twice."
End Sub

Public Sub Run_UsingClassWithTeardown_RunsTeardownOnceAfterEachTest()
    Dim Mock As New Mock2TestsWithTeardown
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Mock)
    
    Call Fixture.Run
    
    Dim Actual As CallTrace
    Set Actual = Mock.Calls
    Dim Expected As New CallTrace
    Call Expected.Add(Constants.TEST1_METHOD)
    Call Expected.Add(Constants.TEARDOWN_METHOD)
    Call Expected.Add(Constants.TEST2_METHOD)
    Call Expected.Add(Constants.TEARDOWN_METHOD)
    Assert.IsTrue Actual.Equals(Expected), "Teardown should be called twice."
End Sub

Public Sub Run_UsingClassWithFixtureSetup_RunsFixtureSetupOnce()
    Dim Mock As New Mock2TestsWithFixtureSetup
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Mock)
    
    Call Fixture.Run
    
    Dim Actual As CallTrace
    Set Actual = Mock.Calls
    Dim Expected As New CallTrace
    Call Expected.Add(Constants.FIXTURESETUP_METHOD)
    Call Expected.Add(Constants.TEST_ANY)
    Call Expected.Add(Constants.TEST_ANY)
    Assert.IsTrue Actual.Equals(Expected)
End Sub

Public Sub Run_UsingClassWithFixtureTeardown_RunsFixtureTeardownOnce()
    Dim Mock As New Mock2TestsWithFixtureTeardown
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Mock)
    
    Call Fixture.Run
    
    Dim Actual As CallTrace
    Set Actual = Mock.Calls
    Dim Expected As New CallTrace
    Call Expected.Add(Constants.TEST_ANY)
    Call Expected.Add(Constants.TEST_ANY)
    Call Expected.Add(Constants.FIXTURETEARDOWN_METHOD)
    Assert.IsTrue Actual.Equals(Expected)
End Sub

Public Sub Constructor_WhenCalled_StoresTestClass()
    Dim Fake As New Fake1Test
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Fake)
    
    Dim Actual As Object
    Set Actual = Fixture.TestClass
    
    Assert.IsTrue Actual Is Fake, "Wrong test class returned."
End Sub

Public Sub Sort_WhenCalled_SortsTestsAscending()
    Dim Fake As New Fake4TestsInDescendingOrder
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Fake)
    
    Call Fixture.Sort
    
    Dim Actual As TestList
    Set Actual = Fixture.Tests
    Dim Expected As New TestList
    Call Expected.Add(NewStubTest("ATest"))
    Call Expected.Add(NewStubTest("BTest"))
    Call Expected.Add(NewStubTest("CTest"))
    Call Expected.Add(NewStubTest("DTest"))
    Assert.IsTrue Actual.Equals(Expected), "Tests not sorted."
End Sub

Public Sub Run_WhenCalled_StoresAssertCountInResult()
    Dim Stub As New Stub2TestsWithFullAsserts
    Stub.FixtureSetupAssertCount = 3
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Stub)

    Dim Result As TestResult
    Set Result = Fixture.Run

    Dim Actual As Long
    Actual = Result.AssertCount
    Assert.AreEqual 3, Actual
End Sub

Public Sub Run_WhenCalled_StoresAssertCountInResult2()
    Dim Stub As New Stub2TestsWithFullAsserts
    Stub.FixtureTeardownAssertCount = 2
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Stub)

    Dim Result As TestResult
    Set Result = Fixture.Run

    Dim Actual As Long
    Actual = Result.AssertCount
    Assert.AreEqual 2, Actual
End Sub

Public Sub Run_WhenCalled_DoesNotStoreTestCaseAssertCounts()
    Dim Stub As New Stub2TestsWithFullAsserts
    Stub.Test1AssertCount = 2
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Stub)

    Dim Result As TestResult
    Set Result = Fixture.Run

    Dim Actual As Long
    Actual = Result.AssertCount
    Assert.AreEqual 0, Actual
End Sub

Public Sub Run_WithFailingFixtureSetup_DoesNotRunTests()
    Dim Mock As New Mock2TestsWithFixtureSetup
    Call Mock.Errors.Add(Constants.FIXTURESETUP_METHOD, 5)
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Mock)
    
    Call Fixture.Run
    
    Dim Actual As CallTrace
    Set Actual = Mock.Calls
    Dim Expected As New CallTrace
    Call Expected.Add(Constants.FIXTURESETUP_METHOD)
    Assert.IsTrue Actual.Equals(Expected)
End Sub

Public Sub Run_WithFailingFixtureSetup_ChildTestResultsAreFailure()
    Dim Stub As New Stub1TestWithFullContext
    Call Stub.Errors.Add(Constants.FIXTURESETUP_METHOD, 5)
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Stub)
    Dim Result As TestResult
    Set Result = Fixture.Run
    
    Dim ChildResult As TestResult
    Set ChildResult = Result.Results(0)
    
    Dim Actual As ResultState
    Actual = ChildResult.ResultState
    Assert.IsTrue Actual = FailureState
End Sub

Public Sub Run_WithEventListener_CallsListenerMethods()
    Dim Mock As New MockEventListener
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(New FakeEmpty)
    
    Call Fixture.Run(Mock)
    
    Dim Actual As CallTrace
    Set Actual = Mock.Calls
    Dim Expected As New CallTrace
    Call Expected.Add(Constants.TESTSUITESTARTED_LISTENERMETHOD)
    Call Expected.Add(Constants.TESTSUITEFINISHED_LISTENERMETHOD)
    Assert.IsTrue Actual.Equals(Expected)
End Sub

Public Sub Run_WithFailingFixtureSetup_SetsTestCaseSourceTestsToFailingResult()
    Dim Stub As New Stub1TestCaseSrcAndFixtureSetup
    Call Stub.Errors.Add("FixtureSetup", 5, , "FixtureSetup had an error.")
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Stub)
    
    Dim Result As TestResult
    Set Result = Fixture.Run.Results(0).Results(0)
    
    Dim Actual As ResultState
    Actual = Result.ResultState
    Assert.IsTrue Actual = FailureState
End Sub

Public Sub Run_WithRunStateSetToNotRunnable_DoesNotRunTests()
    Dim Mock As New Mock2TestsWithFullContext
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Mock)
    Fixture.RunState = NotRunnableRunState
    
    Call Fixture.Run
    
    Dim Actual As CallTrace
    Set Actual = Mock.Calls
    Dim Expected As New CallTrace
    Assert.IsTrue Actual.Equals(Expected)
End Sub

Public Sub Run_WithRunStateSetToNotRunnable_ReturnsResultSetToNotRunnable()
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(New Fake1Test)
    Fixture.RunState = NotRunnableRunState

    Dim Result As TestResult
    Set Result = Fixture.Run

    Assert.IsTrue Result.ResultState = NotRunnableState
End Sub

Public Sub Run_WithNonPassingFilter_DoesNotRun()
    Dim Stub As New StubFilter
    Stub.Matches = False
    Stub.Passes = False
    Dim Mock As New Mock2TestsWithFullContext
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Mock)
    
    Call Fixture.Run(Nothing, Stub)
    
    Dim Actual As CallTrace
    Set Actual = Mock.Calls
    Dim Expected As New CallTrace
    Assert.IsTrue Actual.Equals(Expected)
End Sub

Public Sub Run_WithNonPassingFilter_ReturnsInconclusiveResult()
    Dim Stub As New StubFilter
    Stub.Matches = False
    Stub.Passes = False
    Dim Mock As New Mock1Test
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Mock)
    
    Dim Result As TestResult
    Set Result = Fixture.Run(Nothing, Stub)
    
    Dim Actual As ResultState
    Actual = Result.ResultState
    Assert.IsTrue Actual = InconclusiveState
End Sub

Public Sub Run_WithNonPassingFilterAndNonRunnableState_ReturnsNonRunnableResult()
    Dim Stub As New StubFilter
    Stub.Matches = False
    Stub.Passes = False
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(New Fake1Test)
    Fixture.RunState = NotRunnableRunState
    
    Dim Result As TestResult
    Set Result = Fixture.Run(Nothing, Stub)
    
    Dim Actual As ResultState
    Actual = Result.ResultState
    Assert.IsTrue Actual = NotRunnableState
End Sub

Public Sub Run_WithRunStateSetToNotRunnable_CallsEventListenerMethods()
    Dim Mock As New MockEventListener
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(New Fake1Test)
    Fixture.RunState = NotRunnableRunState
    
    Call Fixture.Run(Mock)
    
    Dim Actual As CallTrace
    Set Actual = Mock.Calls
    Dim Expected As New CallTrace
    Call Expected.Add(Constants.TESTSUITESTARTED_LISTENERMETHOD)
    Call Expected.Add(Constants.TESTSUITEFINISHED_LISTENERMETHOD)
    Assert.IsTrue Actual.Equals(Expected)
End Sub

Public Sub Run_WhenCalledTwice_RunsTestsTwice()
    Dim Mock As New Mock1Test
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Mock)
    
    Call Fixture.Run
    Call Fixture.Run
    
    Dim Actual As CallTrace
    Set Actual = Mock.Calls
    Dim Expected As New CallTrace
    Call Expected.Add("Test1")
    Call Expected.Add("Test1")
    Assert.IsTrue Actual.Equals(Expected)
End Sub

Public Sub Run_WithIgnoreInFixtureSetup_RunsTests()
    Dim Mock As New Mock2TestsWithFixtureSetup
    Call Mock.Errors.Add(Constants.FIXTURESETUP_METHOD, AssertCode.IgnoreCode)
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Mock)
    
    Call Fixture.Run
    
    Dim Actual As CallTrace
    Set Actual = Mock.Calls
    Dim Expected As New CallTrace
    Call Expected.Add(Constants.FIXTURESETUP_METHOD)
    Call Expected.Add(Constants.TEST1_METHOD)
    Call Expected.Add(Constants.TEST2_METHOD)
    Assert.IsTrue Actual.Equals(Expected)
End Sub

Public Sub Run_WithIgnoreInFixtureTeardown_Continues()
    Dim Mock As New Mock2TestsWithFixtureTeardown
    Call Mock.Errors.Add(Constants.FIXTURETEARDOWN_METHOD, AssertCode.IgnoreCode)
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Mock)
    
    Dim Result As TestResult
    Set Result = Fixture.Run
    
    Dim Actual As ResultState
    Actual = Result.ResultState
    Assert.IsTrue Actual = SuccessState
End Sub


Private Function NewStubTest(ByVal Name As String) As ITest
    Dim Stub As New StubTest
    Stub.Name = Name
    Set NewStubTest = Stub
End Function






