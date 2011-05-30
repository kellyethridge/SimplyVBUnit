VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestCaseTests"
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
' Module: TestCaseTests
'
Option Explicit


Public Sub Constructor_WhenCalled_StoresName()
    Dim TestCase As TestCase
    Set TestCase = Sim.NewTestCase(Constants.TEST1_METHOD, MakeTestMethod)
    
    Dim Actual As String
    Actual = TestCase.Name
    
    Debug.Assert Actual = Constants.TEST1_METHOD
End Sub

Public Sub Constructor_WithEmptyName_RaisesError()
    On Error GoTo errTrap
    
    Sim.NewTestCase "", MakeTestMethod
    On Error GoTo 0
    Assert.Fail "An error should be raised."
    
errTrap:
    AssertError Err, ErrorCode.Argument, "TestCase.Init", "Name cannot be empty."
End Sub

Public Sub Constructor_WithNothingTestMethod_RaisesError()
    On Error GoTo errTrap
    
    Sim.NewTestCase "Name", Nothing
    On Error GoTo 0
    Assert.Fail "An error should be raised."
    
errTrap:
    AssertError Err, ErrorCode.ArgumentNull, "TestCase.Init", "TestMethod cannot be Nothing."
End Sub

Public Sub FullName_WhenCalled_ReturnsNameWithParentName()
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(New Fake1Test)
    Dim TestCase As TestCase
    Set TestCase = Fixture.Tests(0)

    Dim Actual As String
    Actual = TestCase.FullName
    
    Assert.IsTrue Actual = "Fake1Test.Test1", "Wrong full name."
End Sub

Public Sub FullName_WithNoParent_ReturnsName()
    Dim TestCase As TestCase
    Set TestCase = Sim.NewTestCase("Test", MakeTestMethod)
    Set TestCase.Parent = Nothing
    
    Dim Actual As String
    Actual = TestCase.FullName
    
    Assert.IsTrue Actual = "Test", "Wrong test name."
End Sub

Public Sub Run_WithFailingAssert_ReturnsFailingTestResult()
    Dim TestCase As TestCase
    Set TestCase = MakeFailingTestCase
    
    Dim Result As TestResult
    Set Result = TestCase.Run
    
    Assert.IsTrue Result.IsFailure, "Result should be failure."
End Sub

Public Sub Run_WithFailingAssert_ReturnsFailingTestResult2()
    Dim TestCase As TestCase
    Set TestCase = MakeFailingTestCase
    
    Dim Result As TestResult
    Set Result = TestCase.Run
    
    Dim Actual As String
    Actual = Result.Message
    Assert.IsTrue Actual = "Failed", "Result should contain message."
End Sub

Public Sub Run_WithPassingTest_ReturnsSuccessTestResult()
    Dim TestClass As New Stub1Test
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(TestClass)
    Dim TestCase As TestCase
    Set TestCase = Fixture.Tests(0)
    
    Dim Result As TestResult
    Set Result = TestCase.Run
    
    Dim Actual As Boolean
    Actual = Result.IsSuccess
    Assert.IsTrue Actual = True, "Result should be success."
End Sub

Public Sub Run_WithUnknownErrorInTest_ReturnsErrorTestResult()
    Dim TestClass As New Stub1Test
    TestClass.Errors.Add Constants.TEST1_METHOD, 1, , "unknown error"
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(TestClass)
    Dim TestCase As TestCase
    Set TestCase = Fixture.Tests(0)
    
    Dim Result As TestResult
    Set Result = TestCase.Run
    
    Dim Actual As Boolean
    Actual = Result.IsError
    Assert.IsTrue Actual = True, "Result should be error."
End Sub

Public Sub Run_WithUnknownErrorTest_ReturnsErrorTestResult2()
    Dim TestClass As New Stub1Test
    TestClass.Errors.Add Constants.TEST1_METHOD, 1, , "unknown error"
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(TestClass)
    Dim TestCase As TestCase
    Set TestCase = Fixture.Tests(0)
    
    Dim Result As TestResult
    Set Result = TestCase.Run
    
    Dim Actual As String
    Actual = Result.Message
    Assert.IsTrue Actual = "unknown error", "Wrong error message."
End Sub


Public Sub Run_WithSetupDefined_RunsSetup()
    Dim TestClass As New Mock2TestsWithSetup
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(TestClass)
    Dim TestCase As ITest
    Set TestCase = Fixture.Tests(0)
    
    TestCase.Run
    
    AssertCalls TestClass.Calls, Constants.SETUP_METHOD, Constants.TEST1_METHOD
End Sub

Public Sub Run_WithTeardownDefined_RunsTeardown()
    Dim TestClass As New Mock2TestsWithTeardown
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(TestClass)
    Dim TestCase As ITest
    Set TestCase = Fixture.Tests(0)
    
    TestCase.Run
    
    AssertCalls TestClass.Calls, Constants.TEST1_METHOD, Constants.TEARDOWN_METHOD
End Sub

Public Sub Run_WithErrorInSetup_SkipsTestRunsTeardown()
    Dim TestClass As New Mock1TestWithSetupTeardown
    TestClass.Errors.Add "Setup", AssertCode.FailureCode
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(TestClass)
    Dim TestCase As ITest
    Set TestCase = Fixture.Tests(0)
    
    TestCase.Run
        
    AssertCalls TestClass.Calls, Constants.SETUP_METHOD, Constants.TEARDOWN_METHOD
End Sub

Public Sub Run_WithAssertFailInTest_RunsTeardown()
    Dim Mock As New Mock1TestWithTeardown
    Mock.Errors.Add Constants.TEST1_METHOD, AssertCode.FailureCode
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Mock)
    Dim TestCase As ITest
    Set TestCase = Fixture.Tests(0)

    TestCase.Run
    
    AssertCalls Mock.Calls, Constants.TEST1_METHOD, Constants.TEARDOWN_METHOD
End Sub

Public Sub Run_WithAssertFailInTest_ReturnsFailingResult()
    Dim Mock As New Mock1TestWithTeardown
    Mock.Errors.Add Constants.TEST1_METHOD, AssertCode.FailureCode
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Mock)
    Dim TestCase As ITest
    Set TestCase = Fixture.Tests(0)

    Dim Result As TestResult
    Set Result = TestCase.Run
    
    Dim Actual As Boolean
    Actual = Result.IsFailure
    Assert.IsTrue Actual = True, "Result should be failure."
End Sub

Public Sub TestClass_WhenCalled_ReturnsTestClass()
    Dim Fake As New Fake1Test
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Fake)
    Dim Test As ITest
    Set Test = Fixture.Tests(0)
    
    Dim Actual As Object
    Set Actual = Test.TestClass
    
    Assert.IsTrue Actual Is Fake, "Wrong test class returned."
End Sub

Public Sub Run_ParameterizedTest_RunsTest()
    Dim Mock As New Mock1TestCaseSource
    Mock.Value1 = "arg1"
    Mock.Value2 = "arg2"
    Mock.Value3 = "arg3"
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Mock)
    Dim Test As ITest
    Set Test = Fixture.Tests(0)
    
    Test.Run
    
    AssertCalls Mock.Calls, "GetTestCases.Test1", "Test1.arg1.arg2.arg3"
End Sub

Public Sub Run_WithErrorInSetup_SetsResultToError()
    Dim Stub As New Stub1TestWithFullContext
    Stub.Errors.Add Constants.SETUP_METHOD, ErrorCode.Argument
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Stub)
    Dim Test As ITest
    Set Test = Fixture.Tests(0)
    
    Dim Actual As TestResult
    Set Actual = Test.Run
    
    Assert.IsTrue Actual.IsError, "Result should be error."
End Sub

Public Sub Run_WithErrorInSetup_SetsResultToError2()
    Dim Stub As New Stub1TestWithFullContext
    Stub.Errors.Add Constants.SETUP_METHOD, ErrorCode.Argument, , "In Setup"
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Stub)
    Dim Test As ITest
    Set Test = Fixture.Tests(0)
    
    Dim Actual As TestResult
    Set Actual = Test.Run
    
    Assert.IsTrue Actual.Message = "Error 5, In Setup", "Wrong result message."
End Sub

Public Sub Run_WithErrorInTeardown_SetsResultToError()
    Dim Stub As New Stub1TestWithFullContext
    Stub.Errors.Add Constants.TEARDOWN_METHOD, 5, , "In Teardown"
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Stub)
    Dim Test As ITest
    Set Test = Fixture.Tests(0)
    
    Dim Result As TestResult
    Set Result = Test.Run
    
    Dim Actual As Boolean
    Actual = Result.IsError
    Assert.IsTrue Actual = True, "Result should be error."
End Sub

Public Sub Run_WithErrorInTeardown_SetsResultToError2()
    Dim Stub As New Stub1TestWithFullContext
    Stub.Errors.Add Constants.TEARDOWN_METHOD, 5, , "In Teardown"
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Stub)
    Dim Test As ITest
    Set Test = Fixture.Tests(0)
    
    Dim Result As TestResult
    Set Result = Test.Run
    
    Dim Actual As String
    Actual = Result.Message
    Assert.IsTrue Actual = "Error 5, In Teardown", "Wrong result message."
End Sub

Public Sub Run_WithErrorInTest_RunsTeardown()
    Dim Mock As New Mock1TestWithTeardown
    Mock.Errors.Add Constants.TEST1_METHOD, 1
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Mock)
    Dim Test As ITest
    Set Test = Fixture.Tests(0)
    
    Dim Result As TestResult
    Set Result = Test.Run
    
    AssertCalls Mock.Calls, Constants.TEST1_METHOD, Constants.TEARDOWN_METHOD
End Sub

Public Sub Run_WhenCalled_StoresAssertCountInResult()
    Dim Stub As New Stub2TestsWithFullAsserts
    Stub.Test1AssertCount = 2
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Stub)
    Dim Test As ITest
    Set Test = Fixture.Tests(0)

    Dim Result As TestResult
    Set Result = Test.Run

    Dim Actual As Long
    Actual = Result.AssertCount
    Assert.AreEqual 2, Actual
End Sub

Public Sub Run_WhenCalled_StoresAssertCountInResult2()
    Dim Stub As New Stub2TestsWithFullAsserts
    Stub.Test2AssertCount = 3
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Stub)
    Dim Test As ITest
    Set Test = Fixture.Tests(1)

    Dim Result As TestResult
    Set Result = Test.Run

    Dim Actual As Long
    Actual = Result.AssertCount
    Assert.AreEqual 3, Actual
End Sub

Public Sub Run_WhenCalled_StoresAssertCountInResult3()
    Dim Stub As New Stub2TestsWithFullAsserts
    Stub.SetupAssertCount = 2
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Stub)
    Dim Test As ITest
    Set Test = Fixture.Tests(1)

    Dim Result As TestResult
    Set Result = Test.Run

    Dim Actual As Long
    Actual = Result.AssertCount
    Assert.AreEqual 2, Actual
End Sub

Public Sub Run_WhenCalled_StoresAssertCountInResult4()
    Dim Stub As New Stub2TestsWithFullAsserts
    Stub.TeardownAssertCount = 3
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Stub)
    Dim Test As ITest
    Set Test = Fixture.Tests(1)

    Dim Result As TestResult
    Set Result = Test.Run

    Dim Actual As Long
    Actual = Result.AssertCount
    Assert.AreEqual 3, Actual
End Sub

Public Sub Run_WhenCalled_DoesNotStoreFixtureAssertCounts()
    Dim Stub As New Stub2TestsWithFullAsserts
    Stub.FixtureSetupAssertCount = 1
    Stub.FixtureTeardownAssertCount = 2
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Stub)
    Dim Test As ITest
    Set Test = Fixture.Tests(0)

    Dim Result As TestResult
    Set Result = Test.Run

    Dim Actual As Long
    Actual = Result.AssertCount
    Assert.AreEqual 0, Actual
End Sub

Public Sub Run_WithTestCaseSource_RunsSetup()
    Dim Mock As New Mock2TestCaseSourceFullContext
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Mock)
    Dim Test As ITest
    Set Test = Fixture.Tests(0)
    
    Test.Run
    
    AssertCalls Mock.Calls, Constants.SETUP_METHOD, Constants.TEST_ANY, Constants.TEARDOWN_METHOD
End Sub

Public Sub Run_WithEventListener_CallsListenerMethods()
    Dim Mock As New MockEventListener
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(New Fake1Test)
    Dim Test As ITest
    Set Test = Fixture.Tests(0)
    
    Test.Run Mock
    
    AssertCalls Mock.Calls, Constants.TESTCASESTARTED_LISTENERMETHOD, Constants.TESTCASEFINISHED_LISTENERMETHOD
End Sub

Public Sub Run_WithAssertPass_ReturnsSuccess()
    Dim Stub As New Stub1Test
    Stub.Errors.Add Constants.TEST1_METHOD, AssertCode.SuccessCode
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Stub)
    Dim Test As ITest
    Set Test = Fixture.Tests(0)
    
    Dim Result As TestResult
    Set Result = Test.Run
    
    Dim Actual As ResultState
    Actual = Result.ResultState
    Dim Expected As ResultState
    Expected = SuccessState
    Assert.IsTrue Actual = Expected
End Sub

Public Sub Run_WithAssertPassMessage_IncludesMessageInResult()
    Const SUCCESS_MESSAGE   As String = "success message"
    
    Dim Stub As New Stub1Test
    Stub.Errors.Add Constants.TEST1_METHOD, AssertCode.SuccessCode, , SUCCESS_MESSAGE
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Stub)
    Dim Test As ITest
    Set Test = Fixture.Tests(0)
    
    Dim Result As TestResult
    Set Result = Test.Run

    Dim Actual As String
    Actual = Result.Message
    Assert.IsTrue Actual = SUCCESS_MESSAGE
End Sub

Public Sub Run_WithUnknownError_CallsListenerMethod()
    Const ERR_CODE As Long = 5
    
    Dim Mock As New MockEventListener
    Dim Stub As New Stub1Test
    Stub.Errors.Add Constants.TEST1_METHOD, ERR_CODE
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Stub)
    Dim Test As ITest
    Set Test = Fixture.Tests(0)
    
    Test.Run Mock
    
    AssertCalls Mock.Calls, Constants.TESTCASESTARTED_LISTENERMETHOD, _
                            Constants.UNHANDLEDERROR_LISTENERMETHOD, _
                            Constants.TESTCASEFINISHED_LISTENERMETHOD
End Sub

Public Sub Run_WithUnknownError_CallsListenerMethod2()
    Const ERR_CODE      As Long = 9
    Const ERR_SOURCE    As String = "source"
    Const ERR_DESC      As String = "description"
    
    Dim Mock As New MockEventListener
    Dim Stub As New Stub1Test
    Stub.Errors.Add Constants.TEST1_METHOD, ERR_CODE, ERR_SOURCE, ERR_DESC
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Stub)
    Dim Test As ITest
    Set Test = Fixture.Tests(0)
    
    Test.Run Mock
    
    Dim Actual As ErrorInfo
    Set Actual = Mock.ErrorInfo
    Dim Expected As ErrorInfo
    Set Expected = Sim.NewErrorInfo(ERR_CODE, ERR_SOURCE, ERR_DESC)
    
    Assert.IsTrue Actual.Equals(Expected)
End Sub

Public Sub Run_WithRunStateSetToNotRunnable_DoesNotExecuteTest()
    Dim Mock As New Mock1Test
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Mock)
    Dim Test As TestCase
    Set Test = Fixture.Tests(0)
    Test.RunState = NotRunnableRunState
    
    Test.Run
    
    Dim Actual As CallTrace
    Set Actual = Mock.Calls
    Dim Expected As New CallTrace
    Assert.IsTrue Actual.Equals(Expected)
End Sub

Public Sub Run_WithRunStateSetToNotRunnable_ReturnsNotRunnableResult()
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(New Fake1Test)
    Dim Test As TestCase
    Set Test = Fixture.Tests(0)
    Test.RunState = NotRunnableRunState
    
    Dim Result As TestResult
    Set Result = Test.Run
    
    Dim Actual As ResultState
    Actual = Result.ResultState
    Assert.IsTrue Actual = NotRunnableState
End Sub

Public Sub Run_WhenCalled_BeginsNewContextForEachTestCase()
    Dim FakeOut As New FakeTextWriter
    Dim Mock    As New Mock3TestsWithWriters
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Mock)
    Dim Runner  As TestRunner
    Set Runner = Sim.NewTestRunner(Fixture)

    With New TestContext
        Set .Out = FakeOut
        Runner.Run
    End With
    
    Assert.IsFalse Mock.Test1Writer Is FakeOut
    Assert.IsTrue Mock.Test2Writer Is FakeOut
    Assert.IsTrue Mock.Test3Writer Is FakeOut
End Sub

Public Sub Run_WithRunStateSetToNotRunnable_CallsListenerMethods()
    Dim Mock As New MockEventListener
    Dim Test As TestCase
    Set Test = Sim.NewTestCase(Constants.TEST1_METHOD, MakeTestMethod)
    Test.RunState = NotRunnableRunState
    
    Test.Run Mock
    
    AssertCalls Mock.Calls, Constants.TESTCASESTARTED_LISTENERMETHOD, _
                            Constants.TESTCASEFINISHED_LISTENERMETHOD
End Sub

Public Sub Run_WithIgnoredTest_ReturnsIgnoredResult()
    Dim Stub As New Stub1Test
    Stub.Errors.Add Constants.TEST1_METHOD, AssertCode.IgnoreCode
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Stub)
    Dim Test As TestCase
    Set Test = Fixture.Tests(0)
    
    Dim Result As TestResult
    Set Result = Test.Run
    
    Dim Actual As ResultState
    Actual = Result.ResultState
    Assert.IsTrue Actual = IgnoredState
End Sub

Public Sub Run_WithIgnoredTest_ReturnsIgnoreReason()
    Dim Stub As New Stub1Test
    Stub.Errors.Add Constants.TEST1_METHOD, AssertCode.IgnoreCode, , "Ignore message"
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Stub)
    Dim Test As TestCase
    Set Test = Fixture.Tests(0)
    
    Dim Result As TestResult
    Set Result = Test.Run
    
    Dim Actual As String
    Actual = Result.Message
    Assert.IsTrue Actual = "Ignore message"
End Sub

Public Sub Run_WithIgnoredInSetup_RunsTest()
    Dim Mock As New Mock2TestsWithSetup
    Mock.Errors.Add Constants.SETUP_METHOD, AssertCode.IgnoreCode
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Mock)
    
    Fixture.Run
    
    AssertCalls Mock.Calls, Constants.SETUP_METHOD, _
                            Constants.TEST1_METHOD, _
                            Constants.SETUP_METHOD, _
                            Constants.TEST2_METHOD
End Sub

Public Sub Run_WithIgnoredInTeardown_DoesNothing()
    Dim Mock As New Mock2TestsWithTeardown
    Mock.Errors.Add Constants.TEARDOWN_METHOD, AssertCode.IgnoreCode
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Mock)
    
    Fixture.Run
    
    AssertCalls Mock.Calls, Constants.TEST1_METHOD, _
                            Constants.TEARDOWN_METHOD, _
                            Constants.TEST2_METHOD, _
                            Constants.TEARDOWN_METHOD
End Sub

Public Sub CountTestCases_WithPassingFilter_ReturnsOne()
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(New Fake1Test)
    Dim Test As TestCase
    Set Test = Fixture.Tests(0)
    
    Dim Actual As Long
    Actual = Test.CountTestCases(TestFilter.EmptyFilter)
    
    Const Expected As Long = 1
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub CountTestCases_WithNonpassingFilter_ReturnsZero()
    Dim Stub As New StubFilter
    Stub.Passes = False
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(New Fake1Test)
    Dim Test As TestCase
    Set Test = Fixture.Tests(0)
    
    Dim Actual As Long
    Actual = Test.CountTestCases(Stub)
    
    Const Expected As Long = 0
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Categories_WhenDefinedInTestName_AddsToList()
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(New Mock1TestWithCategory)
    Dim Test As ITest
    Set Test = Fixture.Tests(0)
    
    Dim Actual As Boolean
    Actual = Test.Categories.Contains("DB")
    
    Assert.IsTrue Actual
End Sub

Public Sub Name_WithCategoryDefined_RemovesCategoryFromName()
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(New Mock1TestWithCategory)
    Dim Test As ITest
    Set Test = Fixture.Tests(0)

    Dim Actual As String
    Actual = Test.Name
    
    Assert.That Actual, Iz.EqualTo("Test1")
End Sub

Public Sub Name_WithMultipleCategoriesDefined_AddsToList()
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(New Mock1TestWith2Categories)
    Dim Test As ITest
    Set Test = Fixture.Tests(0)
    
    Dim Actual1 As Boolean
    Dim Actual2 As Boolean
    Actual1 = Test.Categories.Contains("DB")
    Actual2 = Test.Categories.Contains("General")
    
    Assert.IsTrue Actual1
    Assert.IsTrue Actual2
End Sub

Public Sub Run_UsingTestWithExpectedError_ReturnsSuccess()
    Dim Stub As New Stub1TestThrows
    Stub.Number = 5
    Stub.Errors.Add Constants.TEST1_METHOD, 5
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Stub)
    Dim Test As ITest
    Set Test = Fixture.Tests(0)

    Dim Result As TestResult
    Set Result = Test.Run

    Assert.IsTrue Result.IsSuccess
End Sub

Public Sub Run_WithExpectedErrorWithWrongErrorRaised_ReturnsFailure()
    Dim Stub As New Stub1TestThrows
    Stub.Number = 5
    Stub.Errors.Add Constants.TEST1_METHOD, 6
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Stub)
    Dim Test As ITest
    Set Test = Fixture.Tests(0)

    Dim Result As TestResult
    Set Result = Test.Run

    Assert.IsTrue Result.IsFailure
End Sub

Public Sub Run_UsingTestCaseSourceWithExpectedErrorRaised_ReturnsSuccess()
    Dim Stub As New Stub1TestCaseSourceThrows
    Stub.Number = 5
    Stub.Errors.Add Constants.TEST1_METHOD, 5
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Stub)
    Dim Test As ITest
    Set Test = Fixture.Tests(0)
    
    Dim Result As TestResult
    Set Result = Test.Run
    
    Assert.IsTrue Result.IsSuccess
End Sub

Public Sub Run_WithExpectedErrorWithNoErrorRaised_ReturnsFailure()
    Dim Stub As New Stub1TestThrows
    Stub.Number = 5
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Stub)
    Dim Test As ITest
    Set Test = Fixture.Tests(0)
    
    Dim Result As TestResult
    Set Result = Test.Run
    
    Dim Actual As Boolean
    Actual = Result.IsFailure
    Assert.IsTrue Actual = True
End Sub

Public Sub Run_UsingTestCaseSourceWithExpect_UsesConstraint()
    Dim Mock As New Mock1TestCaseSourceExpect
    Mock.ParameterValue = 5
    Mock.Constraint.Matches = True
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Mock)
    Dim Test As ITest
    Set Test = Fixture.Tests(0)
    
    Test.Run
    
    AssertCalls Mock.Constraint.Calls, "Resolve", "Matches"
End Sub

Public Sub Run_UsingTestCaseSourceWithExpect_UsesConstraint2()
    Dim Mock As New Mock1TestCaseSourceExpect
    Mock.ParameterValue = 5
    Mock.Constraint.Matches = True
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Mock)
    Dim Test As ITest
    Set Test = Fixture.Tests(0)
    
    Test.Run
    
    Dim Actual As Variant
    Actual = Mock.Constraint.Actual
    Assert.That Actual, Iz.EqualTo(5)
End Sub



Private Function MakeFailingTestCase() As TestCase
    Dim TestClass As New Stub1Test
    TestClass.Errors.Add Constants.TEST1_METHOD, AssertCode.FailureCode, , "Failed"
    
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(TestClass)

    Set MakeFailingTestCase = Fixture.Tests(0)
End Function

Private Function MakeTestMethod() As ITestMethod
    Dim TestClass As New Fake1Test
    Dim Member As MemberInfo
    Set Member = tli.InterfaceInfoFromObject(TestClass).GetMember(Constants.TEST1_METHOD)
    Set MakeTestMethod = Sim.NewSimpleTestMethod(TestClass, Member)
End Function
