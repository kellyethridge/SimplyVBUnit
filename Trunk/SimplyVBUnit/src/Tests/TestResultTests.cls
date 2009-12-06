VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestResultTests"
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
' Module: TestResultTests
'
Option Explicit
Implements ITestCaseSource

Public Sub Constructor_WithTest_StoresTest()
    On Error GoTo errTrap
    Dim Fake As New FakeTest
    Dim Result As TestResult
    
    Set Result = Sim.NewTestResult(Fake)
    
    Dim Actual As ITest
    Set Actual = Result.Test
    Debug.Assert Actual Is Fake
    Exit Sub
    
errTrap:
    Debug.Assert False
End Sub

Public Sub SetResult_WithInconclusive_IsSuccessReturnsFalse()
    Dim Result As TestResult
    Set Result = MakeTestResult
    
    Call Result.SetResult(InconclusiveState)
    
    Dim Actual As Boolean
    Actual = Result.IsSuccess
    Debug.Assert Actual = False
End Sub

Public Sub SetResult_WithInconclusive_IsFailureReturnsFalse()
    Dim Result As TestResult
    Set Result = MakeTestResult
    
    Call Result.SetResult(InconclusiveState)
    
    Dim Actual As Boolean
    Actual = Result.IsFailure
    Debug.Assert Actual = False
End Sub

Public Sub SetResult_WithSuccessState_IsSuccessReturnsTrue()
    Dim Result As TestResult
    Set Result = MakeTestResult
    
    Call Result.SetResult(SuccessState)
    
    Dim Actual As Boolean
    Actual = Result.IsSuccess
    Debug.Assert Actual = True
End Sub

Public Sub SetResult_WithFailureState_IsFailureReturnsTrue()
    Dim Result As TestResult
    Set Result = MakeTestResult
    
    Call Result.SetResult(FailureState)
    
    Dim Actual As Boolean
    Actual = Result.IsFailure
    Debug.Assert Actual = True
End Sub

Public Sub SetResult_WithErrorState_IsErrorReturnsTrue()
    Dim Result As TestResult
    Set Result = MakeTestResult
    
    Call Result.SetResult(ErrorState)
    
    Dim Actual As Boolean
    Actual = Result.IsError
    Debug.Assert Actual = True
End Sub

Public Sub Error_WhenCalled_SetsResultState()
    Dim Result As TestResult
    Set Result = MakeTestResult
    
    Call Result.Error
    
    Dim Actual As ResultState
    Actual = Result.ResultState
    Debug.Assert Actual = ErrorState
End Sub

Public Sub Error_WithMessage_StoresMessage()
    Dim Result As TestResult
    Set Result = MakeTestResult
    
    Call Result.Error("Error")
    
    Dim Actual As String
    Actual = Result.Message
    Debug.Assert Actual = "Error"
End Sub

Public Sub Success_WhenCalled_SetsResultState()
    On Error GoTo errTrap
    Dim Result As TestResult
    Set Result = Sim.NewTestResult(New FakeTest)
    
    Call Result.Success
    
    Dim Actual As ResultState
    Actual = Result.ResultState
    Debug.Assert Actual = SuccessState
    Exit Sub
    
errTrap:
    Debug.Assert False
End Sub

Public Sub Success_WithMessage_StoresMessage()
    Dim Result As TestResult
    Set Result = Sim.NewTestResult(New FakeTest)
    
    Call Result.Success("message")
    
    Dim Actual As String
    Actual = Result.Message
    Debug.Assert Actual = "message"
End Sub

Public Sub Failure_WhenCalled_SetsResultState()
    On Error GoTo errTrap
    Dim Result As TestResult
    Set Result = Sim.NewTestResult(New FakeTest)
    
    Call Result.Failure
    
    Dim Actual As ResultState
    Actual = Result.ResultState
    Debug.Assert Actual = FailureState
    Exit Sub
    
errTrap:
    Debug.Assert False
End Sub

Public Sub Success_AfterFailureIsCalled_SetsFailureToFalse()
    On Error GoTo errTrap
    Dim Result As TestResult
    Set Result = Sim.NewTestResult(New FakeTest)
    
    Call Result.Failure
    Call Result.Success
    
    Dim Actual As Boolean
    Actual = Result.IsFailure
    Debug.Assert Actual = False
    Exit Sub
    
errTrap:
    Debug.Assert False
End Sub

Public Sub Failure_AfterSuccessIsCalled_SetsSuccessToFalse()
    On Error GoTo errTrap
    Dim Result As TestResult
    Set Result = Sim.NewTestResult(New FakeTest)
    
    Call Result.Success
    Call Result.Failure
    
    Dim Actual As Boolean
    Actual = Result.IsSuccess
    Debug.Assert Actual = False
    Exit Sub
    
errTrap:
    Debug.Assert False
End Sub

Public Sub AddResult_WithResult_StoresResult()
    On Error GoTo errTrap
    Dim Parent As TestResult
    Set Parent = Sim.NewTestResult(New FakeTest)
    Dim Child As TestResult
    Set Child = Sim.NewTestResult(New FakeTest)
    
    Call Parent.AddResult(Child)
    
    Dim Actual As TestResult
    Set Actual = Parent.Results(0)
    Debug.Assert Actual Is Child
    Exit Sub
    
errTrap:
    Debug.Assert False
End Sub

Public Sub AddResult_AddFailingResult_SetsParentResultToFailing()
    Dim Parent As TestResult
    Set Parent = Sim.NewTestResult(New FakeTest)
    Dim Child As TestResult
    Set Child = Sim.NewTestResult(New FakeTest)
    
    Call Child.Failure
    Call Parent.AddResult(Child)
    
    Dim Actual As Boolean
    Actual = Parent.IsFailure
    Debug.Assert Actual = True
End Sub

Public Sub Failure_WithMessage_StoresMessage()
    Dim Result As TestResult
    Set Result = Sim.NewTestResult(New FakeTest)
    
    Call Result.Failure("failure message")
    
    Dim Actual As String
    Actual = Result.Message
    Assert.IsTrue Actual = "failure message"
End Sub

Public Sub Constructor_WhenCalled_SetsResultStateToInconclusive()
    Dim Result As TestResult
    Set Result = Sim.NewTestResult(New FakeTest)
    
    Dim Actual As ResultState
    Actual = Result.ResultState
    
    Assert.IsTrue Actual = InconclusiveState, "Wrong result state."
End Sub

Public Sub Name_WhenCalled_ReturnsTestName()
    Const TEST_NAME As String = "Test Name"
    Dim Stub As New StubTest
    Stub.Name = TEST_NAME
    Dim Result As TestResult
    Set Result = Sim.NewTestResult(Stub)
    
    Dim Actual As String
    Actual = Result.Name
    
    Assert.IsTrue Actual = TEST_NAME, "Wrong test name."
End Sub

Public Sub FullName_WhenCalled_ReturnsTestFullName()
    Const FULL_NAME As String = "Test FullName"
    Dim Stub As New StubTest
    Stub.FullName = FULL_NAME
    Dim Result As TestResult
    Set Result = Sim.NewTestResult(Stub)
    
    Dim Actual As String
    Actual = Result.FullName
    
    Assert.IsTrue Actual = FULL_NAME, "Wrong test full name."
End Sub

Public Sub HasResults_WithNoChildResults_ReturnsFalse()
    Dim Result As TestResult
    Set Result = Sim.NewTestResult(New FakeTest)
    
    Dim Actual As Boolean
    Actual = Result.HasResults
    
    Assert.IsTrue Actual = False, "Should not have results."
End Sub

Public Sub HasResults_WithChildResult_ReturnsTrue()
    Dim Parent As TestResult
    Set Parent = MakeTestResult
    Dim Child As TestResult
    Set Child = MakeTestResult
    Call Parent.AddResult(Child)
    
    Dim Actual As Boolean
    Actual = Parent.HasResults
    
    Assert.IsTrue Actual = True, "Should have results."
End Sub

Public Sub AddResult_WithNothing_RaisesError()
    Dim Result As TestResult
    Set Result = MakeTestResult
    
    On Error GoTo errTrap
    Call Result.AddResult(Nothing)
    On Error GoTo 0
    
errTrap:
    Call AssertError(Err, ErrorCode.ArgumentNull, "TestResult.AddResult", "Result cannot be Nothing.")
End Sub

Public Sub SetResult_WithFailureSite_StoresFailureSite()
    Dim Result As TestResult
    Set Result = MakeTestResult
    
    Call Result.SetResult(FailureState, , FailureSite.SetupSite)
    
    Dim Actual As FailureSite
    Actual = Result.FailureSite
    Assert.IsTrue Actual = SetupSite, "Wrong failure site stored."
End Sub

Public Sub SetResult_WithDefaultSite_StoresDefault()
    Dim Result As TestResult
    Set Result = MakeTestResult
    
    Call Result.SetResult(ErrorState)
    
    Dim Actual As FailureSite
    Actual = Result.FailureSite
    Assert.IsTrue Actual = TestSite, "Wrong default site."
End Sub

Public Sub Failure_WithFailureSite_StoresFailureSite()
    Dim Result As TestResult
    Set Result = MakeTestResult
    
    Call Result.Failure("", FailureSite.ChildSite)
    
    Dim Actual As FailureSite
    Actual = Result.FailureSite
    Assert.IsTrue Actual = ChildSite, "Wrong failure site."
End Sub

Public Sub Error_WithFailureSite_StoresFailureSite()
    Dim Result As TestResult
    Set Result = MakeTestResult
    
    Call Result.Error("", FailureSite.ParentSite)
    
    Dim Actual As FailureSite
    Actual = Result.FailureSite
    Assert.IsTrue Actual = ParentSite, "Wrong failure site."
End Sub

Public Sub AddResult_WithErrorChild_SetsResultToFailure()
    Dim Parent As TestResult: Set Parent = MakeTestResult
    Dim Child As TestResult: Set Child = MakeTestResult
    Call Child.Error
    
    Call Parent.AddResult(Child)
    
    Dim Actual As ResultState
    Actual = Parent.ResultState
    Assert.IsTrue Actual = FailureState, "Wrong result state."
End Sub

Public Sub Error_WithTeardownSite_AppendsMessageInResult()
    Dim Result As TestResult
    Set Result = MakeTestResult
    Call Result.Failure("First message")
    
    Call Result.Error("second message", TeardownSite)
    
    Dim Actual As String
    Actual = Result.Message
    Const Expected As String = "First message" & vbCrLf & "Teardown: second message"
    Assert.IsTrue Actual = Expected, "Message should be appended."
End Sub

Public Sub SetResult_AsErrorInTeardownSite_AppensMessageInResult()
    Dim Result As TestResult
    Set Result = MakeTestResult
    Call Result.SetResult(FailureState, "First message")
    
    Call Result.SetResult(ErrorState, "second message", TeardownSite)
    
    Dim Actual As String
    Actual = Result.Message
    Const Expected As String = "First message" & vbCrLf & "Teardown: second message"
    Assert.IsTrue Actual = Expected, "Message should be appended."
End Sub

Public Sub AddResult_WithResultStateCombinations_SetsParentResultState(ByVal ParentState As ResultState, ByVal ChildState As ResultState, ByVal ExpectedState As ResultState)
    Dim Parent As TestResult
    Dim Child As TestResult
    Set Parent = MakeTestResult
    Set Child = MakeTestResult
    Call Parent.SetResult(ParentState)
    Call Child.SetResult(ChildState)
    
    Call Parent.AddResult(Child)
    
    Dim Actual As ResultState
    Actual = Parent.ResultState
    Assert.IsTrue Actual = ExpectedState, "Wrong result state."
End Sub

Public Sub Ignore_WhenCalled_SetsResultStateToIgnored()
    Dim Result As TestResult
    Set Result = MakeTestResult
    
    Call Result.Ignore
    
    Dim Actual As ResultState
    Actual = Result.ResultState
    Assert.IsTrue Actual = IgnoredState
End Sub

Public Sub Ignore_WhenCalled_IsIgnoredReturnsTrue()
    Dim Result As TestResult
    Set Result = MakeTestResult
    
    Call Result.Ignore
    
    Dim Actual As Boolean
    Actual = Result.IsIgnored
    Assert.IsTrue Actual = True
End Sub

Public Sub Ignore_WithReason_StoresReason()
    Dim Result As TestResult
    Set Result = MakeTestResult
    
    Call Result.Ignore("because")
    
    Dim Actual As String
    Actual = Result.Message
    Assert.IsTrue Actual = "because"
End Sub

Public Sub Invalid_WhenCalled_SetsResultStateToNotRunnable()
    Dim Result As TestResult
    Set Result = MakeTestResult
    
    Call Result.Invalid
    
    Dim Actual As ResultState
    Actual = Result.ResultState
    Assert.IsTrue Actual = NotRunnableState
End Sub

Public Sub Invalid_WithReason_StoresReason()
    Dim Result As TestResult
    Set Result = MakeTestResult
    
    Call Result.Invalid("because")
    
    Dim Actual As String
    Actual = Result.Message
    Assert.IsTrue Actual = "because"
End Sub



Private Function MakeTestResult() As TestResult
    Set MakeTestResult = Sim.NewTestResult(New FakeTest)
End Function

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SVBUnit.TestCaseDataBuilder)
    Select Case Test.MethodName
        Case "AddResult_WithResultStateCombinations_SetsParentResultState"
            Call Test.Use(InconclusiveState, InconclusiveState, InconclusiveState)
            Call Test.Use(InconclusiveState, NotRunnableState, InconclusiveState)
            Call Test.Use(InconclusiveState, IgnoredState, InconclusiveState)
            Call Test.Use(InconclusiveState, SuccessState, SuccessState)
            Call Test.Use(InconclusiveState, FailureState, FailureState)
            Call Test.Use(InconclusiveState, ErrorState, FailureState)
            Call Test.Use(InconclusiveState, CancelledState, CancelledState)
            Call Test.Use(NotRunnableState, InconclusiveState, NotRunnableState)
            Call Test.Use(NotRunnableState, NotRunnableState, NotRunnableState)
            Call Test.Use(NotRunnableState, IgnoredState, NotRunnableState)
            Call Test.Use(NotRunnableState, SuccessState, NotRunnableState)
            Call Test.Use(NotRunnableState, FailureState, FailureState)
            Call Test.Use(NotRunnableState, ErrorState, FailureState)
            Call Test.Use(NotRunnableState, CancelledState, CancelledState)
            Call Test.Use(IgnoredState, InconclusiveState, IgnoredState)
            Call Test.Use(IgnoredState, NotRunnableState, IgnoredState)
            Call Test.Use(IgnoredState, IgnoredState, IgnoredState)
            Call Test.Use(IgnoredState, SuccessState, IgnoredState)
            Call Test.Use(IgnoredState, FailureState, FailureState)
            Call Test.Use(IgnoredState, ErrorState, FailureState)
            Call Test.Use(IgnoredState, CancelledState, CancelledState)
            Call Test.Use(SuccessState, InconclusiveState, SuccessState)
            Call Test.Use(SuccessState, NotRunnableState, SuccessState)
            Call Test.Use(SuccessState, IgnoredState, SuccessState)
            Call Test.Use(SuccessState, SuccessState, SuccessState)
            Call Test.Use(SuccessState, FailureState, FailureState)
            Call Test.Use(SuccessState, ErrorState, FailureState)
            Call Test.Use(SuccessState, CancelledState, CancelledState)
            Call Test.Use(FailureState, InconclusiveState, FailureState)
            Call Test.Use(FailureState, NotRunnableState, FailureState)
            Call Test.Use(FailureState, IgnoredState, FailureState)
            Call Test.Use(FailureState, SuccessState, FailureState)
            Call Test.Use(FailureState, FailureState, FailureState)
            Call Test.Use(FailureState, ErrorState, FailureState)
            Call Test.Use(FailureState, CancelledState, CancelledState)
            Call Test.Use(ErrorState, InconclusiveState, ErrorState)
            Call Test.Use(ErrorState, NotRunnableState, ErrorState)
            Call Test.Use(ErrorState, IgnoredState, ErrorState)
            Call Test.Use(ErrorState, SuccessState, ErrorState)
            Call Test.Use(ErrorState, FailureState, ErrorState)
            Call Test.Use(ErrorState, ErrorState, ErrorState)
            Call Test.Use(ErrorState, CancelledState, CancelledState)
            Call Test.Use(CancelledState, InconclusiveState, CancelledState)
            Call Test.Use(CancelledState, NotRunnableState, CancelledState)
            Call Test.Use(CancelledState, IgnoredState, CancelledState)
            Call Test.Use(CancelledState, SuccessState, CancelledState)
            Call Test.Use(CancelledState, FailureState, FailureState)
            Call Test.Use(CancelledState, ErrorState, FailureState)
            Call Test.Use(CancelledState, CancelledState, CancelledState)
    End Select
End Sub
