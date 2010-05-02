Attribute VB_Name = "modMain"
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
' Module: modMain
'
Option Explicit

#If False Then
    Dim Iz
#End If

Private mBootstraps As Long
Private mPassed     As Long
Private mFailed     As Long

Private Sub RunTestClassTests()
    Init
    
    Dim Suite As New TestSuite
    
    Suite.Add New TestFixtureTests
    Suite.Add New TestResultTests
    Suite.Add New TestListTests
    Suite.Add New TestResultListTests
    Suite.Add New TestSuiteTests
    Suite.Add New TestCaseTests
    Suite.Add New AssertionsTests
    Suite.Add New NoArgTestMethodTests
    Suite.Add New ErrorInfoTests
    Suite.Add New MemberQueryTests
    Suite.Add New TestCaseDataBuilderTests
    Suite.Add New TestCaseDataTests
    Suite.Add New ArgOnlyTestMethodTests
    Suite.Add New TestCaseBuilderTests
    Suite.Add New TestListEnumeratorTests
    Suite.Add New StringBuilderTests
    Suite.Add New TextMessageWriterTests
    Suite.Add New EqualConstraintTests
    Suite.Add New IzTests
    Suite.Add New ArrayEnumeratorTests
    Suite.Add New EnumVariantEnumeratorTests
    Suite.Add New TestOutputTests
    Suite.Add New StackTests
    Suite.Add New TestContextTests
    Suite.Add New TestContextManagerTests
    Suite.Add New TestRunnerTests
    Suite.Add New EmptyFilterTests
    Suite.Add New FullNameFilterTests
    Suite.Add New CategoryListTests
    Suite.Add New CategoryParserTests
    Suite.Add New OrFilterTests
    Suite.Add New AndFilterTests
    Suite.Add New NotFilterTests
    Suite.Add New AndConstraintTests
    Suite.Add New OrConstraintTests
    Suite.Add New NotConstraintTests
    Suite.Add New ErrorHelperTests
    Suite.Add New ThrowsConstraintTests
    Suite.Add New TestCaseModifierTests
    Suite.Add New ArgWithReturnTestMethodTests
    Suite.Add New ConstraintBuilderTests
    Suite.Add New AndOperatorTests
    Suite.Add New OrOperatorTests
    Suite.Add New MsgUtilsTests
    Suite.Add New ConstraintExpressionTests
    Suite.Add New ToleranceTests
    
    
    Dim Result As TestResult
    Set Result = Suite.Run
    PrintResults Result
    PrintSummary Result
End Sub



Private Sub Init()
    Debug.Print String$(50, "-")
    Debug.Print "Running tests"
    mPassed = 0
    mFailed = 0
End Sub

Private Sub PrintSummary(ByVal Result As TestResult)
    Debug.Print String$(50, "-")
    Debug.Print "Total : " & Result.Test.TestCount + mBootstraps
    Debug.Print "Passed: " & mPassed + mBootstraps
    Debug.Print "Failed: " & mFailed
    Debug.Print "Time  : " & Result.Time & "ms"
End Sub

Private Sub RunBootstrapTests()
    mBootstraps = 0
    Call RunBootstrapTestClass(New BootstrapUtilitiesTests)
    Call RunBootstrapTestClass(New BootstrapCallTraceTests)
    Call RunBootstrapTestClass(New BootstrapCallErrorTests)
    Call RunBootstrapTestClass(New BootstrapMockOneSubTestClassTests)
    Call RunBootstrapTestClass(New BootstrapStubOneSubTestClassTests)
    Call RunBootstrapTestClass(New BootstrapMockTestsWithSetupTests)
    Call RunBootstrapTestClass(New BootstrapTestFixtureTests)
    Call RunBootstrapTestClass(New BootstrapTestSuiteTests)
End Sub

Private Sub RunBootstrapTestClass(ByVal TestClass As IBootstrapTestClass)
    mBootstraps = mBootstraps + TestClass.Run
End Sub

Private Sub Main()
    RunBootstrapTests
    RunTestClassTests
End Sub

Private Sub PrintResults(ByVal Result As TestResult, Optional ByVal Indent As Long)
    If Result.IsFailure Or Result.IsError Then
        Debug.Print Space$(Indent); Result.Test.Name & ": " & Result.Message
        If Not Result.Test.IsSuite Then
            mFailed = mFailed + 1
        End If
    ElseIf Not Result.Test.IsSuite Then
        mPassed = mPassed + 1
    End If
    
    Dim Child As TestResult
    For Each Child In Result.Results
        Call PrintResults(Child, Indent + 4)
    Next
End Sub
