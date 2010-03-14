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
    Call Init
    
    Dim Suite As New TestSuite
    
    Call Suite.Add(New TestFixtureTests)
    Call Suite.Add(New TestResultTests)
    Call Suite.Add(New TestListTests)
    Call Suite.Add(New TestResultListTests)
    Call Suite.Add(New TestSuiteTests)
    Call Suite.Add(New TestCaseTests)
    Call Suite.Add(New AssertionsTests)
    Call Suite.Add(New NoArgTestMethodTests)
    Call Suite.Add(New ErrorInfoTests)
    Call Suite.Add(New MemberQueryTests)
    Call Suite.Add(New TestCaseDataBuilderTests)
    Call Suite.Add(New TestCaseDataTests)
    Call Suite.Add(New ArgOnlyTestMethodTests)
    Call Suite.Add(New TestCaseBuilderTests)
    Call Suite.Add(New TestListEnumeratorTests)
    Call Suite.Add(New StringBuilderTests)
    Call Suite.Add(New TextMessageWriterTests)
    Call Suite.Add(New EqualConstraintTests)
    Call Suite.Add(New IzTests)
    Call Suite.Add(New ArrayEnumeratorTests)
    Call Suite.Add(New EnumVariantEnumeratorTests)
    Call Suite.Add(New TestOutputTests)
    Call Suite.Add(New StackTests)
    Call Suite.Add(New TestContextTests)
    Call Suite.Add(New TestContextManagerTests)
    Call Suite.Add(New TestRunnerTests)
    Call Suite.Add(New EmptyFilterTests)
    Call Suite.Add(New FullNameFilterTests)
    Call Suite.Add(New CategoryListTests)
    Call Suite.Add(New CategoryParserTests)
    Call Suite.Add(New OrFilterTests)
    Call Suite.Add(New AndFilterTests)
    Call Suite.Add(New NotFilterTests)
    Call Suite.Add(New AndConstraintTests)
    Call Suite.Add(New OrConstraintTests)
    Call Suite.Add(New NotConstraintTests)
    Call Suite.Add(New ErrorHelperTests)
    Call Suite.Add(New ThrowsConstraintTests)
    Call Suite.Add(New TestCaseModifierTests)
    Call Suite.Add(New ArgWithReturnTestMethodTests)
    Call Suite.Add(New ConstraintBuilderTests)
    Call Suite.Add(New AndOperatorTests)
    Call Suite.Add(New OrOperatorTests)
    Call Suite.Add(New MsgUtilsTests)
    
    
    Dim Result As TestResult
    Set Result = Suite.Run
    Call PrintResults(Result)
    Call PrintSummary(Result)
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
