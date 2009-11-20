VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestFixture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
' Module: TestFixture
'
Option Explicit
Implements ITest


Private mTestClass              As Object
Private mParent                 As Long
Private mTests                  As New TestList
Private mFixtureSetup           As ITestMethod
Private mFixtureTeardown        As ITestMethod
Private mOldAssertCount         As Long
Private mTicks                  As Currency
Private mRunState               As RunState
Private mTestCaseContext        As ContextMethods


Public Property Get RunState() As RunState
    RunState = mRunState
End Property

Public Property Let RunState(ByVal RHS As RunState)
    mRunState = RHS
End Property

Public Sub Sort(Optional ByVal Comparer As ITestComparer)
    Call mTests.Sort(Comparer)
End Sub

Public Property Get Parent() As ITest
    Set Parent = TestUtils.GetObjectFromPointer(mParent)
End Property

Public Property Set Parent(ByVal RHS As ITest)
    mParent = ObjPtr(RHS)
End Property

Public Property Get TestCount() As Long
    Dim Test    As ITest
    Dim Result  As Long

    For Each Test In mTests
        Result = Result + Test.TestCount
    Next

    TestCount = Result
End Property

Public Function Run(Optional ByVal Listener As IEventListener) As TestResult
    If Listener Is Nothing Then
        Set Listener = NullListener.Instance
    End If
    
    Dim Result As TestResult
    Set Result = Sim.NewTestResult(Me)
    
    Call Result.Success
    Call Listener.TestFixtureStarted(Me)
    Call InitMetrics
    
    If RunFixtureSetup(Result) = True Then
        Call RunTests(Result)
    End If
    
    Call RunFixtureTeardown
    Call RecordMetrics(Result)
    Call Listener.TestFixtureFinished(Result)
    
    Set Run = Result
End Function

Private Sub RunTests(ByVal Result As TestResult)
    Dim Test As ITest
    For Each Test In mTests
        Result.AddResult Test.Run
    Next
End Sub

Private Sub InitMetrics()
    Call ResetAssertCount
    Call StartTiming
End Sub

Private Sub ResetAssertCount()
    mOldAssertCount = Assert.SetCount(0)
End Sub

Private Sub StartTiming()
    mTicks = Timing.GetTicks
End Sub

Private Sub RecordMetrics(ByVal Result As TestResult)
    Result.Time = StopTiming
    Result.AssertCount = RestoreAssertCount
End Sub

Private Function StopTiming() As Currency
    StopTiming = Timing.GetTime(mTicks)
End Function

Private Function RestoreAssertCount() As Long
    RestoreAssertCount = Assert.SetCount(mOldAssertCount)
End Function

Public Property Get Name() As String
    Name = TypeName(mTestClass)
End Property

Public Property Get FullName() As String
    FullName = TestUtils.BuildFullName(Me)
End Property

Public Property Get Tests() As TestList
    Set Tests = mTests
End Property

Public Property Get IsFixture() As Boolean
    IsFixture = True
End Property

Public Property Get IsSuite() As Boolean
    IsSuite = False
End Property

Public Property Get IsTestCase() As Boolean
    IsTestCase = False
End Property

Public Property Get HasParent() As Boolean
    HasParent = (mParent <> vbNullPtr)
End Property

Public Property Get TestClass() As Object
    Set TestClass = mTestClass
End Property

Friend Sub Init(ByVal TestClass As Object)
    If TestClass Is Nothing Then _
        Call Err.Raise(ErrorCode.ArgumentNull, Resource.GetString(Src_TestFixture_Init), Resource.GetString(Arg_NullTestClass))
        
    Set mTestClass = TestClass
    Call BuildFixture
End Sub

Private Sub BuildFixture()
    Dim Result As QueryResult
    Set Result = QueryTestClass
    
    Call AssignContextMethods(Result)
    Call BuildTestCasesFromMembers(Result.TestCaseMembers)
    Call SetContext(mTests)
End Sub

Private Function QueryTestClass() As QueryResult
    Dim Query As New MemberQuery
    Set QueryTestClass = Query.Execute(mTestClass)
End Function

Private Sub AssignContextMethods(ByVal Result As QueryResult)
    Call AssignTestCaseContext(Result.SetupMember, Result.TeardownMember)
    Call AssignTestFixtureContext(Result.FixtureSetupMember, Result.FixtureTeardownMember)
End Sub

Private Sub AssignTestCaseContext(ByVal SetupMember As MemberInfo, ByVal TeardownMember As MemberInfo)
    Dim SetupMethod     As ITestMethod
    Dim TeardownMethod  As ITestMethod
    
    Set SetupMethod = BuildContextMethod(SetupMember)
    Set TeardownMethod = BuildContextMethod(TeardownMember)
    Set mTestCaseContext = Sim.NewContextMethods(SetupMethod, TeardownMethod)
End Sub

Private Sub AssignTestFixtureContext(ByVal FixtureSetupMember As MemberInfo, ByVal FixtureTeardownMember As MemberInfo)
    Set mFixtureSetup = BuildContextMethod(FixtureSetupMember)
    Set mFixtureTeardown = BuildContextMethod(FixtureTeardownMember)
End Sub

Private Function BuildContextMethod(ByVal Member As MemberInfo) As ITestMethod
    If ContextMemberWasFound(Member) Then
        Set BuildContextMethod = Sim.NewNoArgTestMethod(mTestClass, Member)
    Else
        Set BuildContextMethod = NullMethod.Instance
    End If
End Function

Private Function ContextMemberWasFound(ByVal Member As MemberInfo) As Boolean
    ContextMemberWasFound = (Not Member Is Nothing)
End Function

Private Sub BuildTestCasesFromMembers(ByVal Members As Collection)
    Dim Member As MemberInfo
    
    For Each Member In Members
        Call BuildTestCases(Member)
    Next
End Sub

Private Sub BuildTestCases(ByVal Member As MemberInfo)
    Dim Test As ITest
    
    For Each Test In TestCaseBuilder.BuildFrom(mTestClass, Member)
        Call AddTestCase(Test)
    Next
End Sub

Private Sub AddTestCase(ByVal Test As ITest)
    Set Test.Parent = Me
    Call mTests.Add(Test)
End Sub

Private Sub SetContext(ByVal Tests As TestList)
    Dim Test As ITest
    For Each Test In Tests
        If Test.IsTestCase Then
            Call AssignContextToTestCase(Test)
        ElseIf Test.IsSuite Then
            Call SetContext(Test.Tests)
        End If
    Next
End Sub

Private Sub AssignContextToTestCase(ByVal Test As TestCase)
    Set Test.Context = mTestCaseContext
End Sub

Private Function RunFixtureSetup(ByVal FixtureResult As TestResult) As Boolean
    On Error GoTo errTrap
    Call mFixtureSetup.Run
    
    RunFixtureSetup = True
    Exit Function
    
errTrap:
    Call FixtureResult.Error("FixtureSetup failed.", SetupSite)
    Call FailAllTests(mTests, FixtureResult)
End Function

Private Sub FailAllTests(ByVal Tests As TestList, ByVal ParentResult As TestResult)
    Dim Test As ITest
    For Each Test In Tests
        Dim Result As TestResult
        Set Result = Sim.NewTestResult(Test)
        Call Result.Failure("FixtureSetup failed.", ParentSite)
        Call ParentResult.AddResult(Result)
        
        If Not Test.Tests Is Nothing Then
            Call FailAllTests(Test.Tests, Result)
        End If
    Next
End Sub

Private Sub RunFixtureTeardown()
    Call mFixtureTeardown.Run
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
    Call TestUtils.ClearParents(mTests)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ITest Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get ITest_FullName() As String
    ITest_FullName = FullName
End Property

Private Property Get ITest_HasParent() As Boolean
    ITest_HasParent = HasParent
End Property

Private Property Get ITest_IsFixture() As Boolean
    ITest_IsFixture = IsFixture
End Property

Private Property Get ITest_IsSuite() As Boolean
    ITest_IsSuite = IsSuite
End Property

Private Property Get ITest_IsTestCase() As Boolean
    ITest_IsTestCase = IsTestCase
End Property

Private Property Get ITest_Parent() As ITest
    Set ITest_Parent = Parent
End Property

Private Property Set ITest_Parent(ByVal RHS As ITest)
    Set Parent = RHS
End Property

Private Sub ITest_Reset()

End Sub

Private Function ITest_Run(Optional ByVal Listener As IEventListener) As TestResult
    Set ITest_Run = Run(Listener)
End Function

Private Property Get ITest_Name() As String
    ITest_Name = Name
End Property

Private Property Let ITest_RunState(ByVal RHS As RunState)
    Me.RunState = RHS
End Property

Private Property Get ITest_RunState() As RunState
    ITest_RunState = Me.RunState
End Property

Private Sub ITest_Sort(Optional ByVal Comparer As ITestComparer)
    Call Sort(Comparer)
End Sub

Private Property Get ITest_TestClass() As Object
    Set ITest_TestClass = TestClass
End Property

Private Property Get ITest_TestCount() As Long
    ITest_TestCount = TestCount
End Property

Private Property Get ITest_Tests() As TestList
    Set ITest_Tests = Tests
End Property