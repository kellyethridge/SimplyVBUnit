VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestCase"
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
' Module: TestCase
'
Option Explicit
Implements ITest

Private mTestMethod     As ITestMethod
Private mName           As String
Private mParent         As Long
Private mOldAssertCount As Long
Private mTicks          As Currency
Private mRunState       As RunState
Private mContext        As ContextMethods


Public Property Get RunState() As RunState
    RunState = mRunState
End Property

Public Property Let RunState(ByVal RHS As RunState)
    mRunState = RHS
End Property

Public Property Get HasParent() As Boolean
    HasParent = (mParent <> vbNullPtr)
End Property

Public Property Get TestCount() As Long
    TestCount = 1
End Property

Public Property Get Tests() As TestList
    Set Tests = Nothing
End Property

Public Property Get TestClass() As Object
    Set TestClass = mTestMethod.TestClass
End Property

Public Property Get Name() As String
    Name = mName
End Property

Public Property Get FullName() As String
    FullName = TestUtils.BuildFullName(Me)
End Property

Public Property Get IsFixture() As Boolean
    IsFixture = False
End Property

Public Property Get IsSuite() As Boolean
    IsSuite = False
End Property

Public Property Get IsTestCase() As Boolean
    IsTestCase = True
End Property

Public Property Set Parent(ByVal RHS As ITest)
    mParent = ObjPtr(RHS)
End Property

Public Property Get Parent() As ITest
    Set Parent = TestUtils.GetObjectFromPointer(mParent)
End Property

Public Function Run(Optional ByVal Listener As IEventListener) As TestResult
    If Listener Is Nothing Then
        Set Listener = NullListener.Instance
    End If
    
    Dim Result As TestResult
    Set Result = Sim.NewTestResult(Me)
    
    Call Listener.TestCaseStarted(Me)
    
    If Me.RunState = RunnableRunState Then
        Call RunTestCase(Result, Listener)
    Else
        Call Result.Failure("FixtureSetup failed.", ParentSite)
    End If
    
    Call Listener.TestCaseFinished(Result)
    
    Set Run = Result
End Function

Private Sub RunTestCase(ByVal Result As TestResult, ByVal Listener As IEventListener)
    Call InitMetrics
    
    If RunSetup(Result) Then
        Call RunTestMethod(Result, Listener)
    End If
    
    Call RunTeardown(Result)
    Call Err.Clear
    Call RecordMetrics(Result)
End Sub

Private Sub InitMetrics()
    Call ResetAssertCount
    Call StartTiming
End Sub

Private Sub RecordMetrics(ByVal Result As TestResult)
    Result.Time = StopTiming
    Result.AssertCount = RestoreAssertCount
End Sub

Private Sub ResetAssertCount()
    mOldAssertCount = Assert.SetCount(0)
End Sub

Private Sub StartTiming()
    mTicks = Timing.GetTicks
End Sub

Private Function StopTiming() As Currency
    StopTiming = Timing.GetTime(mTicks)
End Function

Private Function RestoreAssertCount() As Long
    RestoreAssertCount = Assert.SetCount(mOldAssertCount)
End Function

Private Function RunSetup(ByVal Result As TestResult) As Boolean
    On Error GoTo errTrap
    Call mContext.RunSetup
    
    RunSetup = True
    Exit Function
    
errTrap:
    Call Result.Error("Error " & Err.Number & ", " & Err.Description, SetupSite)
End Function

Private Sub RunTestMethod(ByVal Result As TestResult, ByVal Listener As IEventListener)
    On Error Resume Next
    Call mTestMethod.Run

    Select Case Err.Number
        Case AssertCode.SuccessCode, ErrorCode.NoError
            Call Result.Success(Err.Description)
            mRunState = ExecutedRunState
            
        Case AssertCode.FailureCode
            Call Result.Failure(Err.Description, TestSite)
        Case Else
            Call Result.Error(Err.Description, TestSite)
            
            Dim Info As ErrorInfo
            Set Info = ErrorInfo.FromErr(Err)
            Call Listener.UnhandledError(Info)
    End Select
    
    Call Err.Clear
End Sub

Private Sub RunTeardown(ByVal Result As TestResult)
    On Error GoTo errTrap
    Call mContext.RunTeardown
    
    Exit Sub
    
errTrap:
    Call Result.Error("Error " & Err.Number & ", " & Err.Description, TeardownSite)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Name As String, ByVal TestMethod As ITestMethod)
    If Len(Name) = 0 Then _
        Call Err.Raise(ErrorCode.Argument, Resource.GetString(Src_TestCase_Init), Resource.GetString(Arg_EmptyName))
    If TestMethod Is Nothing Then _
        Call Err.Raise(ErrorCode.ArgumentNull, Resource.GetString(Src_TestCase_Init), Resource.GetString(Arg_NullTestMethod))
        
    mName = Name
    Set mTestMethod = TestMethod
End Sub

Friend Property Get Context() As ContextMethods
    Set Context = mContext
End Property

Friend Property Set Context(ByVal RHS As ContextMethods)
    If Not RHS Is Nothing Then
        Set mContext = RHS
    Else
        Set mContext = NullContext.Instance
    End If
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Me.RunState = RunnableRunState
    Set mContext = NullContext.Instance
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

Private Property Get ITest_Name() As String
    ITest_Name = Name
End Property

Private Sub ITest_Reset()

End Sub

Private Function ITest_Run(Optional ByVal Listener As IEventListener) As TestResult
    Set ITest_Run = Run(Listener)
End Function

Private Property Let ITest_RunState(ByVal RHS As RunState)
    Me.RunState = RHS
End Property

Private Property Get ITest_RunState() As RunState
    ITest_RunState = Me.RunState
End Property

Private Sub ITest_Sort(Optional ByVal Comparer As ITestComparer)
    ' do nothing
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
