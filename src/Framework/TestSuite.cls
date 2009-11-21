VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestSuite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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
' Module: TestSuite
'
Option Explicit
Implements ITest

Private Const DEF_NAME As String = "TestSuite"


Private mTests      As New TestList
Private mParent     As Long
Private mName       As String
Private mRunState   As RunState



Public Property Get RunState() As RunState
    RunState = mRunState
End Property

Public Property Let RunState(ByVal RHS As RunState)
    mRunState = RHS
End Property

Public Sub Sort(Optional ByVal Comparer As ITestComparer)
    Call mTests.Sort(Comparer)
    
    Dim Test As ITest
    For Each Test In mTests
        Call Test.Sort(Comparer)
    Next
End Sub

Public Property Get Parent() As ITest
    Set Parent = TestUtils.GetObjectFromPointer(mParent)
End Property

Public Property Set Parent(ByVal RHS As ITest)
    mParent = ObjPtr(RHS)
End Property

Public Property Get HasParent() As Boolean
    HasParent = (mParent <> vbNullPtr)
End Property

Public Sub Add(ByVal TestClass As Object)
    If TestClass Is Nothing Then _
        Call Err.Raise(ErrorCode.ArgumentNull, Resource.GetString(Src_TestSuite_Add), Resource.GetString(Arg_NullTestClass))
    
    Dim Test As ITest
    If TypeOf TestClass Is ITest Then
        Set Test = TestClass
    Else
        Set Test = Sim.NewTestFixture(TestClass)
    End If
    
    Set Test.Parent = Me
    Call mTests.Add(Test)
End Sub

Public Property Get TestCount() As Long
    Dim Test    As ITest
    Dim Result  As Long

    For Each Test In mTests
        Result = Result + Test.TestCount
    Next

    TestCount = Result
End Property

Public Property Get Tests() As TestList
    Set Tests = mTests
End Property

Public Function Run(Optional ByVal Listener As IEventListener) As TestResult
    If Listener Is Nothing Then
        Set Listener = NullListener.Instance
    End If
    
    Dim Result As TestResult
    Set Result = Sim.NewTestResult(Me)
    
    If Me.RunState = RunnableRunState Then
        Call Result.Success
        Call Listener.TestSuiteStarted(Me)
        
        Dim Ticks As Currency
        Ticks = Timing.GetTicks
        
        Dim Test As ITest
        For Each Test In mTests
            Call Result.AddResult(Test.Run)
        Next
        
        Result.Time = Timing.GetTime(Ticks)
        
        Call Listener.TestSuiteFinished(Result)
    Else
        Call Result.SetResult(NotRunnableState)
    End If
    
    Set Run = Result
End Function

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
    IsSuite = True
End Property

Public Property Get IsTestCase() As Boolean
    IsTestCase = False
End Property

Public Property Get TestClass() As Object
    Set TestClass = Nothing
End Property

Friend Sub Init(ByVal Name As String)
    mName = Name
End Sub

Private Sub Class_Initialize()
    Call Init(DEF_NAME)
End Sub

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

Private Property Get ITest_Name() As String
    ITest_Name = Name
End Property

Private Sub ITest_Reset()

End Sub

Private Function ITest_Run(Optional ByVal Listener As IEventListener) As TestResult
    Set ITest_Run = Run(Listener)
End Function

Private Property Let ITest_RunState(ByVal RHS As RunState)
    RunState = RHS
End Property

Private Property Get ITest_RunState() As RunState
    ITest_RunState = RunState
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
