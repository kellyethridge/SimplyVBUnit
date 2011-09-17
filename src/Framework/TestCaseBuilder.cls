VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestCaseBuilder"
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
' Module: TestCaseBuilder
'
Option Explicit

Private mTestCases      As Collection
Private mMember         As MemberInfo
Private mTestClass      As Object
Private mTestSequence   As Long
Private mTestName       As String


Public Function BuildFrom(ByVal TestClass As Object, ByVal Member As MemberInfo) As Collection
    Init TestClass, Member
    BuildTestCases
    Set BuildFrom = Release
End Function

Private Sub Init(ByVal TestClass As Object, ByVal Member As MemberInfo)
    Set mTestCases = New Collection
    Set mMember = Member
    Set mTestClass = TestClass
    mTestSequence = 0
    mTestName = Member.Name
End Sub

Private Sub BuildTestCases()
    If HasNoParameters Then
        BuildSimpleTestMethod
    ElseIf IsParameterized Then
        BuildParameterizedTestMethod
    End If
End Sub

Private Function Release() As Collection
    Set mMember = Nothing
    Set mTestClass = Nothing
    Set Release = mTestCases
    Set mTestCases = Nothing
End Function

Private Function IsParameterized() As Boolean
    IsParameterized = HasParameters And CanProvideArguments
End Function

Private Function HasParameters() As Boolean
    HasParameters = (mMember.Parameters.Count > 0)
End Function

Private Function CanProvideArguments() As Boolean
    CanProvideArguments = (TypeOf mTestClass Is ITestCaseSource)
End Function

Private Function HasNoParameters() As Boolean
    HasNoParameters = Not HasParameters
End Function

Private Sub BuildSimpleTestMethod()
    Dim Method As ITestMethod
    Set Method = Sim.NewSimpleTestMethod(mTestClass, mMember.MemberId)
    AddTestCase Method
End Sub

Private Function CanBuildIndirectly() As Boolean
    CanBuildIndirectly = (TypeOf mTestClass Is ITestCaseSource)
End Function

Private Sub BuildParameterizedTestMethod()
    Dim Suite As TestSuite
    Set Suite = Sim.NewTestSuite(mMember.Name)
    mTestCases.Add Suite
    
    Dim Data As TestCaseData
    For Each Data In GetTestCases
        Dim TestCase As TestCase
        Set TestCase = MakeTestCaseFromData(Data)
        Suite.Add TestCase
    Next
End Sub

Private Function MakeTestCaseFromData(ByVal Data As TestCaseData) As TestCase
    Dim Method As ITestMethod
    Set Method = Sim.NewParameterizedTestMethod(mTestClass, mMember.MemberId, Data)
    
    Set MakeTestCaseFromData = Sim.NewTestCase(NextSequentialTestName, Method)
End Function

Private Function NextSequentialTestName() As String
    mTestSequence = mTestSequence + 1
    NextSequentialTestName = mTestName & "_" & mTestSequence
End Function

Private Function GetTestCases() As Collection
    Dim Builder As TestCaseCollector
    Set Builder = Sim.NewTestCaseCollector(mMember.Name)
    
    Dim Source As ITestCaseSource
    Set Source = mTestClass
    Source.GetTestCases Builder
    
    Set GetTestCases = Builder.DataStore
End Function

Private Sub AddTestCase(ByVal Method As ITestMethod)
    Dim TestCase As TestCase
    Set TestCase = Sim.NewTestCase(mTestName, Method)
    mTestCases.Add TestCase
End Sub










