VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestCaseBuilderTests"
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
' Module: TestCaseBuilderTests
'
Option Explicit

Public Sub BuildFrom_With1Test_Returns1TestCase()
    Dim Fake As New Fake1Test
    Dim Member As MemberInfo
    Set Member = tli.InterfaceInfoFromObject(Fake).GetMember(Constants.TEST1_METHOD)
    Dim Builder As New TestCaseBuilder
    
    Dim Tests As Collection
    Set Tests = Builder.BuildFrom(Fake, Member)
    
    Dim Actual As Long
    Actual = Tests.Count
    Assert.IsTrue Actual = 1, "Should be 1 test case."
End Sub

Public Sub BuildFrom_With1TestCaseSource_CallsSourceForTestCase()
    Dim Mock As New Mock1TestCaseSource
    Dim Member As MemberInfo
    Set Member = tli.InterfaceInfoFromObject(Mock).GetMember(Constants.TEST1_METHOD)
    Dim Builder As New TestCaseBuilder
    
    Call Builder.BuildFrom(Mock, Member)
    
    Call AssertCalls(Mock.Calls, "GetTestCases.Test1")
End Sub

Public Sub BuildFrom_With2TestCaseSource_Returns1TestSuite()
    Dim Mock As New Mock2TestCaseSource
    Dim Member As MemberInfo
    Set Member = tli.InterfaceInfoFromObject(Mock).GetMember(Constants.TEST1_METHOD)
    Dim Builder As New TestCaseBuilder
    
    Dim Tests As Collection
    Set Tests = Builder.BuildFrom(Mock, Member)
    
    Dim Actual As ITest
    Set Actual = Tests(1)
    Assert.IsTrue Actual.IsSuite, "Should be test suite."
End Sub

Public Sub BuildFrom_With2TestCaseSource_Returns2TestCasesWithinTestSuite()
    Dim Mock As New Mock2TestCaseSource
    Dim Member As MemberInfo
    Set Member = tli.InterfaceInfoFromObject(Mock).GetMember(Constants.TEST1_METHOD)
    Dim Builder As New TestCaseBuilder
    
    Dim Tests As Collection
    Set Tests = Builder.BuildFrom(Mock, Member)
    
    Dim Parent As ITest
    Set Parent = Tests(1)
    Dim Actual As Long
    Actual = Parent.TestCount
    Assert.IsTrue Actual = 2, "Wrong test count."
End Sub

Public Sub BuildFrom_With2TestCaseSource_TestCaseNamesEqualParentNamesPlusSequenceNumber()
    Dim Mock As New Mock2TestCaseSource
    Dim Member As MemberInfo
    Set Member = tli.InterfaceInfoFromObject(Mock).GetMember(Constants.TEST1_METHOD)
    Dim Builder As New TestCaseBuilder
    
    Dim Tests As Collection
    Set Tests = Builder.BuildFrom(Mock, Member)
    
    Dim Parent As ITest
    Set Parent = Tests(1)
    Dim Test As ITest
    Set Test = Parent.Tests(0)
    Assert.IsTrue Test.Name = "Test11"
    
    Set Test = Parent.Tests(1)
    Assert.IsTrue Test.Name = "Test12"
End Sub

Public Sub BuildFrom_WithTestWithArgsDoesNotImplementITestCaseSource_ReturnsEmpty()
    Dim Fake As New Fake1TestWithArgsNoSource
    Dim Member As MemberInfo
    Set Member = tli.InterfaceInfoFromObject(Fake).GetMember(Constants.TEST1_METHOD)
    Dim Builder As New TestCaseBuilder
    
    Dim Actual As Collection
    Set Actual = Builder.BuildFrom(Fake, Member)
    
    Assert.IsTrue Actual.Count = 0, "Should be 0 test cases."
End Sub

Public Sub BuildFrom_WithTestNameSet_ReturnsWithTestName()
    Dim Stub As New Stub1TestCaseSourceTestName
    Stub.TestName = "Custom"
    Dim Member As MemberInfo
    Set Member = tli.InterfaceInfoFromObject(Stub).GetMember(Constants.TEST1_METHOD)
    Dim Builder As New TestCaseBuilder
    
    Dim Tests As Collection
    Set Tests = Builder.BuildFrom(Stub, Member)
    
    Dim Parent As ITest
    Set Parent = Tests(1)

    Dim Test As ITest
    Set Test = Parent.Tests(0)
    Dim Actual As String
    Actual = Test.Name
    Assert.That Actual, Iz.EqualTo("Custom")
End Sub























