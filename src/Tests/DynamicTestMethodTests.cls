VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DynamicTestMethodTests"
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
' Module: DynamicTestMethodTests
'
Option Explicit

Public Sub Run_WhenCalled_ExecutesTestClassMethod()
    Dim Mock As New Mock1TestCaseSourceExpect
    Dim Method As DynamicTestMethod
    Set Method = MakeDynamicTestMethod(Mock)

    Method.Run

    AssertCalls Mock.Calls, "Test1.5"
End Sub

Public Sub Run_WhenCalled_AppliesConstraintToResult()
    Dim Stub As New Mock1TestCaseSourceExpect
    Dim Member As MemberInfo
    Set Member = tli.InterfaceInfoFromObject(Stub).VTableInterface.GetMember(Constants.TEST1_METHOD)
    Dim Mock As New MockConstraint
    Mock.Matches = True
    Dim TestCaseData As New TestCaseData
    Set TestCaseData.Expected = Mock
    Dim Args() As Variant
    Args = Array(5)
    TestCaseData.CopyArguments Args
    Dim Method As DynamicTestMethod
    Set Method = Sim.NewDynamicTestMethod(Stub, Member, TestCaseData)

    Method.Run

    AssertCalls Mock.Calls, "Resolve", "Matches"
End Sub

Public Sub Run_WhenCalled_AppliesConstraintToResult2()
    Dim Stub As New Mock1TestCaseSourceExpect
    Dim Member As MemberInfo
    Set Member = tli.InterfaceInfoFromObject(Stub).GetMember(Constants.TEST1_METHOD)
    Dim Mock As New MockConstraint
    Mock.Matches = True
    Dim TestCaseData As New TestCaseData
    Set TestCaseData.Expected = Mock
    Dim Args() As Variant
    Args = Array(5)
    TestCaseData.CopyArguments Args
    Dim Method As DynamicTestMethod
    Set Method = Sim.NewDynamicTestMethod(Stub, Member, TestCaseData)

    Method.Run

    Dim Actual As Variant
    Actual = Mock.Actual
    Assert.That Actual, Iz.EqualTo(5)
End Sub

Public Sub Run_WhenCalled_AssertsResult()
    Dim Stub As New Mock1TestCaseSourceExpect
    Dim Member As MemberInfo
    Set Member = tli.InterfaceInfoFromObject(Stub).GetMember(Constants.TEST1_METHOD)
    Dim Mock As New MockConstraint
    Mock.Matches = False
    Dim TestCaseData As New TestCaseData
    Set TestCaseData.Expected = Mock
    Dim Args() As Variant
    Args = Array(5)
    TestCaseData.CopyArguments Args
    Dim Method As DynamicTestMethod
    Set Method = Sim.NewDynamicTestMethod(Stub, Member, TestCaseData)

    On Error GoTo errTrap
    Method.Run
    On Error GoTo 0

errTrap:
    AssertError Err, AssertCode.FailureCode
End Sub


Private Function MakeDynamicTestMethod(ByVal TestClass As Object) As DynamicTestMethod
    Dim Member As MemberInfo
    Set Member = tli.InterfaceInfoFromObject(TestClass).VTableInterface.GetMember(Constants.TEST1_METHOD)
    Dim TestCaseData As New TestCaseData
    Set TestCaseData.Expected = PassingConstraint
    Dim Args() As Variant
    Args = Array(5)
    TestCaseData.CopyArguments Args
    
    Set MakeDynamicTestMethod = Sim.NewDynamicTestMethod(TestClass, Member, TestCaseData)
End Function

Private Function PassingConstraint() As MockConstraint
    Set PassingConstraint = New MockConstraint
    PassingConstraint.Matches = True
End Function
