VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArgWithReturnTestMethodTests"
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
' Module: ArgWithReturnTestMethodTests
'
Option Explicit

Public Sub Run_WhenCalled_ExecutesTestClassMethod()
    Dim Mock As New Mock1TestCaseSourceExpect
    Dim Member As MemberInfo
    Set Member = tli.InterfaceInfoFromObject(Mock).GetMember("Test1")
    Dim TestCaseData As New TestCaseData
    Set TestCaseData.Expected = PassingConstraint
    Dim Args() As Variant
    Args = Array(5)
    Call TestCaseData.CopyArguments(Args)
    Dim Method As ArgWithReturnTestMethod
    Set Method = Sim.NewArgWithReturnTestMethod(Mock, Member, TestCaseData)
    
    Call Method.Run
    
    Dim Actual As CallTrace
    Set Actual = Mock.Calls
    Dim Expected As New CallTrace
    Expected.Add "Test1.5"
    Assert.IsTrue Actual.Equals(Expected), "Method should be executed."
End Sub

Public Sub Run_WhenCalled_AppliesConstraintToResult()
    Dim Stub As New Mock1TestCaseSourceExpect
    Dim Member As MemberInfo
    Set Member = tli.InterfaceInfoFromObject(Stub).GetMember("Test1")
    Dim Mock As New MockConstraint
    Mock.Matches = True
    Dim TestCaseData As New TestCaseData
    Set TestCaseData.Expected = Mock
    Dim Args() As Variant
    Args = Array(5)
    Call TestCaseData.CopyArguments(Args)
    Dim Method As ArgWithReturnTestMethod
    Set Method = Sim.NewArgWithReturnTestMethod(Stub, Member, TestCaseData)
    
    Call Method.Run
    
    Dim Actual As CallTrace
    Set Actual = Mock.Calls
    Dim Expected As New CallTrace
    Call Expected.Add("Resolve")
    Call Expected.Add("Matches")
    Assert.IsTrue Actual.Equals(Expected)
End Sub

Public Sub Run_WhenCalled_AppliesConstraintToResult2()
    Dim Stub As New Mock1TestCaseSourceExpect
    Dim Member As MemberInfo
    Set Member = tli.InterfaceInfoFromObject(Stub).GetMember("Test1")
    Dim Mock As New MockConstraint
    Mock.Matches = True
    Dim TestCaseData As New TestCaseData
    Set TestCaseData.Expected = Mock
    Dim Args() As Variant
    Args = Array(5)
    Call TestCaseData.CopyArguments(Args)
    Dim Method As ArgWithReturnTestMethod
    Set Method = Sim.NewArgWithReturnTestMethod(Stub, Member, TestCaseData)
    
    Call Method.Run
    
    Dim Actual As Variant
    Actual = Mock.Actual
    Assert.That Actual, Iz.EqualTo(5)
End Sub

Public Sub Run_WhenCalled_AssertsResult()
    Dim Stub As New Mock1TestCaseSourceExpect
    Dim Member As MemberInfo
    Set Member = tli.InterfaceInfoFromObject(Stub).GetMember("Test1")
    Dim Mock As New MockConstraint
    Mock.Matches = False
    Dim TestCaseData As New TestCaseData
    Set TestCaseData.Expected = Mock
    Dim Args() As Variant
    Args = Array(5)
    Call TestCaseData.CopyArguments(Args)
    Dim Method As ArgWithReturnTestMethod
    Set Method = Sim.NewArgWithReturnTestMethod(Stub, Member, TestCaseData)
    
    On Error GoTo errTrap
    Call Method.Run
    On Error GoTo 0
    Assert.Fail "Should raise and error"
    
errTrap:
    Dim Actual As Long
    Actual = Err.Number
    Assert.IsTrue Actual = AssertCode.FailureCode
End Sub



Private Function PassingConstraint() As MockConstraint
    Set PassingConstraint = New MockConstraint
    PassingConstraint.Matches = True
End Function
