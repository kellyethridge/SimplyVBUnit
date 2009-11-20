VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArgOnlyTestMethodTests"
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
' Module: ArgsOnlyTestMethodTests
'
Option Explicit

Public Sub Constructor_WithNothingTestClass_RaisesError()
    On Error GoTo errTrap
    
    Dim Method As ArgOnlyTestMethod
    Dim Member As MemberInfo
    Set Member = tli.InterfaceInfoFromObject(New Fake1Test).GetMember("Test1")
    
    Set Method = Sim.NewArgOnlyTestMethod(Nothing, Member, New TestCaseData)
    On Error GoTo 0
    Assert.Fail "An error should be raised."
    
errTrap:
    Dim Actual As ErrorInfo
    Set Actual = ErrorInfo.FromErr(Err)
    Dim Expected As ErrorInfo
    Set Expected = Sim.NewErrorInfo(ErrorCode.ArgumentNull, "ArgOnlyTestMethod.Init", "TestClass cannot be Nothing.")
    Assert.IsTrue Actual.Equals(Expected), "Wrong error."
End Sub

Public Sub Constructors_WithNothingMember_RaisesError()
    On Error GoTo errTrap
    
    Dim Method As ArgOnlyTestMethod
    
    Set Method = Sim.NewArgOnlyTestMethod(New Fake1Test, Nothing, New TestCaseData)
    On Error GoTo 0
    Assert.Fail "An error should be raised."
    
errTrap:
    Dim Actual As ErrorInfo
    Set Actual = ErrorInfo.FromErr(Err)
    Dim Expected As ErrorInfo
    Set Expected = Sim.NewErrorInfo(ErrorCode.ArgumentNull, "ArgOnlyTestMethod.Init", "Member cannot be Nothing.")
    Assert.IsTrue Actual.Equals(Expected), "Wrong error."
End Sub

Public Sub Constructors_WithNothingTestCaseData_RaisesError()
    On Error GoTo errTrap
    
    Dim Method As ArgOnlyTestMethod
    Dim Member As MemberInfo
    Set Member = tli.InterfaceInfoFromObject(New Fake1Test).GetMember("Test1")

    Set Method = Sim.NewArgOnlyTestMethod(New Fake1Test, Member, Nothing)
    On Error GoTo 0
    Assert.Fail "An error should be raised."
    
errTrap:
    Assert.IsTrue Err.Number = ErrorCode.ArgumentNull
End Sub

Public Sub Run_WhenCalled_ExecutesTestClassMethod()
    Dim Mock As New Mock1TestCaseSource
    Dim Member As MemberInfo
    Set Member = tli.InterfaceInfoFromObject(Mock).GetMember("Test1")
    Dim TestCaseData As New TestCaseData
    Dim Args() As Variant
    Args = Array("one", "two", 3)
    TestCaseData.CopyArguments Args
    
    Dim Method As ArgOnlyTestMethod
    Set Method = Sim.NewArgOnlyTestMethod(Mock, Member, TestCaseData)
    
    Method.Run
    
    Dim Actual As CallTrace
    Set Actual = Mock.Calls
    Dim Expected As New CallTrace
    Expected.Add "Test1.one.two.3"
    Assert.IsTrue Actual.Equals(Expected), "Method should be executed."
End Sub

