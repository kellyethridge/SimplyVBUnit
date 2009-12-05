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
    Dim Data    As New TestCaseData
    Dim Member  As MemberInfo
    Set Member = MakeTest1Member(New Fake1Test)
    On Error GoTo errTrap
    
    Call Sim.NewArgOnlyTestMethod(Nothing, Member, Data)
    On Error GoTo 0
    
errTrap:
    Call AssertError(Err, ErrorCode.ArgumentNull, "ArgOnlyTestMethod.Init", "TestClass cannot be Nothing.")
End Sub

Public Sub Constructors_WithNothingMember_RaisesError()
    On Error GoTo errTrap
    
    Call Sim.NewArgOnlyTestMethod(New Fake1Test, Nothing, New TestCaseData)
    On Error GoTo 0
    
errTrap:
    Call AssertError(Err, ErrorCode.ArgumentNull, "ArgOnlyTestMethod.Init", "Member cannot be Nothing.")
End Sub

Public Sub Constructors_WithNothingTestCaseData_RaisesError()
    On Error GoTo errTrap
    
    Call Sim.NewArgOnlyTestMethod(New Fake1Test, MakeTest1Member(New Fake1Test), Nothing)
    On Error GoTo 0
    
errTrap:
    Call AssertError(Err, ErrorCode.ArgumentNull, "ArgOnlyTestMethod.Init", "TestCaseData cannot be Nothing.")
End Sub

Public Sub Run_WhenCalled_ExecutesTestClassMethod()
    Dim Mock As New Mock1TestCaseSource
    Dim Member As MemberInfo
    Set Member = MakeTest1Member(Mock)
    Dim Data As TestCaseData
    Set Data = MakeTestData("one", "two", 3)
    
    Dim Method As ArgOnlyTestMethod
    Set Method = Sim.NewArgOnlyTestMethod(Mock, Member, Data)
    
    Call Method.Run
    
    Call AssertCalls(Mock.Calls, "Test1.one.two.3")
End Sub


Private Function MakeTest1Member(ByVal TestClass As Object) As MemberInfo
    Set MakeTest1Member = tli.InterfaceInfoFromObject(TestClass).GetMember(Constants.TEST1_METHOD)
End Function

Private Function MakeTestData(ParamArray Args() As Variant) As TestCaseData
    Dim Data As New TestCaseData
    Dim LocalArgs() As Variant
    LocalArgs = Args
    Call Data.CopyArguments(LocalArgs)
    Set MakeTestData = Data
End Function
