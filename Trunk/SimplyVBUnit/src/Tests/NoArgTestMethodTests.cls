VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NoArgTestMethodTests"
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
' Module: NoArgTestMethodTests
'
Option Explicit

Public Sub Constructor_WithNothingTestClass_RaisesError()
    On Error GoTo errTrap
    
    Dim Method As NoArgTestMethod
    Dim Member As MemberInfo
    Set Member = tli.InterfaceInfoFromObject(New Fake1Test).GetMember("Test1")
    
    Set Method = Sim.NewNoArgTestMethod(Nothing, Member)
    On Error GoTo 0
    Assert.Fail "An error should be raised."
    
errTrap:
    Assert.IsTrue Err.Number = ErrorCode.ArgumentNull
End Sub

Public Sub Constructor_WithNothingMember_RaisesError()
    On Error GoTo errTrap
    
    Dim Method As NoArgTestMethod
    
    Set Method = Sim.NewNoArgTestMethod(New Fake1Test, Nothing)
    On Error GoTo 0
    Assert.Fail "An error should be raised."
    
errTrap:
    Assert.IsTrue Err.Number = ErrorCode.ArgumentNull
End Sub

Public Sub Run_WhenCalled_ExecutesTestClassMethod()
    Dim TestClass As New Mock1Test
    Dim Member As MemberInfo
    Set Member = tli.InterfaceInfoFromObject(TestClass).GetMember("Test1")
    
    Dim Method As NoArgTestMethod
    Set Method = Sim.NewNoArgTestMethod(TestClass, Member)
    
    Method.Run
    
    Dim Actual As CallTrace
    Set Actual = TestClass.Calls
    Dim Expected As New CallTrace
    Expected.Add "Test1"
    Assert.IsTrue Actual.Equals(Expected), "Method should be executed."
End Sub

Public Sub Constructor_WhenCalled_StoresTestClass()
    Dim Fake As New Fake1Test
    Dim Member As MemberInfo
    Set Member = tli.InterfaceInfoFromObject(Fake).GetMember("Test1")
    Dim Method As NoArgTestMethod
    
    Set Method = Sim.NewNoArgTestMethod(Fake, Member)
    
    Dim Actual As Object
    Set Actual = Method.TestClass
    Assert.IsTrue Actual Is Fake, "Wrong test class returned."
End Sub
