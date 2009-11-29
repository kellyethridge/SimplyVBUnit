VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AssertionsTests"
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
' Module: AssertionsTests
'
Option Explicit

Public Sub Fail_WhenCalled_RaisesCorrectError()
    On Error GoTo errTrap
    
    Assert.Fail
    Debug.Assert False
    
errTrap:
    Debug.Assert Err.Number = AssertCode.FailureCode
End Sub

Public Sub Fail_WithMessage_IncludesMessageWithError()
    On Error GoTo errTrap
    
    Const ASSERT_MESSAGE As String = "message"
    Assert.Fail ASSERT_MESSAGE
    
errTrap:
    Debug.Assert Err.Description = ASSERT_MESSAGE
End Sub

Public Sub IsTrue_WithTrueCondition_DoesNothing()
    On Error GoTo errTrap
    
    Assert.IsTrue True
    Exit Sub
    
errTrap:
    Debug.Assert False
End Sub

Public Sub IsTrue_WithFalseCondition_RaisesCorrectError()
    On Error GoTo errTrap
    
    Assert.IsTrue False
    Debug.Assert False
    
errTrap:
    Debug.Assert Err.Number = AssertCode.FailureCode
End Sub

Public Sub IsTrue_WithMessageOnFalseCondition_IncludesMessageWithError()
    On Error GoTo errTrap
    
    Const ASSERT_MESSAGE As String = "message"
    Assert.IsTrue False, ASSERT_MESSAGE
    
errTrap:
    Const Expected As String = ASSERT_MESSAGE & vbCrLf & _
                               "Expected: True" & vbCrLf & _
                               "But was : False" & vbCrLf
    Debug.Assert Err.Description = Expected
End Sub

Public Sub That_WithConstraint_ExecutesConstraint()
    Dim Mock As New MockConstraint
    Mock.Matches = True
    
    Assert.That 1, Mock
    
    Dim Actual As CallTrace
    Set Actual = Mock.Calls
    Dim Expected As New CallTrace
    Expected.Add "Resolve"
    Expected.Add "Matches"
    Assert.IsTrue Actual.Equals(Expected)
End Sub

Public Sub That_WithFailingConstraint_RaisesCorrectError()
    Dim Stub As New StubConstraint
    Stub.Matches = False

    On Error GoTo errTrap
    Assert.That 1, Stub
    
    On Error GoTo 0
    Debug.Assert False

errTrap:
    Debug.Assert Err.Number = AssertCode.FailureCode
End Sub

Public Sub That_WithFailingConstraint_WritesConstraintMessage()
    Dim Mock As New MockConstraint
    Mock.Matches = False

    On Error GoTo errTrap
    Assert.That 1, Mock
    Debug.Assert False
    
errTrap:
    Dim Actual As CallTrace
    Set Actual = Mock.Calls
    Dim Expected As New CallTrace
    Expected.Add "Resolve"
    Expected.Add "Matches"
    Expected.Add "WriteMessageTo"
    Assert.IsTrue Actual.Equals(Expected)
End Sub

Public Sub That_WithFailingConstraint_IncludesConstraintMessageWithError()
    Dim Stub As New StubConstraint
    Stub.Matches = False
    Stub.Message = "failed"
    
    On Error GoTo errTrap
    Assert.That 1, Stub
    Debug.Assert False
    
errTrap:
    Dim Actual As ErrorInfo
    Set Actual = ErrorInfo.FromErr(Err)
    Dim Expected As ErrorInfo
    Set Expected = Sim.NewErrorInfo(AssertCode.FailureCode, "SVBUnit", "failed")
    Assert.IsTrue Actual.Equals(Expected)
End Sub

Public Sub That_WithMessage_PrependsConstraintMessage()
    Dim Stub As New StubConstraint
    Stub.Matches = False
    Stub.Message = "failed"
    
    On Error GoTo errTrap
    Assert.That 1, Stub, "the message"
    Debug.Assert False
    
errTrap:
    Dim Actual As ErrorInfo
    Set Actual = ErrorInfo.FromErr(Err)
    Dim Expected As ErrorInfo
    Set Expected = Sim.NewErrorInfo(AssertCode.FailureCode, "SVBUnit", "the message" & vbCrLf & "failed")
    Assert.IsTrue Actual.Equals(Expected)
End Sub

Public Sub IsFalse_WithFalseCondtion_DoesNothing()
    On Error GoTo errTrap
    
    Assert.IsFalse False
    Exit Sub
    
errTrap:
    Assert.Fail "Should not raise an error."
End Sub

Public Sub IsFalse_WithFailingCondition_RaisesCorrectError()
    On Error GoTo errTrap
    
    Assert.IsFalse True
    On Error GoTo 0
    Assert.Fail "Should raise error."
    
errTrap:
    Dim Actual As Long
    Actual = Err.Number
    Assert.IsTrue Actual = AssertCode.FailureCode, "Wrong error code."
End Sub

Public Sub IsFalse_WithMessage_IncludesMessage()
    On Error GoTo errTrap
    
    Const ASSERT_MESSAGE As String = "message"
    Assert.IsFalse True, ASSERT_MESSAGE
    
errTrap:
    Const Expected As String = ASSERT_MESSAGE & vbCrLf & _
                               "Expected: False" & vbCrLf & _
                               "But was : True" & vbCrLf
    Assert.IsTrue Err.Description = Expected, "Wrong message."
End Sub

Public Sub AreEqual_WithEqualValues_DoesNothing()
    On Error GoTo errTrap
    
    Assert.AreEqual 1, 1
    Exit Sub
    
errTrap:
    Assert.Fail "Should not raise error."
End Sub

Public Sub AreEqual_WithUnequalValues_RaisesCorrectError()
    On Error GoTo errTrap
    
    Assert.AreEqual 1, 2
    On Error GoTo 0
    Assert.Fail "An error should be raised."
    
errTrap:
    Dim Actual As Long
    Actual = Err.Number
    Assert.IsTrue Actual = AssertCode.FailureCode
End Sub

Public Sub AreEqual_WithMessage_IncludesMessage()
    On Error GoTo errTrap
    
    Const ASSERT_MESSAGE As String = "message"
    Assert.AreEqual 1, 2, ASSERT_MESSAGE
    
errTrap:
    Const Expected As String = ASSERT_MESSAGE & vbCrLf & _
                               "Expected: 1" & vbCrLf & _
                               "But was : 2" & vbCrLf
    Assert.IsTrue Err.Description = Expected, "Wrong message."
End Sub

Public Sub Pass_WhenCalled_RaisesCorrectError()
    On Error GoTo errTrap
    
    Assert.Pass
    On Error GoTo 0
    Assert.Fail "An error should be raised."
    
errTrap:
    Assert.IsTrue Err.Number = AssertCode.SuccessCode
End Sub

Public Sub Pass_WithMessage_IncludesMessageInError()
    On Error GoTo errTrap
    
    Assert.Pass "message"
    
errTrap:
    Assert.IsTrue Err.Description = "message"
End Sub

Public Sub Ignore_WhenCalled_RaisesCorrectError()
    On Error GoTo errTrap
    
    Assert.Ignore
    On Error GoTo 0
    Assert.Fail "An error should be raised."
    
errTrap:
    Assert.That Err.Number, Iz.EqualTo(AssertCode.IgnoreCode)
End Sub

Public Sub Ignore_WhenCalledWithMessage_IncludesMessageInError()
    On Error GoTo errTrap
    
    Assert.Ignore "message"
    
errTrap:
    Dim Actual As String
    Actual = Err.Description
    Assert.That Actual, Iz.EqualTo("message")
End Sub

Public Sub IsVBEmpty_WithNonemptyValue_RaisesCorrectError()
    On Error GoTo errTrap
    
    Assert.IsVBEmpty 1, "Not empty"
    On Error GoTo 0
    Assert.Fail "An error should be raised."
    
errTrap:
    Dim Actual As Long
    Actual = Err.Number
    Assert.That Actual, Iz.EqualTo(AssertCode.FailureCode)
End Sub

Public Sub IsVBEmpty_WithEmpty_DoesNothing()
    On Error GoTo errTrap
    
    Assert.IsVBEmpty Empty
    Exit Sub
    
errTrap:
    Assert.Fail "No error should be raised."
End Sub

Public Sub IsVBNull_WithNonnullValue_RaisesCorrectError()
    On Error GoTo errTrap
    
    Assert.IsVBNull 1, "Not null"
    On Error GoTo 0
    Assert.Fail "An error should be raised."
    
errTrap:
    Dim Actual As Long
    Actual = Err.Number
    Assert.That Actual, Iz.EqualTo(AssertCode.FailureCode)
End Sub







