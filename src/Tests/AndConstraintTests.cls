VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AndConstraintTests"
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
' Module: AndConstraintTests
'
Option Explicit

Public Sub Matches_WithTwoTrueConstraints_ReturnsTrue()
    Dim StubLHS As New StubConstraint
    StubLHS.Matches = True
    Dim StubRHS As New StubConstraint
    StubRHS.Matches = True
    Dim c As AndConstraint
    Set c = Sim.NewAndConstraint(StubLHS, StubRHS)
    
    Dim Actual As Boolean
    Actual = c.Matches(True)
    
    Assert.IsTrue Actual
End Sub

Public Sub Matches_WithFailingLHSConstraint_ReturnsFalse()
    Dim StubLHS As New StubConstraint
    StubLHS.Matches = False
    Dim StubRHS As New StubConstraint
    StubRHS.Matches = True
    Dim c As AndConstraint
    Set c = Sim.NewAndConstraint(StubLHS, StubRHS)
    
    Dim Actual As Boolean
    Actual = c.Matches(True)
    
    Assert.IsFalse Actual
End Sub

Public Sub Matches_WithFailingLHSConstraint_DoesNotCallRHSConstraint()
    Dim StubLHS As New StubConstraint
    StubLHS.Matches = False
    Dim MockRHS As New MockConstraint
    MockRHS.Matches = True
    Dim c As AndConstraint
    Set c = Sim.NewAndConstraint(StubLHS, MockRHS)
    
    Call c.Matches(True)

    Dim Actual As CallTrace
    Set Actual = MockRHS.Calls
    Dim Expected As New CallTrace
    Assert.IsTrue Actual.Equals(Expected)
End Sub

Public Sub Constructor_WithLeftSideAsNothing_RaisesError()
    On Error GoTo errTrap
    Dim c As AndConstraint
    
    Set c = Sim.NewAndConstraint(Nothing, New StubConstraint)
    On Error GoTo 0
    Assert.Fail "An error should be raised."
    
errTrap:
    Dim Actual As ErrorInfo
    Set Actual = ErrorInfo.FromErr(Err)
    Dim Expected As ErrorInfo
    Set Expected = Sim.NewErrorInfo(ErrorCode.ArgumentNull, "AndConstraint.Init", "LeftSide cannot be Nothing.")
    Assert.IsTrue Actual.Equals(Expected)
End Sub

Public Sub Constructor_WithRightSideAsNothing_RaisesError()
    On Error GoTo errTrap
    Dim c As AndConstraint
    
    Set c = Sim.NewAndConstraint(New StubConstraint, Nothing)
    On Error GoTo 0
    Assert.Fail "An error should be raised."
    
errTrap:
    Dim Actual As ErrorInfo
    Set Actual = ErrorInfo.FromErr(Err)
    Dim Expected As ErrorInfo
    Set Expected = Sim.NewErrorInfo(ErrorCode.ArgumentNull, "AndConstraint.Init", "RightSide cannot be Nothing.")
    Assert.IsTrue Actual.Equals(Expected)
End Sub

Public Sub WriteMessageTo_WhenCalled_WritesLeftAndRightMessages()
    Dim LeftConstraint  As New StubConstraint
    Dim RightConstraint As New StubConstraint
    LeftConstraint.Description = "equal left"
    RightConstraint.Description = "equal right"
    Dim w As New TextMessageWriter
    Dim c As AndConstraint
    Set c = Sim.NewAndConstraint(LeftConstraint, RightConstraint)
    Call c.Matches(1)
    
    Call c.WriteMessageTo(w)
    
    Dim Actual As String
    Actual = w.ToString
    Const Expected As String = "Expected: equal left and equal right" & vbCrLf & _
                               "But was : 1" & vbCrLf
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub










