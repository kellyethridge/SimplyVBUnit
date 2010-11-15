VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NotConstraintTests"
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
' Module: NotConstraintTests
'
Option Explicit

Public Sub Matches_WithPassingBaseConstraint_ReturnsFalse()
    Dim Stub As New StubConstraint
    Stub.Matches = True
    Dim c As NotConstraint
    Set c = Sim.NewNotConstraint(Stub)
    
    Dim Actual As Boolean
    Actual = c.Matches(True)
    
    Assert.IsFalse Actual
End Sub

Public Sub Matches_WithFailingBaseConstraint_ReturnsTrue()
    Dim Stub As New StubConstraint
    Stub.Matches = False
    Dim c As NotConstraint
    Set c = Sim.NewNotConstraint(Stub)
    
    Dim Actual As Boolean
    Actual = c.Matches(True)
    
    Assert.IsTrue Actual
End Sub

Public Sub Constructor_WithNullBaseConstraint_RaisesError()
    Assert.Throws ErrorCode.ArgumentNull
    
    Sim.NewNotConstraint Nothing
End Sub

Public Sub WriteMessageTo_WhenCalled_WritesBaseConstraintMessage()
    Dim Stub As New StubConstraint
    Stub.Description = "equals"
    Stub.ActualValue = 1
    Dim c As IConstraint
    Set c = Sim.NewNotConstraint(Stub)
    c.Matches 1
    Dim w As New TextMessageWriter
    
    c.WriteMessageTo w
    
    Dim Actual As String
    Actual = w.ToString
    Const Expected As String = "  Expected: not equals" & vbCrLf & _
                               "  But was : 1" & vbCrLf
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub
