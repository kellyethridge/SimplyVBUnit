VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ThrowsConstraintTests"
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
' Module: ThrowsConstraintTests
'
Option Explicit

Public Sub Matches_WithNoError_ReturnsFalse()
    Dim ExpectedInfo As ErrorInfo
    Set ExpectedInfo = Sim.NewErrorInfo(100, , "My Description")
    Dim ActualInfo As ErrorInfo
    Set ActualInfo = Sim.NewErrorInfo(0)
    
    Dim c As ThrowsConstraint
    Set c = Sim.NewThrowsConstraint(ExpectedInfo)
    
    Assert.That c.Matches(ActualInfo), Iz.False
End Sub

Public Sub Matches_WithEqualErrorNumber_ReturnsTrue()
    Dim ExpectedInfo As ErrorInfo
    Set ExpectedInfo = Sim.NewErrorInfo(100)
    Dim c As ThrowsConstraint
    Set c = Sim.NewThrowsConstraint(ExpectedInfo)
    Dim ActualInfo As ErrorInfo
    Set ActualInfo = Sim.NewErrorInfo(100)
    
    Dim Actual As Boolean
    Actual = c.Matches(ActualInfo)
    
    Assert.IsTrue Actual
End Sub

Public Sub Matches_WithEqualDescriptions_ReturnsTrue()
    Dim ExpectedInfo As ErrorInfo
    Set ExpectedInfo = Sim.NewErrorInfo(100, , "My Description")
    Dim c As ThrowsConstraint
    Set c = Sim.NewThrowsConstraint(ExpectedInfo)
    Dim ActualInfo As ErrorInfo
    Set ActualInfo = Sim.NewErrorInfo(100, , "My Description")

    Dim Actual As Boolean
    Actual = c.Matches(ActualInfo)
    
    Assert.That Actual, Iz.True
End Sub

Public Sub Matches_WithUnequalDescriptions_ReturnsFalse()
    Dim ExpectedInfo As ErrorInfo
    Set ExpectedInfo = Sim.NewErrorInfo(100, , "My Description")
    Dim c As ThrowsConstraint
    Set c = Sim.NewThrowsConstraint(ExpectedInfo)
    Dim ActualInfo As ErrorInfo
    Set ActualInfo = Sim.NewErrorInfo(100, , "Wrong Description")
    
    Dim Actual As Boolean
    Actual = c.Matches(ActualInfo)
    
    Assert.That Actual, Iz.False
End Sub

Public Sub Matches_WithDescriptionsOfWrongCase_ReturnsFalse()
    Dim ExpectedInfo As ErrorInfo
    Set ExpectedInfo = Sim.NewErrorInfo(100, , "My Description")
    Dim c As ThrowsConstraint
    Set c = Sim.NewThrowsConstraint(ExpectedInfo)
    Dim ActualInfo As ErrorInfo
    Set ActualInfo = Sim.NewErrorInfo(100, , "my description")
    
    Dim Actual As Boolean
    Actual = c.Matches(ActualInfo)
    
    Assert.That Actual, Iz.False
End Sub

Public Sub Matches_WithIgnoreCase_ReturnsTrue()
    Dim ExpectedInfo As ErrorInfo
    Set ExpectedInfo = Sim.NewErrorInfo(100, , "My Description")
    Dim c As ThrowsConstraint
    Set c = Sim.NewThrowsConstraint(ExpectedInfo).IgnoreCase
    Dim ActualInfo As ErrorInfo
    Set ActualInfo = Sim.NewErrorInfo(100, , "my description")
    
    Dim Actual As Boolean
    Actual = c.Matches(ActualInfo)

    Assert.That Actual, Iz.True
End Sub

Public Sub WriteMessageTo_WithMatchesNotCalled_WritesExpectedErrorMessage()
    Dim c As ThrowsConstraint
    Set c = Sim.NewThrowsConstraint(Sim.NewErrorInfo(1))
    Dim w As New TextMessageWriter

    Call c.WriteMessageTo(w)

    Dim Actual As String
    Actual = w.ToString
    Const Expected As String = "Expected: an error" & vbCrLf & "But was : no error raised" & vbCrLf
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub WriteMessageTo_WithWrongErrorRaised_WritesWrongErrorMessage()
    Dim c As ThrowsConstraint
    Set c = Sim.NewThrowsConstraint(Sim.NewErrorInfo(1))
    Dim Info As ErrorInfo
    Set Info = Sim.NewErrorInfo(100)
    Call c.Matches(Info)
    
    Dim w As New TextMessageWriter
    Call c.WriteMessageTo(w)
    
    Dim Actual As String
    Actual = w.ToString
    Const Expected = "Expected: error 1" & vbCrLf & "But was : error 100" & vbCrLf
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub














