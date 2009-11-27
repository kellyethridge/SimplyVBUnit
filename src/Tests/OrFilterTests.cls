VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrFilterTests"
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
' Module: OrFilterTests
'
Option Explicit

Public Sub Match_WithMatchingFilter_ReturnsTrue()
    Dim Test As New StubTest
    Test.FullName = "Test1"
    Dim Subfilter As New FullNameFilter
    Call Subfilter.Add("Test1")
    Dim Filter As New OrFilter
    Call Filter.Add(Subfilter)
    
    Dim Actual As Boolean
    Actual = Filter.Match(Test)
    
    Assert.IsTrue Actual
End Sub

Public Sub Match_WithOneOfTwoMatchingFilters_ReturnsTrue()
    Dim Test As New StubTest
    Test.FullName = "Test1"
    Dim Subfilter1 As New FullNameFilter
    Call Subfilter1.Add("WrongName")
    Dim Subfilter2 As New FullNameFilter
    Call Subfilter2.Add("Test1")
    Dim Filter As New OrFilter
    Call Filter.Add(Subfilter1)
    Call Filter.Add(Subfilter2)
    
    Dim Actual As Boolean
    Actual = Filter.Match(Test)
    
    Assert.IsTrue Actual
End Sub

Public Sub Match_WithNoMatchingFilter_ReturnsFalse()
    Dim Test As New FakeTest
    Dim Stub As New StubFilter
    Stub.Matches = False
    Dim Filter As New OrFilter
    Call Filter.Add(Stub)
    
    Dim Actual As Boolean
    Actual = Filter.Match(Test)
    
    Assert.IsFalse Actual
End Sub

Public Sub Pass_WithPassingFilter_ReturnsTrue()
    Dim Test As New FakeTest
    Dim Stub As New StubFilter
    Stub.Matches = True
    Dim Filter As New OrFilter
    Call Filter.Add(Stub)
    
    Dim Actual As Boolean
    Actual = Filter.Pass(Test)
    
    Assert.IsTrue Actual
End Sub

Public Sub Pass_WithParentOfMatchingTest_ReturnsTrue()
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(New Fake1Test)
    Dim Subfilter As New FullNameFilter
    Call Subfilter.Add("Fake1Test.Test1")
    Dim Filter As New OrFilter
    Call Filter.Add(Subfilter)
    
    Dim Actual As Boolean
    Actual = Filter.Pass(Fixture)
    
    Assert.IsTrue Actual
End Sub

Public Sub Pass_WithChildOfMatchingFixture_ReturnsTrue()
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(New Fake1Test)
    Dim Test As ITest
    Set Test = Fixture.Tests(0)
    Dim Subfilter As New FullNameFilter
    Call Subfilter.Add("Fake1Test")
    Dim Filter As New OrFilter
    Call Filter.Add(Subfilter)
    
    Dim Actual As Boolean
    Actual = Filter.Pass(Test)
    
    Assert.IsTrue Actual
End Sub























