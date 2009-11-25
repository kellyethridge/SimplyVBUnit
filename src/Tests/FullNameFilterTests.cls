VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FullNameFilterTests"
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
' Module: FullNameFilterTests
'
Option Explicit

Public Sub Match_WithEqualFullNames_ReturnsTrue()
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(New Fake2Tests)
    Dim Test As ITest
    Set Test = Fixture.Tests(0)
    
    Dim Filter As New FullNameFilter
    Call Filter.Add("Fake2Tests.Test1")
    
    Dim Actual As Boolean
    Actual = Filter.Match(Test)
    
    Assert.IsTrue Actual = True
End Sub

Public Sub Match_WithUnequalFullNames_ReturnsFalse()
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(New Fake2Tests)
    Dim Test As ITest
    Set Test = Fixture.Tests(0)
    
    Dim Filter As New FullNameFilter
    Call Filter.Add("UnequalName")
    
    Dim Actual As Boolean
    Actual = Filter.Match(Test)
    
    Assert.IsFalse Actual = True
End Sub

Public Sub Match_WithMultipleEqualNames_ReturnsTrue()
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(New Fake2Tests)
    Dim Test1 As ITest
    Dim Test2 As ITest
    Set Test1 = Fixture.Tests(0)
    Set Test2 = Fixture.Tests(1)
    
    Dim Filter As New FullNameFilter
    Call Filter.Add("Fake2Tests.Test1")
    Call Filter.Add("Fake2Tests.Test2")
    
    Dim Actual1 As Boolean
    Dim Actual2 As Boolean
    Actual1 = Filter.Match(Test1)
    Actual2 = Filter.Match(Test2)
    
    Assert.IsTrue Actual1 = True
    Assert.IsTrue Actual2 = True
End Sub

Public Sub Match_WithEqualNamesDifferenceCasing_ReturnsTrue()
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(New Fake2Tests)
    Dim Test As ITest
    Set Test = Fixture.Tests(0)
    
    Dim Filter As New FullNameFilter
    Call Filter.Add("Fake2Tests.Test1")
    
    Dim Actual As Boolean
    Actual = Filter.Match(Test)
    
    Assert.IsTrue Actual = True
End Sub

Public Sub Pass_WithExactMatch_ReturnsTrue()
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(New Fake2Tests)
    Dim Test As ITest
    Set Test = Fixture.Tests(0)
    
    Dim Filter As New FullNameFilter
    Call Filter.Add("Fake2Tests.Test1")
    
    Dim Actual As Boolean
    Actual = Filter.Pass(Test)
    
    Assert.That Actual, Iz.True
End Sub

Public Sub Pass_WithMatchingDecendent_ReturnsTrue()
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(New Fake2Tests)
    
    Dim Filter As New FullNameFilter
    Call Filter.Add("Fake2Tests.Test2")
    
    Dim Actual As Boolean
    Actual = Filter.Pass(Fixture)
    
    Assert.That Actual, Iz.True
End Sub

Public Sub Pass_WithNoMatchingDescendents_ReturnsFalse()
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(New Fake2Tests)
    
    Dim Filter As New FullNameFilter
    Call Filter.Add("Fake2Tests.None")
    
    Dim Actual As Boolean
    Actual = Filter.Pass(Fixture)
    
    Assert.That Actual, Iz.False
End Sub

Public Sub Pass_WithMatchingParent_ReturnsTrue()
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(New Fake2Tests)
    Dim Suite As TestSuite
    Set Suite = Sim.NewTestSuite("ParentSuite")
    Call Suite.Add(Fixture)
    
    Dim Filter As New FullNameFilter
    Call Filter.Add("ParentSuite")
    
    Dim Actual As Boolean
    Actual = Filter.Pass(Fixture)
    
    Assert.That Actual, Iz.True
End Sub

Public Sub Pass_WithNoMatchingParent_ReturnsFalse()
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(New Fake2Tests)
    Dim Suite As TestSuite
    Set Suite = Sim.NewTestSuite("ParentSuite")
    Call Suite.Add(Fixture)
    
    Dim Filter As New FullNameFilter
    Call Filter.Add("NotParentSuite")
    
    Dim Actual As Boolean
    Actual = Filter.Pass(Fixture)
    
    Assert.That Actual, Iz.False
End Sub

Public Sub Add_WithExistingNames_DoesNothing()
    On Error GoTo errTrap
    
    Dim Filter As New FullNameFilter
    Call Filter.Add("Test")
    Call Filter.Add("Test")
    Exit Sub
    
errTrap:
    Assert.Fail "Should not raise an error."
End Sub

































