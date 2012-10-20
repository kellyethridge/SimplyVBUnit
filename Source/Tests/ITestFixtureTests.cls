VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ITestFixtureTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Copyright 2011 Kelly Ethridge
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
' Module: ITestFixtureTests
'
Option Explicit

Public Sub Count_WhenCalled_ShouldReturn5()
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(New MockITestFixture)
    
    Dim Actual As Long
    Actual = Fixture.TestCount
    
    Assert.AreEqual 5, Actual
End Sub

Public Sub Run_OnFirstTest_RunsMethods()
    Dim Mock As New MockITestFixture
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(Mock)
    Dim Filter As New FullNameFilter
    Filter.Add "MockITestFixture.Test4*"
    
    Fixture.Run , Filter
    
    AssertCalls Mock.Calls, "FixtureSetup", "Setup", "Test4.2", "Teardown", "Setup", "Test4.3", "Teardown"
End Sub
