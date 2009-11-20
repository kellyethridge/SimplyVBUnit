VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BootstrapTestSuiteTests"
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
' Module: BootstrapTestSuiteTests
'
Option Explicit
Implements IBootstrapTestClass

Public Sub Run()
    Add_WithTestClass_CreatesTestFixture
    Add_WithTestFixture_StoresTestFixture
    Add_WithTwoTestClasses_Creates2TestFixtures
    TestCount_With2TestsPerTestClass_Returns4
    Run_WithTestClass_ExecutesTest
    Add_WithTestSuite_StoresTestSuite
    Run_WithChildTestSuite_RunsChildTestSuite
    
End Sub

Private Sub Add_WithTestClass_CreatesTestFixture()
    Dim TestClass As New FakeEmpty
    Dim Suite As New TestSuite
    
    Suite.Add TestClass
    
    Dim Actual As Object
    Set Actual = Suite.Tests(0)
    Debug.Assert TypeOf Actual Is TestFixture
End Sub

Private Sub Add_WithTestFixture_StoresTestFixture()
    Dim TestClass As New FakeEmpty
    Dim Fixture As TestFixture
    Set Fixture = Sim.NewTestFixture(TestClass)
    Dim Suite As New TestSuite
    
    Suite.Add Fixture
    
    Dim Actual As Object
    Set Actual = Suite.Tests(0)
    Debug.Assert Actual Is Fixture
End Sub

Private Sub Add_WithTwoTestClasses_Creates2TestFixtures()
    Dim Suite As New TestSuite
    
    Suite.Add New Fake1Test
    Suite.Add New Fake1Test
    
    Dim Actual As Long
    Actual = Suite.Tests.Count
    Debug.Assert Actual = 2
End Sub

Private Sub TestCount_With2TestsPerTestClass_Returns4()
    Dim Suite As New TestSuite
    
    Suite.Add New Fake2Tests
    Suite.Add New Fake2Tests
    
    Dim Actual As Long
    Actual = Suite.TestCount
    Debug.Assert Actual = 4
End Sub

Private Sub Run_WithTestClass_ExecutesTest()
    Dim TestClass As New Mock1Test
    Dim Suite As New TestSuite
    Suite.Add TestClass
    
    Suite.Run
    
    Dim Actual As Long
    Actual = TestClass.Calls.Count
    Debug.Assert Actual = 1
End Sub

Private Sub Add_WithTestSuite_StoresTestSuite()
    Dim Suite1 As New TestSuite
    Dim Suite2 As New TestSuite
    
    Suite1.Add Suite2
    
    Dim Actual As Object
    Set Actual = Suite1.Tests(0)
    Debug.Assert Actual Is Suite2
End Sub

Private Sub Run_WithChildTestSuite_RunsChildTestSuite()
    Dim TestClass As New Mock1Test
    Dim ChildSuite As New TestSuite
    ChildSuite.Add TestClass
    Dim ParentSuite As New TestSuite
    ParentSuite.Add ChildSuite
    
    ParentSuite.Run
    
    Dim Actual As Long
    Actual = TestClass.Calls.Count
    Debug.Assert Actual = 1
End Sub


Private Sub IBootstrapTestClass_Run()
    Run
End Sub
