VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestCaseDataBuilderTests"
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
' Module: TestCaseDataBuilderTests
'
Option Explicit

Public Sub Constructor_WithMethodName_StoresMethodName()
    Const METHOD_NAME As String = "Method1"
    Dim Builder As TestCaseDataBuilder
    Set Builder = Sim.NewTestCaseDataBuilder(METHOD_NAME, New Collection)
    
    Dim Actual As String
    Actual = Builder.MethodName
    
    Assert.IsTrue Actual = METHOD_NAME, "Wrong method name."
End Sub

Public Sub Using_With1Arg_Adds1TestCaseDataToStore()
    Dim Store As New Collection
    Dim Builder As TestCaseDataBuilder
    Set Builder = Sim.NewTestCaseDataBuilder("Test1", Store)
    
    Builder.Using 1
    
    Dim Actual As Long
    Actual = Store.Count
    Assert.IsTrue Actual = 1, "Wrong TestCaseData count."
End Sub

Public Sub Using_With1Arg_Adds1TestCaseDataToStore2()
    Dim Store As New Collection
    Dim Builder As TestCaseDataBuilder
    Set Builder = Sim.NewTestCaseDataBuilder("Test1", Store)
    
    Builder.Using 1
    
    Dim TestCaseData As TestCaseData
    Set TestCaseData = Store(1)
    Dim Actual() As Variant
    Actual = TestCaseData.Arguments
    Assert.IsTrue Actual(0) = 1, "Wrong argument."
End Sub

' Arguments need to be in reverse order to be properly used by the TLIApplication.Invoke methods.
Public Sub Using_With3Args_StoresArgsInReverse()
    Dim Store As New Collection
    Dim Builder As TestCaseDataBuilder
    Set Builder = Sim.NewTestCaseDataBuilder("Test1", Store)
    
    Builder.Using "one", "two", 3
    
    Dim TestCaseData As TestCaseData
    Set TestCaseData = Store(1)
    Dim Actual() As Variant
    Actual = TestCaseData.Arguments
    Assert.IsTrue Actual(0) = 3, "Expected 3"
    Assert.IsTrue Actual(1) = "two", "Expected 'two'"
    Assert.IsTrue Actual(2) = "one", "Expected 'one'"
End Sub

Public Sub Using_With2TestCaseDatas_StoresTestCaseDatas()
    Dim Store As New Collection
    Dim Builder As TestCaseDataBuilder
    Set Builder = Sim.NewTestCaseDataBuilder("Test1", Store)
    
    Builder.Using 1, 2
    Builder.Using "one", "two"
    
    Dim Actual As Long
    Actual = Store.Count
    Assert.IsTrue Actual = 2, "Wrong TestCaseData count."
End Sub

Public Sub Using_WithNoArgs_DoesNotCreateTestCaseData()
    Dim Store As New Collection
    Dim Builder As TestCaseDataBuilder
    Set Builder = Sim.NewTestCaseDataBuilder("Test1", Store)
    
    Builder.Using
    
    Dim Actual As Long
    Actual = Store.Count
    Assert.IsTrue Actual = 0, "Should be 0 TestCaseDatas."
End Sub
