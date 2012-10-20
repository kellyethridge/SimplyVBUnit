VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EmptyFilterTests"
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
' Module: EmptyFilterTests
'
Option Explicit
Implements ITestCaseSource


Private Sub ITestCaseSource_GetTestCases(ByVal Test As SVBUnit.TestCaseCollector)
    Select Case Test.MethodName
        Case "Match_WithTestNames_ReturnsTrue"
            Call Test.Use("NAME")
            Call Test.Use("name")
            Call Test.Use("")
            
        Case "Pass_WithTestNames_ReturnsTrue"
            Call Test.Use("NAME")
            Call Test.Use("name")
            Call Test.Use("")
    End Select
End Sub

Public Sub Match_WithTestNames_ReturnsTrue(ByVal FullName As String)
    Dim Stub As New StubTest
    Stub.FullName = FullName
    
    Dim Actual As Boolean
    Actual = TestFilter.EmptyFilter.Match(Stub)
    
    Assert.IsTrue Actual = True
End Sub

Public Sub Pass_WithTestNames_ReturnsTrue(ByVal FullName As String)
    Dim Stub As New StubTest
    Stub.FullName = FullName
    
    Dim Actual As Boolean
    Actual = TestFilter.EmptyFilter.Pass(Stub)
    
    Assert.IsTrue Actual = True
End Sub
