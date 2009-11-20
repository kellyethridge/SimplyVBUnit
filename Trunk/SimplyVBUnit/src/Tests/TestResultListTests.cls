VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestResultListTests"
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
' Module: TestResultListTests
'
Option Explicit

Public Sub Constructor_WhenCalled_CreatesEmptyList()
    On Error GoTo errTrap
    Dim List As New TestResultList

    Dim Actual As Long
    Actual = List.Count

    Debug.Assert Actual = 0
    Exit Sub
    
errTrap:
    Debug.Assert False
End Sub

Public Sub Add_WithTestResult_AddsResultToList()
    On Error GoTo errTrap
    Dim List As New TestResultList
    Dim Result As TestResult
    Set Result = MakeResult

    List.Add Result

    Dim Actual As Long
    Actual = List.Count
    Debug.Assert Actual = 1
    Exit Sub
    
errTrap:
    Debug.Assert False
End Sub

Public Sub Add_WithTestResult_AddsResultToList2()
    On Error GoTo errTrap
    Dim List As New TestResultList
    Dim Result As TestResult
    Set Result = MakeResult

    List.Add Result

    Dim Actual As TestResult
    Set Actual = List(0)
    Debug.Assert Actual Is Result
    Exit Sub
    
errTrap:
    Debug.Assert False
End Sub

Public Sub Add_WithNothing_ThrowsException()
    On Error GoTo errTrap

    Dim List As New TestResultList

    List.Add Nothing
    Debug.Assert False
    
errTrap:
    Debug.Assert Err.Number = ErrorCode.ArgumentNull
End Sub

Public Sub Item_WithIndexTooLow_ThrowsException()
    On Error GoTo errTrap

    Dim List As TestResultList
    Set List = MakeListWithOneResult

    Const TOO_LOW As Long = -1
    Dim Result As TestResult
    Set Result = List(TOO_LOW)
    Debug.Assert False
    
errTrap:
    Debug.Assert Err.Number = ErrorCode.IndexOutOfRange
End Sub

Public Sub Item_WithIndexTooHigh_ThrowsException()
    On Error GoTo errTrap

    Dim List As TestResultList
    Set List = MakeListWithOneResult

    Dim TOO_HIGH As Long
    TOO_HIGH = 1
    Dim Result As TestResult
    Set Result = List(TOO_HIGH)
    Debug.Assert False
    
errTrap:
    Debug.Assert Err.Number = ErrorCode.IndexOutOfRange
End Sub

Public Sub NewEnum_WhenCalled_ReturnsEnumerator()
    On Error GoTo errTrap
    Dim List As New TestResultList

    Dim Enumerator As IUnknown
    Set Enumerator = List.NewEnum
    
    Debug.Assert Not Enumerator Is Nothing
    Exit Sub
    
errTrap:
    Debug.Assert False
End Sub









Private Function MakeListWithOneResult() As TestResultList
    Set MakeListWithOneResult = New TestResultList
    Call MakeListWithOneResult.Add(MakeResult)
End Function

Private Function MakeResult() As TestResult
    Set MakeResult = Sim.NewTestResult(New FakeTest)
End Function

