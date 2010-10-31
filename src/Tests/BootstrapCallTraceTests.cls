VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BootstrapCallTraceTests"
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
' Module: CallStackTests
'
Option Explicit
Implements IBootstrapTestClass

Private mCount As Long

Public Function Run() As Long
    mCount = 0
    
    RunTest Constructor_WhenCalled_IsEmpty
    RunTest Add_WhenCalled_StoresCallOnStack
    RunTest Count_WhenCalled_ReturnsCallCount
    RunTest Equals_WithSameCalls_ReturnsTrue
    RunTest Equals_WithSameCallsInDifferentOrder_ReturnsFalse
    RunTest Equals_WithDifferentSizeStacks_ReturnsFalse
    RunTest Equals_WithDifferentCase_ReturnsTrue
    RunTest Equals_WithPatternMatch_ReturnsTrue
    
    Run = mCount
End Function

Private Sub RunTest(ByRef v As Variant)
    mCount = mCount + 1
End Sub

Private Function Constructor_WhenCalled_IsEmpty()
    Dim Calls As New CallTrace
    
    Dim Actual As Long
    Actual = Calls.Count
    
    Debug.Assert Actual = 0
End Function

Public Function Add_WhenCalled_StoresCallOnStack()
    Dim Calls As New CallTrace
    
    Calls.Add "Method1"
    
    Dim Actual As String
    Actual = Calls(1)
    Debug.Assert Actual = "Method1"
End Function

Public Function Count_WhenCalled_ReturnsCallCount()
    Dim Calls As New CallTrace
    Calls.Add "Method"
    
    Dim Actual As Long
    Actual = Calls.Count
    
    Debug.Assert Actual = 1
End Function

Public Function Equals_WithSameCalls_ReturnsTrue()
    Dim Actual As New CallTrace
    Actual.Add "Test1"
    Actual.Add "Test2"
    
    Dim Expected As New CallTrace
    Expected.Add "Test1"
    Expected.Add "Test2"
    
    Debug.Assert Actual.Equals(Expected)
End Function

Public Function Equals_WithSameCallsInDifferentOrder_ReturnsFalse()
    Dim Actual As New CallTrace
    Actual.Add "Test1"
    Actual.Add "Test2"
    
    Dim Expected As New CallTrace
    Expected.Add "Test2"
    Expected.Add "Test1"
    
    Debug.Assert Actual.Equals(Expected) = False
End Function

Public Function Equals_WithDifferentSizeStacks_ReturnsFalse()
    Dim Actual As New CallTrace
    Actual.Add "Test1"
    Actual.Add "Test2"
    
    Dim Expected As New CallTrace
    Expected.Add "Test1"
    
    Debug.Assert Actual.Equals(Expected) = False
End Function

Public Function Equals_WithDifferentCase_ReturnsTrue()
    Dim Actual As New CallTrace
    Actual.Add "test1"
    
    Dim Expected As New CallTrace
    Expected.Add "TEST1"
    
    Debug.Assert Actual.Equals(Expected) = True
End Function

Public Function Equals_WithPatternMatch_ReturnsTrue()
    Dim Actual As New CallTrace
    Actual.Add "Test1"
    
    Dim Expected As New CallTrace
    Expected.Add "Test*"
    
    Debug.Assert Actual.Equals(Expected) = True
End Function




Private Function IBootstrapTestClass_Run() As Long
    IBootstrapTestClass_Run = Run
End Function
