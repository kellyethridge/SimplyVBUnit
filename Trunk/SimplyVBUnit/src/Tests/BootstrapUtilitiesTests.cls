VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BootstrapUtilitiesTests"
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
' Module: BootstrapUtilitiesTests
'
Option Explicit
Implements IBootstrapTestClass

Private mCount As Long

Public Function Run() As Long
    mCount = 0
    
    RunTest Quote_WithText_ReturnsTextInQuotes
    RunTest AssertError_WithNoError_RaisesError
    RunTest AssertError_WithWrongErrorNumber_RaisesError
    RunTest AssertError_WithWrongErrorSource_RaisesError
    RunTest AssertError_WithWrongDescription_RaisesError
    RunTest AssertError_WithCorrectError_Passes
    RunTest AssertCalls_WithWrongCall_RaisesError
    RunTest AssertCalls_WithWrongNumberOfCalls_RaisesError
    RunTest AssertCalls_WithCorrectCalls_Passes
    RunTest AssertNoCalls_WithCalls_RaisesError
    RunTest AssertNoCalls_WithNoCallsMade_Passes
    
    Run = mCount
End Function

Private Sub RunTest(ByRef v As Variant)
    mCount = mCount + 1
End Sub

Private Function Quote_WithText_ReturnsTextInQuotes()
    Const TEST_TEXT As String = "text"
    
    Dim Actual As String
    Actual = Quote(TEST_TEXT)
    
    Const Expected As String = """text"""
    Debug.Assert Actual = Expected
End Function

Private Function AssertError_WithNoError_RaisesError()
    On Error GoTo errTrap
    
    AssertError Err, 1
    Debug.Assert False ' should raise error
    
errTrap:
    Debug.Assert Err.Number = AssertCode.FailureCode
End Function

Private Function AssertError_WithWrongErrorNumber_RaisesError()
    On Error GoTo errTrap
    
    Err.Number = 1
    Err.Source = "a"
    Err.Description = "b"
    
    AssertError Err, 2, "a", "b"
    Debug.Assert False ' should raise error
    
errTrap:
    Debug.Assert Err.Number = AssertCode.FailureCode
End Function

Private Function AssertError_WithWrongErrorSource_RaisesError()
    On Error GoTo errTrap
    
    Err.Number = 1
    Err.Source = "a"
    Err.Description = "b"
    
    AssertError Err, 1, "b", "b"
    Debug.Assert False ' should raise error
    
errTrap:
    Debug.Assert Err.Number = AssertCode.FailureCode
End Function

Private Function AssertError_WithWrongDescription_RaisesError()
    On Error GoTo errTrap
    
    Err.Number = 1
    Err.Source = "a"
    Err.Description = "b"
    
    AssertError Err, 1, "a", "c"
    Debug.Assert False ' should raise error
    
errTrap:
    Debug.Assert Err.Number = AssertCode.FailureCode
End Function

Private Function AssertError_WithCorrectError_Passes()
    On Error GoTo errTrap
    
    Err.Number = 1
    Err.Source = "a"
    Err.Description = "b"
    
    AssertError Err, 1, "a", "b"
    Exit Function
    
errTrap:
    Debug.Assert False ' should not raise error
End Function

Private Function AssertCalls_WithWrongCall_RaisesError()
    Dim Actual As New CallTrace
    Actual.Add "Method1"
    
    On Error GoTo errTrap
    
    AssertCalls Actual, "Method2"
    On Error GoTo 0
    Debug.Assert False ' an error should be raised
    
errTrap:
    Debug.Assert Err.Number = AssertCode.FailureCode
End Function

Private Function AssertCalls_WithWrongNumberOfCalls_RaisesError()
    Dim Actual As New CallTrace
    Actual.Add "Method1"
    
    On Error GoTo errTrap
    
    AssertCalls Actual, "Method1", "Method2"
    On Error GoTo 0
    Debug.Assert False ' an error should be raised
    
errTrap:
    Debug.Assert Err.Number = AssertCode.FailureCode
End Function

Private Function AssertCalls_WithCorrectCalls_Passes()
    Dim Actual As New CallTrace
    Actual.Add "Method1"
    
    On Error GoTo errTrap
    AssertCalls Actual, "Method1"
    Exit Function
    
errTrap:
    Debug.Assert False ' no error should be raised
End Function

Private Function AssertNoCalls_WithCalls_RaisesError()
    Dim Actual As New CallTrace
    Actual.Add "Method1"
    
    On Error GoTo errTrap
    AssertNoCalls Actual
    On Error GoTo 0
    Debug.Assert False ' an error should be raised
    
errTrap:
    Debug.Assert Err.Number = AssertCode.FailureCode
End Function

Private Function AssertNoCalls_WithNoCallsMade_Passes()
    Dim Actual As New CallTrace
    
    On Error GoTo errTrap
    AssertNoCalls Actual
    Exit Function
    
errTrap:
    Debug.Assert False ' no error should be raised
End Function



Private Function IBootstrapTestClass_Run() As Long
    IBootstrapTestClass_Run = Run
End Function
