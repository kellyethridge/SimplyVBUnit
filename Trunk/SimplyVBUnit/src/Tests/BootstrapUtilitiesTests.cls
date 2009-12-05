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
    
    Call RunTest(Quote_WithText_ReturnsTextInQuotes)
    Call RunTest(AssertError_WithNoError_RaisesError)
    Call RunTest(AssertError_WithWrongErrorNumber_RaisesError)
    Call RunTest(AssertError_WithWrongErrorSource_RaisesError)
    Call RunTest(AssertError_WithWrongDescription_RaisesError)
    Call RunTest(AssertError_WithCorrectError_Passes)
    
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
    Call Err.Clear
    
    Call AssertError(Err, 1)
    Debug.Assert False ' should raise error
    
errTrap:
    Debug.Assert Err.Number = AssertCode.FailureCode
End Function

Private Function AssertError_WithWrongErrorNumber_RaisesError()
    On Error GoTo errTrap
    
    Err.Number = 1
    Err.Source = "a"
    Err.Description = "b"
    
    Call AssertError(Err, 2, "a", "b")
    Debug.Assert False ' should raise error
    
errTrap:
    Debug.Assert Err.Number = AssertCode.FailureCode
End Function

Private Function AssertError_WithWrongErrorSource_RaisesError()
    On Error GoTo errTrap
    
    Err.Number = 1
    Err.Source = "a"
    Err.Description = "b"
    
    Call AssertError(Err, 1, "b", "b")
    Debug.Assert False ' should raise error
    
errTrap:
    Debug.Assert Err.Number = AssertCode.FailureCode
End Function

Private Function AssertError_WithWrongDescription_RaisesError()
    On Error GoTo errTrap
    
    Err.Number = 1
    Err.Source = "a"
    Err.Description = "b"
    
    Call AssertError(Err, 1, "a", "c")
    Debug.Assert False ' should raise error
    
errTrap:
    Debug.Assert Err.Number = AssertCode.FailureCode
End Function

Private Function AssertError_WithCorrectError_Passes()
    On Error GoTo errTrap
    
    Err.Number = 1
    Err.Source = "a"
    Err.Description = "b"
    
    Call AssertError(Err, 1, "a", "b")
    Exit Function
    
errTrap:
    Debug.Assert False ' should not raise error
End Function




Private Function IBootstrapTestClass_Run() As Long
    IBootstrapTestClass_Run = Run
End Function