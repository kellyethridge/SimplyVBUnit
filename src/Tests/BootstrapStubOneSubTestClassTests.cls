VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BootstrapStubOneSubTestClassTests"
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
' Module: BootstrapStubOneSubTestClassTests
'
Option Explicit
Implements IBootstrapTestClass

Private mCount As Long

Public Function Run() As Long
    mCount = 0
    
    Call RunTest(Test1_WhenCalled_DoesNothing)
    Call RunTest(Test1_WhenSetToFail_RaisesCorrectErrorCode)
    Call RunTest(Test1_WithMessageSet_IncludesMessageWithError)
    
    Run = mCount
End Function

Private Sub RunTest(ByRef v As Variant)
    mCount = mCount + 1
End Sub

Private Function Test1_WhenCalled_DoesNothing()
    On Error GoTo errTrap
    Dim TestClass As New Stub1Test
    
    Call TestClass.Test1
    Exit Function
    
errTrap:
    Debug.Assert False
End Function

Private Function Test1_WhenSetToFail_RaisesCorrectErrorCode()
    On Error GoTo errTrap
    Dim TestClass As New Stub1Test
    Call TestClass.Errors.Add(Constants.TEST1_METHOD, AssertCode.FailureCode)
    
    Call TestClass.Test1
    Debug.Assert False ' should have raised an error
    
errTrap:
    Dim Actual As Long
    Actual = Err.Number
    Debug.Assert Actual = AssertCode.FailureCode
End Function

Private Function Test1_WithMessageSet_IncludesMessageWithError()
    Const Expected As String = "message"
    On Error GoTo errTrap
    Dim TestClass As New Stub1Test
    Call TestClass.Errors.Add(Constants.TEST1_METHOD, AssertCode.FailureCode, , Expected)
    
    TestClass.Test1
    
errTrap:
    Dim Actual As String
    Actual = Err.Description
    Debug.Assert Actual = Expected
End Function

Private Function IBootstrapTestClass_Run() As Long
    IBootstrapTestClass_Run = Run
End Function
