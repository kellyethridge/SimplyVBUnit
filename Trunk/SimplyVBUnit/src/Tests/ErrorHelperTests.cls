VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ErrorHelperTests"
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
' Module: ErrorHelperTests
'
Option Explicit

Public Sub CannotBeNothing_WithValidArguments_RaisesCorrectError()
    On Error GoTo errTrap
    Dim Helper As New ErrorHelper
    
    Call Helper.CannotBeNothing(Me, "Method", "Object")
    On Error GoTo 0
    
errTrap:
    Call AssertError(Err, ErrorCode.ArgumentNull, "ErrorHelperTests.Method", "Object cannot be Nothing.")
End Sub

Public Sub CannotBeNothing_WithSourceObjectAsNothing_RaisesError()
    On Error GoTo errTrap
    Dim Helper As New ErrorHelper
    
    Call Helper.CannotBeNothing(Nothing, "Method", "Object")
    On Error GoTo 0
    
errTrap:
    Call AssertError(Err, ErrorCode.ArgumentNull, "ErrorHelper.CannotBeNothing", "SourceObject cannot be Nothing.")
End Sub

Public Sub CannotBeEmpty_WithValidArguments_RaisesCorrectError()
    On Error GoTo errTrap
    Dim Helper As New ErrorHelper
    
    Call Helper.CannotBeEmpty(Me, "Method", "String")
    On Error GoTo 0
    
errTrap:
    Call AssertError(Err, ErrorCode.Argument, "ErrorHelperTests.Method", "String cannot be empty.")
End Sub

Public Sub CannotBeEmpty_WithSourceObjectAsNothing_RaisesError()
    On Error GoTo errTrap
    Dim Helper As New ErrorHelper
    
    Call Helper.CannotBeEmpty(Nothing, "Method", "Object")
    On Error GoTo 0
    
errTrap:
    Call AssertError(Err, ErrorCode.ArgumentNull, "ErrorHelper.CannotBeEmpty", "SourceObject cannot be Nothing.")
End Sub

Public Sub CannotBeNothing_WithEmptySourceMethod_RaisesError()
    On Error GoTo errTrap
    Dim Helper As New ErrorHelper
    
    Call Helper.CannotBeNothing(Me, "", "Object")
    On Error GoTo 0
    
errTrap:
    Call AssertError(Err, ErrorCode.Argument, "ErrorHelper.CannotBeNothing", "SourceMethod cannot be empty.")
End Sub

Public Sub CannotBeNothing_WithEmptyArgName_RaisesError()
    On Error GoTo errTrap
    Dim Helper As New ErrorHelper
    
    Call Helper.CannotBeNothing(Me, "Method", "")
    On Error GoTo 0
    
errTrap:
    Call AssertError(Err, ErrorCode.Argument, "ErrorHelper.CannotBeNothing", "ArgName cannot be empty.")
End Sub

Public Sub CannotBeEmpty_WithEmptySourceMethod_RaisesError()
    On Error GoTo errTrap
    Dim Helper As New ErrorHelper
    
    Call Helper.CannotBeEmpty(Me, "", "Object")
    On Error GoTo 0
    
errTrap:
    Call AssertError(Err, ErrorCode.Argument, "ErrorHelper.CannotBeEmpty", "SourceMethod cannot be empty.")
End Sub

Public Sub CannotBeEmpty_WithEmptyArgName_RaisesError()
    On Error GoTo errTrap
    Dim Helper As New ErrorHelper

    Call Helper.CannotBeEmpty(Me, "Method", "")
    On Error GoTo 0

errTrap:
    Call AssertError(Err, ErrorCode.Argument, "ErrorHelper.CannotBeEmpty", "ArgName cannot be empty.")
End Sub

Public Sub CannotBeNegative_WhenCalled_RaisesCorrectError()
    On Error GoTo errTrap
    Dim Helper As New ErrorHelper
    
    Call Helper.CannotBeNegative(Me, "Method", "Index")
    On Error GoTo 0
    
errTrap:
    Call AssertError(Err, ErrorCode.ArgumentOutOfRange, "ErrorHelperTests.Method", "Index cannot be negative.")
End Sub

Public Sub IndexOutOfRange_WhenCalled_RaisesCorrectError()
    On Error GoTo errTrap
    Dim Helper As New ErrorHelper
    
    Call Helper.IndexOutOfRange(Me, "Method")
    On Error GoTo 0
    
errTrap:
    Call AssertError(Err, ErrorCode.IndexOutOfRange, "ErrorHelperTests.Method", "Index was out of range. Must be non-negative and less than the size of the collection.")
End Sub

Public Sub DoesNotSupportEnumeration_WhenCalled_RaisesCorrectError()
    On Error GoTo errTrap
    Dim Helper As New ErrorHelper
    
    Call Helper.EnumerationNotSupported(Me, "Method", "Object")
    On Error GoTo 0
    
errTrap:
    Call AssertError(Err, ErrorCode.Argument, "ErrorHelperTests.Method", "Object does not support enumeration.")
End Sub
























