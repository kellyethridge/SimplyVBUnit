VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BootstrapCallErrorTests"
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
' Module: BootstrapCallErrorTests
'
Option Explicit
Implements IBootstrapTestClass


Public Sub Run()
    Raise_WithNonexistingMethod_DoesNothing
    Raise_WithExistingError_RaisesError
    Raise_WithExistingError_RaisesError2
    Raise_WithExistingError_RaisesError3
End Sub

Private Sub Raise_WithNonexistingMethod_DoesNothing()
    On Error GoTo errTrap
    
    Dim Errors As New CallError
    
    Errors.Raise "missing"
    Exit Sub
    
errTrap:
    Debug.Assert False
End Sub

Private Sub Raise_WithExistingError_RaisesError()
    On Error GoTo errTrap
    
    Dim Errors As New CallError
    Errors.Add "Test", 1
    
    Errors.Raise "Test"
    Debug.Assert False ' should raise an error
    
errTrap:
    Debug.Assert Err.Number = 1
    Err.Clear
End Sub

Private Sub Raise_WithExistingError_RaisesError2()
    On Error GoTo errTrap
    
    Dim Errors As New CallError
    Errors.Add "Test", 1, "source"
    
    Errors.Raise "Test"
    Debug.Assert False ' should raise an error
    
errTrap:
    Debug.Assert Err.Source = "source"
    Err.Clear
End Sub

Private Sub Raise_WithExistingError_RaisesError3()
    On Error GoTo errTrap
    
    Dim Errors As New CallError
    Errors.Add "Test", 1, "source", "desc"
    
    Errors.Raise "Test"
    Debug.Assert False ' should raise an error
    
errTrap:
    Debug.Assert Err.Description = "desc"
    Err.Clear
End Sub


Private Sub IBootstrapTestClass_Run()
    Run
End Sub
