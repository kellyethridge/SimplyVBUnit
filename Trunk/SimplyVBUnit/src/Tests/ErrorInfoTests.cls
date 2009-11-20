VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ErrorInfoTests"
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
' Module: ErrorInfoTests
'
Option Explicit

Public Sub Constructor_WhenCalled_StoresErrorNumber()
    Dim Info As ErrorInfo
    
    Set Info = Sim.NewErrorInfo(1)
    
    Dim Actual As Long
    Actual = Info.Number
    Assert.IsTrue Actual = 1, "Wrong error number."
End Sub

Public Sub Constructor_WithSource_StoresSource()
    Dim Info As ErrorInfo
    
    Set Info = Sim.NewErrorInfo(1, "source")
    
    Dim Actual As String
    Actual = Info.Source
    Assert.IsTrue Actual = "source", "Wrong source."
End Sub

Public Sub Constructor_WithDescription_StoresDescription()
    Dim Info As ErrorInfo
    
    Set Info = Sim.NewErrorInfo(1, , "description")
    
    Dim Actual As String
    Actual = Info.Description
    Assert.IsTrue Actual = "description", "Wrong description."
End Sub

Public Sub Equals_WithSameValues_ReturnsTrue()
    Dim Info1 As ErrorInfo
    Set Info1 = Sim.NewErrorInfo(1, "source", "desc")
    Dim Info2 As ErrorInfo
    Set Info2 = Sim.NewErrorInfo(1, "source", "desc")
    
    Dim Actual As Boolean
    Actual = Info1.Equals(Info2)
    
    Assert.IsTrue Actual = True, "Should be equal."
End Sub

Public Sub Equals_WithDifferentNumber_ReturnsFalse()
    Dim Info1 As ErrorInfo
    Set Info1 = Sim.NewErrorInfo(1, "source", "desc")
    Dim Info2 As ErrorInfo
    Set Info2 = Sim.NewErrorInfo(2, "source", "desc")
    
    Dim Actual As Boolean
    Actual = Info1.Equals(Info2)
    
    Assert.IsTrue Actual = False, "Should not be equal."
End Sub

Public Sub Equals_WithDifferentSource_ReturnsFalse()
    Dim Info1 As ErrorInfo
    Set Info1 = Sim.NewErrorInfo(1, "source", "desc")
    Dim Info2 As ErrorInfo
    Set Info2 = Sim.NewErrorInfo(1, "source2", "desc")
    
    Dim Actual As Boolean
    Actual = Info1.Equals(Info2)
    
    Assert.IsTrue Actual = False, "Should not be equal."
End Sub

Public Sub Equals_WithDifferentDescription_ReturnsFalse()
    Dim Info1 As ErrorInfo
    Set Info1 = Sim.NewErrorInfo(1, "source", "desc")
    Dim Info2 As ErrorInfo
    Set Info2 = Sim.NewErrorInfo(1, "source", "desc2")
    
    Dim Actual As Boolean
    Actual = Info1.Equals(Info2)
    
    Assert.IsTrue Actual = False, "Should not be equal."
End Sub

Public Sub FromErr_WhenCalled_StoresErrObjectValues()
    Err.Number = 5
    Err.Source = "src"
    Err.Description = "desc"
    
    Dim Actual As ErrorInfo
    Set Actual = ErrorInfo.FromErr(Err)
    Err.Clear
    
    Dim Expected As ErrorInfo
    Set Expected = Sim.NewErrorInfo(5, "src", "desc")
    
    Assert.IsTrue Actual.Equals(Expected), "Should be equal info."
End Sub

Public Sub Equals_WithNothing_ReturnsFalse()
    Dim Info As ErrorInfo
    Set Info = Sim.NewErrorInfo(1)
    
    Dim Actual As Boolean
    Actual = Info.Equals(Nothing)
    
    Assert.IsTrue Actual = False, "Equals should return false."
End Sub


















