VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestContextTests"
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
' Module: TestContextTests
'
Option Explicit

Public Sub Log_WhenCalled_StoresLogWriter()
    Dim Fake    As New FakeTextWriter
    Dim c       As New TestContext
    
    Set c.Log = Fake
    
    Dim Actual As ITextWriter
    Set Actual = c.Log
    Assert.IsTrue Actual Is Fake
End Sub

Public Sub Error_WhenCalled_StoresErrorWriter()
    Dim Fake    As New FakeTextWriter
    Dim c       As New TestContext
    
    Set c.Error = Fake
    
    Dim Actual As ITextWriter
    Set Actual = c.Error
    Assert.IsTrue Actual Is Fake
End Sub

Public Sub Out_WhenCalled_StoredOutWriter()
    Dim Fake    As New FakeTextWriter
    Dim c       As New TestContext
    
    Set c.Out = Fake
    
    Dim Actual As ITextWriter
    Set Actual = c.Out
    Assert.IsTrue Actual Is Fake
End Sub

















