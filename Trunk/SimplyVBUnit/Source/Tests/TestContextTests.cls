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

Public Sub Constructor_WhenCalled_BeginsNewContext()
    Dim FakeOut As New FakeTextWriter
    
    With New TestContext
        Set .Out = FakeOut
    
        Dim Actual As ITextWriter
        Set Actual = TestContext.Out
    End With
    
    Assert.IsTrue Actual Is FakeOut
End Sub

Public Sub Destructor_WhenCalled_RevertsToPreviousContext()
    Dim FirstOut        As New FakeTextWriter
    Dim SecondOut       As New FakeTextWriter
    Dim FirstActual     As ITextWriter
    Dim SecondActual    As ITextWriter
    
    With New TestContext
        Set .Out = FirstOut
        
        With New TestContext
            Set .Out = SecondOut
            Set SecondActual = TestContext.Out
        End With
        
        Set FirstActual = TestContext.Out
    End With
    
    Assert.IsTrue FirstActual Is FirstOut
    Assert.IsTrue SecondActual Is SecondOut
End Sub















