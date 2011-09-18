VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CategorizationTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Copyright 2011 Kelly Ethridge
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
' Module: CategorizationTests
'
Option Explicit

Public Sub Constructor_WithEmptyName_ThrowsException()
    Assert.Throws ErrorCode.Argument, "Categorization.Init"
    
    Dim c As Categorization
    
    Set c = Sim.NewCategorization("", False)
End Sub

Public Sub Constructor_WhenCalled_StoresParameters()
    Dim c As Categorization
    
    Set c = Sim.NewCategorization("abc", True)
    
    Assert.AreEqual "abc", c.Name
    Assert.IsTrue c.IsFixture
End Sub

Public Sub Add_WithEmptyCategory_ThrowsException()
    Dim c As Categorization
    Set c = Sim.NewCategorization("a", False)
    
    Assert.Throws ErrorCode.Argument, "Categorization.Add"
    
    c.Add ""
End Sub

Public Sub Add_WithDuplicateCategory_ThrowsException()
    Dim c As Categorization
    Set c = Sim.NewCategorization("a", False)
    c.Add "test"
    
    Assert.Throws ErrorCode.Argument, "Categorization.Add"
    
    c.Add "test"
End Sub

Public Sub Add_WhenCalled_StoresCategory()
    Dim c As Categorization
    Set c = Sim.NewCategorization("a", False)
    
    c.Add "test"
    
    Assert.AreEqual "test", c(1)
End Sub
