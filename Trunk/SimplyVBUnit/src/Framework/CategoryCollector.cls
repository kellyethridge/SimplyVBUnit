VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CategoryCollector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
' Module: CategoryCollector
'
Option Explicit

Private mFixtureCategories As TestCategories
Private mTestCategories As New Collection

Public Sub ForFixture(ParamArray Categories() As Variant)
    Dim Item As Variant
    For Each Item In Categories
        If VarType(Item) = vbString Then
            mFixtureCategories.Add Item
        End If
    Next
End Sub

Public Sub ForTest(ByRef TestName As String, ParamArray Categories() As Variant)
    Dim c As TestCategories
    Set c = Sim.NewTestCategories(Name, False)
    
    Dim Item As Variant
    For Each Item In Categories
        If VarType(Item) = vbString Then
            c.Add Item
        End If
    Next
    
    mTestCategories.Add c, TestName
End Sub



