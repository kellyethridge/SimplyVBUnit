VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CategoryStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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
' Module: CategoryStore
'
Option Explicit

Private mFixtureCategorization  As Categorization
Private mTestCategorizations    As New Collection


Public Property Get FixtureCategorization() As Categorization
    Set FixtureCategorization = mFixtureCategorization
End Property

Public Property Get TestCategorizations() As Collection
    Set TestCategorizations = mTestCategorizations
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef FixtureName As String)
    If Len(FixtureName) = 0 Then _
        Error.CannotBeEmpty Me, "Init", "FixtureName"
        
    Set mFixtureCategorization = Sim.NewCategorization(FixtureName, True)
End Sub
