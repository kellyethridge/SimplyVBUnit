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
' Module: CategoryCollector
'
Option Explicit

Private mStore As CategoryStore

Public Sub AddFixtureCategories(ParamArray Categories() As Variant)
    If UBound(Categories) < LBound(Categories) Then _
        Error.Argument Me, "AddFixtureCategories", "Categories", "No categories are defined."
    
    Dim Item As Variant
    For Each Item In Categories
        If VarType(Item) <> vbString Then _
            Error.Argument Me, "AddFixtureCategories", "Categories", "Only strings can be passed in."
            
        mStore.FixtureCategorization.Add CStr(Item)
    Next
End Sub

Public Sub AddTestCategories(ByRef TestName As String, ParamArray Categories() As Variant)
    If Len(TestName) = 0 Then _
        Error.CannotBeEmpty Me, "AddTestCategories", "TestName"
    If ContainsKey(mStore.TestCategorizations, TestName) Then _
        Error.Argument Me, "AddTestCategories", "TestName", "Categories for test <" & TestName & "> already exist."
    If UBound(Categories) < LBound(Categories) Then _
        Error.Argument Me, "AddTestCategories", "Categories", "No categories are defined."
        
    Dim c As Categorization
    Set c = Sim.NewCategorization(TestName, False)
    
    Dim Item As Variant
    For Each Item In Categories
        If VarType(Item) <> vbString Then _
            Error.Argument Me, "AddTestCategories", "Categories", "Only strings can be passed in."
            
        c.Add CStr(Item)
    Next
    
    mStore.TestCategorizations.Add c, TestName
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Store As CategoryStore)
    If Store Is Nothing Then _
        Error.CannotBeNothing Me, "Init", "Store"
        
    Set mStore = Store
End Sub



