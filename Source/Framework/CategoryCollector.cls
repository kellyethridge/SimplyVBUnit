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
'The MIT License (MIT)
'Copyright (c) 2012 Kelly Ethridge
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights to
'use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of
'the Software, and to permit persons to whom the Software is furnished to do so,
'subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
'INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
'PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
'FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
'OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
'DEALINGS IN THE SOFTWARE.
'
'
' Module: SComponent.CategoryCollector
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



