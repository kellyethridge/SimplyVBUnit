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
' Module: CategoryCollector
'
Option Explicit

Private mStore As CategoryStore


Public Sub ForFixture(ParamArray Categories() As Variant)
    If UBound(Categories) < LBound(Categories) Then _
        Error.Argument Me, "ForFixture", "Categories", "No categories are defined."
    
    Dim Item As Variant
    For Each Item In Categories
        If VarType(Item) <> vbString Then _
            Error.Argument Me, "ForTest", "Categories", "Only strings can be passed in."
            
        Dim Category As String
        Category = Item
        If Not mStore.FixtureCategorization.Contains(Category) Then
            mStore.FixtureCategorization.Add Category
        End If
    Next
End Sub

Public Sub ForTest(ByRef TestName As String, ParamArray Categories() As Variant)
    If Len(TestName) = 0 Then _
        Error.CannotBeEmpty Me, "ForTest", "TestName"
    If UBound(Categories) < LBound(Categories) Then _
        Error.Argument Me, "ForTest", "Categories", "No categories are defined."
    
    Dim Categorization As Categorization
    Set Categorization = GetCategorization(TestName)
    
    Dim Item As Variant
    For Each Item In Categories
        If VarType(Item) <> vbString Then _
            Error.Argument Me, "ForTest", "Categories", "Only strings can be passed in."
        
        Dim CategoryName As String
        CategoryName = Item
        If Not Categorization.Contains(CategoryName) Then
            Categorization.Add CategoryName
        End If
    Next
End Sub

Public Sub ForCategory(ByRef CategoryName As String, ParamArray Tests() As Variant)
    If Len(CategoryName) = 0 Then _
        Error.CannotBeEmpty Me, "ForCategory", "CategoryName"
    If UBound(Tests) < LBound(Tests) Then _
        Error.Argument Me, "ForCategory", "Tests", "No tests are defined."
    
    Dim Item As Variant
    For Each Item In Tests
        If VarType(Item) <> vbString Then _
            Error.Argument Me, "ForCategory", "Tests", "Only strings can be passed in."
        
        ForTest CStr(Item), CategoryName
    Next
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Store As CategoryStore)
    If Store Is Nothing Then _
        Error.CannotBeNothing Me, "Init", "Store"
        
    Set mStore = Store
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetCategorization(ByRef TestName As String) As Categorization
    Dim Result As Categorization
    If mStore.Contains(TestName) Then
        Set Result = mStore.GetTestCategorization(TestName)
    Else
        Set Result = Sim.NewCategorization(TestName, False)
        mStore.TestCategorizations.Add Result, TestName
    End If
    
    Set GetCategorization = Result
End Function

