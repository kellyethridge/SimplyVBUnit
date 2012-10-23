VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CategoryListController"
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
' Module: SComponent.CategoryListController
'
Option Explicit

Private Const SETTING_SELECTEDCATEGORIES    As String = "Categories.Selected"
Private Const SETTING_CATEGORIESENABLED     As String = "Categories.Enabled"
Private Const SETTING_CATEGORIESEXCLUDED    As String = "Categories.ExcludeSelected"


Private mTests      As ITest
Private mCategories As New Collection
Private mList       As ListBox
Private mEnabled    As CheckBox
Private mExclude    As CheckBox
Private WithEvents mSelectAll As CheckBox
Attribute mSelectAll.VB_VarHelpID = -1


Public Function CreateFilter() As ITestFilter
    If mEnabled.Value = vbChecked Then
        Dim Included As Boolean
        Included = Not (mExclude.Value = vbChecked)
        
        Dim Filter As New CategoryFilter
        Dim i As Long
        For i = 0 To mList.ListCount - 1
            If mList.Selected(i) = Included Then
                Filter.Add mList.List(i)
            End If
        Next
        
        Set CreateFilter = Filter
    End If
End Function

Public Sub SaveState(ByVal Config As UIConfiguration)
    Config.SetValue SETTING_CATEGORIESENABLED, mEnabled.Value = vbChecked
    Config.SetValue SETTING_CATEGORIESEXCLUDED, mExclude.Value = vbChecked
    
    Dim Setting As UISetting
    Dim Settings As New Collection
    Dim i As Long
    For i = 0 To mList.ListCount - 1
        If mList.Selected(i) Then
            Settings.Add UI.NewUISetting(mList.List(i), mList.List(i))
        End If
    Next
    
    Config.SetSettings SETTING_SELECTEDCATEGORIES, Settings
End Sub

Public Sub LoadState(ByVal Config As UIConfiguration)
    mEnabled.Value = IIf(Config.GetValue(SETTING_CATEGORIESENABLED, False), vbChecked, vbUnchecked)
    mExclude.Value = IIf(Config.GetValue(SETTING_CATEGORIESEXCLUDED, False), vbChecked, vbUnchecked)
    
    Dim Selected As New Collection
    Dim Settings As Collection
    Set Settings = Config.GetSettings(SETTING_SELECTEDCATEGORIES)
    Dim Setting As UISetting
    For Each Setting In Settings
        Selected.Add Setting.Value, Setting.Value
    Next
    
    On Error Resume Next
    Dim i As Long
    For i = 0 To mList.ListCount - 1
        Dim Value As Variant
        Value = Selected(mList.List(i))
        If Err.Number = 0 Then
            mList.Selected(i) = True
            Err.Clear
        End If
    Next i
    
    On Error GoTo 0
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Factory Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Tests As ITest, ByVal List As ListBox, ByVal Enabled As CheckBox, ByVal Exclude As CheckBox, ByVal SelectAll As CheckBox)
    Set mTests = Tests
    Set mList = List
    Set mEnabled = Enabled
    Set mExclude = Exclude
    Set mSelectAll = SelectAll
    Load
End Sub

Private Sub Load()
    mList.Enabled = False
    LoadTest mTests
    mList.Enabled = True
End Sub

Private Sub LoadTest(ByVal Test As ITest)
    If Test.HasCategories Then
        AddCategories Test.Categories
    End If
    
    If Not Test.Tests Is Nothing Then
        Dim ChildTest As ITest
        For Each ChildTest In Test.Tests
            LoadTest ChildTest
        Next
    End If
End Sub

Private Sub AddCategories(ByVal Categories As Categorization)
    Dim Category As Variant
    For Each Category In Categories
        If CanAddCategory(Category) Then
            mList.AddItem Category
        End If
    Next
End Sub

Private Function CanAddCategory(ByRef Category As Variant) As Boolean
    On Error GoTo errTrap
    Dim v As Variant
    v = mCategories(Category)
    Exit Function
        
errTrap:
    mCategories.Add Category, Category
    CanAddCategory = True
End Function

Private Sub mSelectAll_Click()
    Dim IsChecked As Boolean
    IsChecked = mSelectAll.Value = vbChecked
    
    Dim i As Long
    For i = 0 To mList.ListCount - 1
        mList.Selected(i) = IsChecked
    Next
End Sub
