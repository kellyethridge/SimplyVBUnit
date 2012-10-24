VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Categorization"
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
' Module: Categorization
'
Option Explicit

Private mIsFixture  As Boolean
Private mName       As String
Private mCategories As New Collection


Public Property Get IsFixture() As Boolean
    IsFixture = mIsFixture
End Property

Public Property Get Name() As String
    Name = mName
End Property

Public Sub Add(ByRef Category As String)
    If Len(Category) = 0 Then _
        Error.CannotBeEmpty Me, "Add", "Category"
    If ContainsKey(mCategories, Category) Then _
        Error.Argument Me, "Add", "Category", "Category <" & Category & "> already exists."
        
    mCategories.Add Category, Category
End Sub

Public Property Get Item(ByRef Index As Variant) As String
Attribute Item.VB_UserMemId = 0
    Item = mCategories(Index)
End Property

Public Property Get Count() As Long
    Count = mCategories.Count
End Property

Public Function Contains(ByRef Category As String) As Boolean
    Contains = ContainsKey(mCategories, Category)
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCategories.[_NewEnum]
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef Name As String, ByVal IsFixture As Boolean)
    If Len(Name) = 0 Then _
        Error.CannotBeEmpty Me, "Init", "Name"
        
    mName = Name
    mIsFixture = IsFixture
End Sub
