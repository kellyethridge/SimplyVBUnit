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
