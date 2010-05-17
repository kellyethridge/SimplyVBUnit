VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CategoryParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
' Module: CategoryParser
'
Option Explicit

Private Const CATEGORY_DELIMETER        As String = "_"
Private Const CATEGORYGROUP_DELIMETER   As String = CATEGORY_DELIMETER & CATEGORY_DELIMETER


Private mName   As String
Private mResult As CategoryParserResult


Public Function Parse(ByRef Name As String) As CategoryParserResult
    Init Name
    ParseCategoriesAndName
    Set Parse = Release
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Init(ByRef Name As String)
    mName = Name
    Set mResult = New CategoryParserResult
End Sub

Private Sub ParseCategoriesAndName()
    Dim i As Long
    i = GetCategoryDelimeterIndex
    
    If HasCategories(i) Then
        AddCategories i
        ParseNameFromCategoryGroup i
    Else
        InitName
    End If
End Sub

Private Function Release() As CategoryParserResult
    Set Release = mResult
    Set mResult = Nothing
    mName = vbNullString
End Function

Private Function GetCategoryDelimeterIndex() As Long
    GetCategoryDelimeterIndex = InStr(mName, CATEGORYGROUP_DELIMETER)
End Function

Private Function HasCategories(ByVal DelimeterIndex As Long) As Boolean
    HasCategories = (DelimeterIndex > 0)
End Function

Private Sub AddCategories(ByVal DelimeterIndex As Long)
    Dim Categories() As String
    Categories = ParseCategoriesFromName(DelimeterIndex)
    
    CreateEmptyCategoryList
    AddCategoriesToList Categories
End Sub

Private Function ParseCategoriesFromName(ByVal DelimeterIndex As Long) As String()
    ParseCategoriesFromName = Split(Mid$(mName, DelimeterIndex + 2), CATEGORY_DELIMETER)
End Function

Private Sub CreateEmptyCategoryList()
    Set mResult.Categories = New CategoryList
End Sub

Private Sub AddCategoriesToList(ByRef Categories() As String)
    Dim i As Long
    For i = 0 To UBound(Categories)
        mResult.Categories.Add Categories(i)
    Next
End Sub

Private Sub ParseNameFromCategoryGroup(ByVal DelimeterIndex As Long)
    mResult.Name = Left(mName, DelimeterIndex - 1)
End Sub

Private Sub InitName()
    mResult.Name = mName
End Sub

