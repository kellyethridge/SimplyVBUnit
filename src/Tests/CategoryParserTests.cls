VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CategoryParserTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
' Module: CategoryParserTests
'
Option Explicit

Public Sub Parse_WithNoCategoryDefined_ReturnsNameOnly()
    Dim Parser As New CategoryParser
    Dim Result As CategoryParserResult
    
    Set Result = Parser.Parse("NameOnly")
    
    Dim Actual As String
    Actual = Result.Name
    Assert.That Actual, Iz.EqualTo("NameOnly")
End Sub

Public Sub Parse_WithNoCategoriesDefined_ReturnsNothingList()
    Dim Parser As New CategoryParser
    Dim Result As CategoryParserResult
    
    Set Result = Parser.Parse("NameOnly")

    Dim Actual As CategoryList
    Set Actual = Result.Categories
    Assert.IsTrue Actual Is Nothing
End Sub

Public Sub Parse_With1CategoryDefined_ReturnsCategoryInList()
    Dim Parser As New CategoryParser
    Dim Result As CategoryParserResult
    
    Set Result = Parser.Parse("Name__Category")
    
    Dim Actual As CategoryList
    Set Actual = Result.Categories
    Assert.IsTrue Actual.Contains("Category")
End Sub

Public Sub Parse_With1CategoryDefined_RemovesCategoryFromName()
    Dim Parser As New CategoryParser
    Dim Result As CategoryParserResult
    
    Set Result = Parser.Parse("Name__Category")

    Dim Actual As String
    Actual = Result.Name
    Assert.That Actual, Iz.EqualTo("Name")
End Sub

Public Sub Parse_WithMultipleCategories_ReturnsCategoriesInList()
    Dim Parser As New CategoryParser
    Dim Result As CategoryParserResult
    
    Set Result = Parser.Parse("Name__Cat1_Cat2_Cat3")
    
    Dim Actual As CategoryList
    Set Actual = Result.Categories
    Assert.IsTrue Actual.Contains("Cat1")
    Assert.IsTrue Actual.Contains("Cat2")
    Assert.IsTrue Actual.Contains("Cat3")
End Sub
