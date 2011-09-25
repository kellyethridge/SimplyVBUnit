VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CategoryFilterTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Match_WithNoTestCategories_ReturnsFalse()
    Dim Test As New FakeTest
    Dim Filter As New CategoryFilter
    Filter.Add "a"
    
    Dim Actual As Boolean
    Actual = Filter.Match(Test)
    
    Assert.IsFalse Actual
End Sub

Public Sub Match_WithNonMatchingTestCategories_ReturnsFalse()
    Dim Test As New FakeTest
    Test.Categories.Add "a"
    Dim Filter As New CategoryFilter
    Filter.Add "b"
    
    Dim Actual As Boolean
    Actual = Filter.Match(Test)
    
    Assert.IsFalse Actual
End Sub

Public Sub Match_WithMatchingCategory_ReturnsTrue()
    Dim Test As New FakeTest
    Test.Categories.Add "a"
    Test.Categories.Add "b"
    Dim Filter As New CategoryFilter
    Filter.Add "b"
    
    Dim Actual As Boolean
    Actual = Filter.Match(Test)
    
    Assert.IsTrue Actual
End Sub
