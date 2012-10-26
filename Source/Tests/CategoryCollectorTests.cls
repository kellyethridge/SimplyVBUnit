VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CategoryCollectorTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
' Module: CategoryCollectorTests
'
Option Explicit

Private Store       As CategoryStore
Private Collector   As CategoryCollector

Public Sub Setup()
    Set Store = Sim.NewCategoryStore("store")
    Set Collector = Sim.NewCategoryCollector(Store)
End Sub

Public Sub Constructor_WithNothing_ThrowsException()
    Assert.Throws ErrorCode.ArgumentNull, "CategoryCollector.Init"
    
    Sim.NewCategoryCollector Nothing
End Sub

Public Sub ForTest_WithEmptyTestName_Throws()
    Assert.Throws ErrorCode.Argument, "CategoryCollector.ForTest"
    
    Collector.ForTest "", "c1", "c2", "c3"
End Sub

Public Sub ForTest_WithNoCategories_Throws()
    Assert.Throws ErrorCode.Argument, "CategoryCollector.ForTest"
    
    Collector.ForTest ""
End Sub

Public Sub ForTest_WithNonStringValues_Throws()
    Assert.Throws ErrorCode.InvalidOperation, "CategoryCollector.ForTest"
    
    Collector.ForTest "test", "two", 2
End Sub

Public Sub ForTest_WithMultipleCategories_AddsToTestsCategories()
    Collector.ForTest "abc", "c1", "c2", "c3"
    
    Dim Expected As Variant
    Expected = Array("c1", "c2", "c3")
    Assert.AreEquivalent Expected, Store.GetTestCategorization("abc")
End Sub

Public Sub ForTest_WithSameCategoryTwice_SkipsAdditionalCategory()
    Collector.ForTest "abc", "c1", "c2", "c2"
    
    Dim Expected As Variant
    Expected = Array("c1", "c2")
    Assert.AreEquivalent Expected, Store.GetTestCategorization("abc")
End Sub

Public Sub ForTest_WhenCalledMultipleTimesForSameTest_AddsCategories()
    Collector.ForTest "abc", "c1"
    
    Collector.ForTest "abc", "c2"
    
    Dim Expected As Variant
    Expected = Array("c1", "c2")
    Assert.AreEquivalent Expected, Store.GetTestCategorization("abc")
End Sub

Public Sub ForCategory_WithEmptyCategoryName_Throws()
    Assert.Throws ErrorCode.Argument, "CategoryCollector.ForCategory"
    
    Collector.ForCategory "", "t1"
End Sub

Public Sub ForCategory_WithNoTests_Throws()
    Assert.Throws ErrorCode.Argument, "CategoryCollector.ForCategory"
    
    Collector.ForCategory ""
End Sub

Public Sub ForCategory_WithNonStringValues_Throws()
    Assert.Throws ErrorCode.InvalidOperation, "CategoryCollector.ForCategory"
    
    Collector.ForCategory "c1", "t1", 2
End Sub

Public Sub ForCategory_WithTestName_CreatesCategorization()
    Collector.ForCategory "c1", "t1"
    
    Dim Expected As Variant
    Expected = Array("c1")
    Assert.AreEquivalent Expected, Store.GetTestCategorization("t1")
End Sub

Public Sub ForCategory_WithMultipleTests_CreatesCategorizationForAllTests()
    Collector.ForCategory "c1", "t1", "t2"
    
    Dim Expected As Variant
    Expected = Array("c1")
    Assert.AreEquivalent Expected, Store.GetTestCategorization("t1")
    Assert.AreEquivalent Expected, Store.GetTestCategorization("t2")
End Sub

Public Sub ForFixture_WithNoCategories_Throws()
    Assert.Throws ErrorCode.Argument, "CategoryCollector.ForFixture"
    
    Collector.ForFixture
End Sub

Public Sub ForFixture_WithMultipleCategories_AddsCategories()
    Collector.ForFixture "c1", "c2"
    
    Dim Expected As Variant
    Expected = Array("c1", "c2")
    Assert.AreEquivalent Expected, Store.FixtureCategorization
End Sub

Public Sub ForFixture_WithMultipleCalls_AddsCategories()
    Collector.ForFixture "c1"
    
    Collector.ForFixture "c2"
    
    Dim Expected As Variant
    Expected = Array("c1", "c2")
    Assert.AreEquivalent Expected, Store.FixtureCategorization
End Sub

Public Sub ForFixture_WithDuplicateCategory_SkipsDuplicate()
    Collector.ForFixture "c1"
    
    Collector.ForFixture "c1"
    
    Dim Expected As Variant
    Expected = Array("c1")
    Assert.AreEquivalent Expected, Store.FixtureCategorization
End Sub


