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

Public Sub Constructor_WithNothing_ThrowsException()
    Assert.Throws ErrorCode.ArgumentNull, "CategoryCollector.Init"
    
    Sim.NewCategoryCollector Nothing
End Sub

Public Sub AddFixtureCategories_WithNoneDefined_ThrowsException()
    Dim Store As CategoryStore
    Set Store = Sim.NewCategoryStore("fix")
    Dim Collector As CategoryCollector
    Set Collector = Sim.NewCategoryCollector(Store)

    Assert.Throws ErrorCode.Argument, "CategoryCollector.AddFixtureCategories"
    
    Collector.AddFixtureCategories
End Sub

Public Sub AddFixtureCategories_WithValues_AddsToStore()
    Dim Store As CategoryStore
    Set Store = Sim.NewCategoryStore("fix")
    Dim Collector As CategoryCollector
    Set Collector = Sim.NewCategoryCollector(Store)
    
    Collector.AddFixtureCategories "one", "two"
    
    Dim Expected() As Variant
    Expected = Array("one", "two")
    Assert.AreEquivalent Expected, Store.FixtureCategorization
End Sub

Public Sub AddFixtureCategories_WithNonStringValue_ThrowsException()
    Dim Store As CategoryStore
    Set Store = Sim.NewCategoryStore("fix")
    Dim Collector As CategoryCollector
    Set Collector = Sim.NewCategoryCollector(Store)

    Assert.Throws ErrorCode.Argument, "CategoryCollector.AddFixtureCategories"
    Collector.AddFixtureCategories "one", 2
End Sub

Public Sub AddTestCategories_WithEmptyName_ThrowsException()
    Dim Store As CategoryStore
    Set Store = Sim.NewCategoryStore("fix")
    Dim Collector As CategoryCollector
    Set Collector = Sim.NewCategoryCollector(Store)
    
    Assert.Throws ErrorCode.Argument, "CategoryCollector.AddTestCategories"
    Collector.AddTestCategories "", "one", "two"
End Sub

Public Sub AddTestCategories_WithDuplicateName_ThrowsException()
    Dim Store As CategoryStore
    Set Store = Sim.NewCategoryStore("fix")
    Dim Collector As CategoryCollector
    Set Collector = Sim.NewCategoryCollector(Store)
    Collector.AddTestCategories "test", "one"
    
    Assert.Throws ErrorCode.InvalidOperation, "CategoryCollector.AddTestCategories"
    Collector.AddTestCategories "test", "two"
End Sub

Public Sub AddTestCategories_WithNonStringValue_ThrowsException()
    Dim Store As CategoryStore
    Set Store = Sim.NewCategoryStore("fix")
    Dim Collector As CategoryCollector
    Set Collector = Sim.NewCategoryCollector(Store)
    
    Assert.Throws ErrorCode.InvalidOperation, "CategoryCollector.AddTestCategories"
    Collector.AddTestCategories "test", "two", 2
End Sub

Public Sub AddTestCategories_WithNoneDefined_ThrowsException()
    Dim Store As CategoryStore
    Set Store = Sim.NewCategoryStore("fix")
    Dim Collector As CategoryCollector
    Set Collector = Sim.NewCategoryCollector(Store)

    Assert.Throws ErrorCode.Argument, "CategoryCollector.AddTestCategories"
    
    Collector.AddTestCategories "test"
End Sub

Public Sub AddTestCategories_WithValues_AddsToStore()
    Dim Store As CategoryStore
    Set Store = Sim.NewCategoryStore("fix")
    Dim Collector As CategoryCollector
    Set Collector = Sim.NewCategoryCollector(Store)

    Collector.AddTestCategories "test", "one", "two"
    
    Dim Expected() As Variant
    Expected = Array("one", "two")
    Assert.AreEquivalent Expected, Store.TestCategorizations(1)
End Sub


