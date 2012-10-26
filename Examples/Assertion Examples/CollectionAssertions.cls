VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CollectionAssertions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' CollectionAssertions
'
' This module shows many of the ways assertions can be made against collections.
' Not all assertions will pass. This is just to demonstrate how to perform assertions with collections.
' Where possible, examples using straight assertions and constraint assertions are performed to
' demonstrate how to perform assertions with different methods.
'
Option Explicit
Implements ICategorizable


Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "Collection Tests"
End Sub

Public Sub IsEmpty_WithEmptyCollection_Passes()
    Dim Actual As New Collection
    
    Assert.That Actual, Iz.Empty
    Assert.IsEmpty Actual
End Sub

Public Sub EquivalentTo_WithCollectionsContainingSameValuesInDifferentOrder_Passes()
    Dim Actual As New Collection
    Actual.Add "a"
    Actual.Add "b"
    Dim Expected As New Collection
    Expected.Add "b"
    Expected.Add "a"
    
    Assert.That Actual, Iz.EquivalentTo(Expected)
    Assert.AreEquivalent Expected, Actual
End Sub

Public Sub HasMember_WithCollectionContainingValue_Passes()
    Dim Actual As New Collection
    Actual.Add "a"
    Actual.Add "b"
    
    Assert.That Actual, Has.Member("b")
End Sub

Public Sub HasMember_WithCollectionNotContainingValue_Fails()
    Dim Actual As New Collection
    Actual.Add "a"

    Assert.That Actual, Has.Member("b"), "This should fail."
End Sub

Public Sub HasMember_WithArrayContainingValue_Passes()
    Dim Actual(0 To 100) As Long
    Actual(50) = 1000
    
    Assert.That Actual, Has.Member(1000)
End Sub

Public Sub Ordered_WithOrderedCollection_Passes()
    Dim Actual As New Collection
    Actual.Add "a"
    Actual.Add "b"
    
    Assert.That Actual, Iz.Ordered
    Assert.Ordered Actual
End Sub

Public Sub Ordered_WithUnorderedCollection_Fails()
    Dim Actual As New Collection
    Actual.Add "b"
    Actual.Add "a"
    
    Assert.That Actual, Iz.Ordered
End Sub

Public Sub DescendingOrder_WithDescendingOrderedArray_Passes()
    Dim Actual(0 To 1) As String
    Actual(0) = "b"
    Actual(1) = "a"
    
    Assert.That Actual, Iz.Ordered.Descending
    Assert.OrderedDescending Actual
End Sub

Public Sub AssertingAllMembers_WithPassingValues_Passes()
    Dim Actual As New Collection
    Actual.Add "c"
    Actual.Add "C"
    
    Assert.That Actual, Has.All.EqualTo("c").IgnoreCase
End Sub

Public Sub AssertingSomeMembers_WithAtLeastOnePassing_Passes()
    Dim Actual As New Collection
    Actual.Add 1
    Actual.Add 2
    Actual.Add 3
    
    Assert.That Actual, Has.Some.GreaterThan(1)
End Sub

Public Sub UniqueMembers_WithDifferentMembers_Passes()
    Dim Actual As New Collection
    Actual.Add 1
    Actual.Add 2
    Actual.Add 3
    
    Assert.That Actual, Iz.Unique
    Assert.Unique Actual
End Sub

Public Sub UniqueMembers_WithNonUniqueMembers_Fails()
    Dim Actual As New Collection
    Actual.Add 1
    Actual.Add 2
    Actual.Add 2
    
    Assert.That Actual, Iz.Unique
End Sub

Public Sub None_WithNoExpectedMembers_Passes()
    Dim Actual(0 To 2) As Long
    Actual(0) = 1
    Actual(1) = 2
    Actual(2) = 3
    
    Assert.That Actual, Has.None.GreaterThan(3)
End Sub

