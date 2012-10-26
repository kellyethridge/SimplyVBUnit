VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArrayAssertions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' This class demonstrates various ways arrays can be tested
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "Array Tests"
End Sub

' tests two identical arrays for equality by checking
' the lower and upper bounds and then each element.
Public Sub TestTwoIdenticalArraysForEquality()
    Dim a1() As Long
    Dim a2() As Long
    
    a1 = CreateLongs(100)
    a2 = a1
    
    Assert.That a1, Iz.EqualTo(a2)
End Sub

Public Sub IsEmpty_WithNullArray_Passes()
    Dim Actual() As Long
    
    Assert.That Actual, Iz.Empty
    Assert.IsEmpty Actual
End Sub

' tests an array for being sorted. the only datatypes that
' are supported are numbers, strings and dates.
Public Sub TestThatArrayIsSorted()
    Dim a() As Long
    a = CreateLongs(100)
    
    Assert.That a, Iz.Ordered
End Sub

' tests an array for being sorted in descending order. the only datatypes that
' are supported are numbers, strings and dates.
Public Sub TestsThatArraysIsSortedInDescendingOrder()
    Dim a As Variant
    a = Array(5, 4, 3, 2, 1, 0)
    
    Assert.That a, Iz.Ordered.Descending
    Assert.OrderedDescending a
End Sub

' this applies a constraint to each element, ensuring each element
' passes the constraint, checking each element is equal to 1.
Public Sub TestThatArrayContainsOnlyTheValue1()
    Dim a As Variant
    a = Array(1, 1, 1, 1, 1)
    
    Assert.That a, Has.All.EqualTo(1)
End Sub

' this applies a constraint to each element, ensuring each element
' passes the constraint, this time checking for references of Nothing.
Public Sub TestEachElementToBeingNothing()
    Dim a() As Object
    ReDim a(0 To 10)
    
    Assert.That a, Has.All.Nothing
End Sub

' tests two arrays for equality using only the number of elements
' and the elements themselves. the bounds are ignored.
Public Sub TestTwoArraysBySizeAndElementsOnly()
    Dim a1() As Long
    ReDim a1(0 To 4)
    
    Dim a2() As Long
    ReDim a2(5 To 9)
    
    Dim i As Long
    For i = 0 To UBound(a1)
        a1(i) = i
        a2(i + 5) = i
    Next i
    
    Assert.That a2, Iz.EqualTo(a1).AsCollection
End Sub

' tests that two arrays contain the same elements, ignoring the order.
Public Sub TestTwoArraysForEquivalence()
    Dim a1 As Variant
    Dim a2 As Variant
    
    a1 = Array(1, 2, 3, 4, 5)
    a2 = Array(3, 4, 1, 2, 5)
    
    Assert.That a1, Iz.EquivalentTo(a2)
End Sub

' checks that there are no elements in an array.
Public Sub TestForAnEmptyArray()
    Dim a As Variant
    a = Array()
    
    Assert.That a, Iz.Empty
    Assert.IsEmpty a
End Sub

' checks that at least one element is equal to the value specified in Member.
Public Sub TestThatAnArrayConstainsASpecificValue()
    Dim a As Variant
    a = Array(1, 2, 3, 4, 5, 6)
    
    Assert.That a, Has.Member(4)
End Sub

Public Sub TestThatTheArrayIsASubsetOfAnotherArray()
    Dim a As Variant
    a = Array(1, 5, 3)
    
    Assert.That a, Iz.SubsetOf(Array(1, 2, 3, 4, 5))
End Sub

Public Sub TestWrongValue()
    Dim a As Variant
    a = Array(0, 1, 2, 3, 3, 5)
    
    Assert.That a, Iz.EqualTo(CreateLongs(6))
End Sub


Private Function CreateLongs(ByVal Count As Long) As Long()
    Dim Result() As Long
    ReDim Result(0 To Count - 1)
    
    Dim i As Long
    For i = 0 To Count - 1
        Result(i) = i
    Next i
    
    CreateLongs = Result
End Function


