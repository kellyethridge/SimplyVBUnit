VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NumericAssertions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' NumericAssertions
'
' This module shows many of the ways assertions can be made against numeric values.
' Not all assertions will pass. This is just to demonstrate how to perform assertions with numbers.
' Where possible, examples using straight assertions and constraint assertions are performed to
' demonstrate how to perform assertions with different methods.
'
Option Explicit


'
' Comparison test using Longs
'
Public Sub Equality_WithEqualValues_Passes()
    Const Actual    As Long = 1
    Const Expected  As Long = 1
    
    Assert.That Actual, Iz.EqualTo(Expected), "This should be equal."
    Assert.AreEqual Expected, Actual, "This should be equal."
End Sub

Public Sub Equality_WithNonEqualValues_Fails()
    Const Actual    As Long = 1
    Const Expected  As Long = 2
    
    Assert.AreEqual Expected, Actual, "This should not pass."
End Sub

Public Sub LessThan_WithLesserValue_Passes()
    Const Actual    As Long = 1
    Const Expected  As Long = 2
    
    Assert.That Actual, Iz.LessThan(Expected), "This should be lesser."
    Assert.Less Actual, Expected, "This should be lesser."
End Sub

Public Sub Inequality_WithNonEqualValues_Passes()
    Const Actual    As Long = 1
    Const Expected  As Long = 2
    
    Assert.That Actual, Iz.Not.EqualTo(Expected), "Should not be equal."
    Assert.AreNotEqual Expected, Actual, "Should not be equal."
End Sub


'
' Comparison tests using mixed datatypes.
'
Public Sub Equality_WithLongAndDouble_Passes()
    Const Actual    As Long = 1
    Const Expected  As Double = 1#
    
    Assert.That Actual, Iz.EqualTo(Expected), "Should be equal."
    Assert.That Actual, Iz.EqualTo(Expected), "Should be equal."
End Sub

Public Sub Equality_WithSameValueWithinTolerance_Passes()
    Const Actual    As Double = 1#
    Const Expected  As Double = 1.00001
    
    Assert.That Actual, Iz.EqualTo(Expected).Within(0.0001), "Should be equal."
    Assert.AreEqualFloats Expected, Actual, 0.0001, "Should be equal."
End Sub

Public Sub Equality_WithDoubleAndCurrency_Passes()
    Const Actual    As Double = 5.5
    Const Expected  As Currency = 5.5@
    
    Assert.AreEqual Expected, Actual, "Should be equal."
    Assert.That Actual, Iz.EqualTo(Expected), "Should be equal."
End Sub


'
' Range comparisons using multiple techniques.
'
Public Sub InRange_WithValueWithinRange_Passes()
    Const Actual    As Long = 5
    
    Assert.That Actual, Iz.InRange(2, 8)
    Assert.InRange Actual, 2, 8
    Assert.That Actual, Iz.GreaterThanOrEqualTo(2).AndAlso.LessThanOrEqualTo(8)
End Sub

Public Sub NotInRange_WithValueOutsideRange_Passes()
    Const Actual As Long = 5
    
    Assert.That Actual, Iz.Not.InRange(7, 10)
    Assert.NotInRange Actual, 7, 10
    Assert.That Actual, Iz.LessThan(7).OrElse.GreaterThan(10)
End Sub

Public Sub GreaterThanAndLessThan_WithValueOutsideRange_Fails()
    Const Actual    As Long = 10
    
    Assert.That Actual, Iz.GreaterThan(4).AndAlso.LessThan(9), "This should fail."
End Sub



