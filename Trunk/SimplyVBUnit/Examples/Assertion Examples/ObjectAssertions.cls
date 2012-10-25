VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ObjectAssertions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' ObjectAssertions
'
' This module shows many of the ways assertions can be made against object values.
' Not all assertions will pass. This is just to demonstrate how to perform assertions with objects.
' Where possible, examples using straight assertions and constraint assertions are performed to
' demonstrate how to perform assertions with different methods.
'
Option Explicit

Public Sub IsNothing_WithNullObject_Passes()
    Dim Actual As Object

    Assert.That Actual, Iz.Nothing
    Assert.IsNothing Actual
End Sub

Public Sub IsNotNothing_WithNonNullObject_Passes()
    Dim Actual As New Collection

    Assert.That Actual, Iz.Not.Nothing
    Assert.IsNotNothing Actual
End Sub

Public Sub SameInstance_WithSameObject_Passes()
    Dim Actual      As Collection
    Dim Expected    As Collection
    Set Actual = New Collection
    Set Expected = Actual

    Assert.That Actual, Iz.SameAs(Expected)
    Assert.AreSame Expected, Actual
End Sub
