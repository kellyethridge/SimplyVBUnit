VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestingPrimitives"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This set of unit test shows how to assert conditions on standard primitive values.

Public Sub AssertingLongs()
    Const Actual As Long = 1
    Const Expected As Long = 1
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub AssertingALongAndDouble()
    Const Actual As Long = 1
    Const Expected As Double = 1#
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Asserting2Doubles()
    Const Actual As Double = 1.0001
    Const Expected As Double = 1#
    
    Assert.That Actual, Iz.EqualTo(Expected).Within(0.0002)
End Sub

Public Sub AssertingOneLessThanTheOther()
    Const Actual As Long = 1
    Const Other As Long = 2
    
    Assert.That Actual, Iz.LessThan(Other)
    
    WriteLine "Hello kelly"
End Sub
