VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OddEvenEqualityComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IEqualityComparer

Private Function IEqualityComparer_Equals(X As Variant, Y As Variant) As Boolean
    On Error GoTo errTrap
    Dim XLong As Long
    Dim YLong As Long
    XLong = CLng(X)
    YLong = CLng(Y)
    
    IEqualityComparer_Equals = (IsOdd(X) = IsOdd(Y))
    
errTrap:
End Function

Private Function IsOdd(ByVal X As Long) As Boolean
    IsOdd = (X And 1) = 1
End Function
