VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MockConstraint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IConstraint
Implements IResolvable

Public Calls    As New CallTrace
Public Matches  As Boolean
Public Actual   As Variant


Private Function IConstraint_Matches(Actual As Variant) As Boolean
    Me.Actual = Actual
    Calls.Add "Matches"
    IConstraint_Matches = Me.Matches
End Function

Private Sub IConstraint_WriteActualValueTo(ByVal Writer As SVBUnit.TextMessageWriter)
    Calls.Add "WriteActualValueTo"
End Sub

Private Sub IConstraint_WriteDescriptionTo(ByVal Writer As SVBUnit.TextMessageWriter)
    Calls.Add "WriteDescriptionTo"
End Sub

Private Sub IConstraint_WriteMessageTo(ByVal Writer As SVBUnit.TextMessageWriter)
    Calls.Add "WriteMessageTo"
End Sub

Private Function IResolvable_Resolve() As SVBUnit.IConstraint
    Calls.Add "Resolve"
    Set IResolvable_Resolve = Me
End Function
