VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StubConstraint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IConstraint

Public ActualValue      As Variant
Public ExpectedValue    As Variant
Public Matches          As Boolean
Public Message          As String


Private Function IConstraint_Matches(Actual As Variant) As Boolean
    IConstraint_Matches = Me.Matches
End Function

Private Function IConstraint_Resolve() As SVBUnit.IConstraint
    Set IConstraint_Resolve = Me
End Function

Private Sub IConstraint_WriteActualValueTo(ByVal Writer As SVBUnit.TextMessageWriter)
    Writer.WriteActualValue Me.ActualValue
End Sub

Private Sub IConstraint_WriteDescriptionTo(ByVal Writer As SVBUnit.TextMessageWriter)
    Writer.WriteExpectedValue Me.ExpectedValue
End Sub

Private Sub IConstraint_WriteMessageTo(ByVal Writer As SVBUnit.TextMessageWriter)
    Writer.WriteText Me.Message
End Sub
