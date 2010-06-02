VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IzStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IzSyntaxHelper


Private Function NewExp() As IConstraintExpression
    Set NewExp = New ConstraintExpression
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IzSyntaxHelper Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IzSyntaxHelper_EqualTo(Expected As Variant) As SimplyVBUnitType.IEqualConstraint
    Set IzSyntaxHelper_EqualTo = NewExp.EqualTo(Expected)
End Function

Private Function IzSyntaxHelper_False() As SimplyVBUnitType.IBasicExpression
    Set IzSyntaxHelper_False = NewExp.False
End Function

Private Function IzSyntaxHelper_GreaterThan(Expected As Variant) As SimplyVBUnitType.IComparisonConstraint
    Set IzSyntaxHelper_GreaterThan = NewExp.GreaterThan(Expected)
End Function

Private Function IzSyntaxHelper_GreaterThanOrEqualTo(Expected As Variant) As SimplyVBUnitType.IComparisonConstraint
    Set IzSyntaxHelper_GreaterThanOrEqualTo = NewExp.GreaterThanOrEqualTo(Expected)
End Function

Private Function IzSyntaxHelper_InRange(From As Variant, To As Variant) As SimplyVBUnitType.IUsingExpression
    Set IzSyntaxHelper_InRange = NewExp.InRange(From, [To])
End Function

Private Function IzSyntaxHelper_LessThan(Expected As Variant) As SimplyVBUnitType.IComparisonConstraint
    Set IzSyntaxHelper_LessThan = NewExp.LessThan(Expected)
End Function

Private Function IzSyntaxHelper_LessThanOrEqualTo(Expected As Variant) As SimplyVBUnitType.IComparisonConstraint
    Set IzSyntaxHelper_LessThanOrEqualTo = NewExp.LessThanOrEqualTo(Expected)
End Function

Private Function IzSyntaxHelper_Missing() As SimplyVBUnitType.IBasicExpression
    Set IzSyntaxHelper_Missing = NewExp.Missing
End Function

Private Function IzSyntaxHelper_Not() As SimplyVBUnitType.IConstraintExpression
    Set IzSyntaxHelper_Not = NewExp.Not
End Function

Private Function IzSyntaxHelper_Nothing() As SimplyVBUnitType.IBasicExpression
    Set IzSyntaxHelper_Nothing = NewExp.Nothing
End Function

Private Function IzSyntaxHelper_True() As SimplyVBUnitType.IBasicExpression
    Set IzSyntaxHelper_True = NewExp.True
End Function

Private Function IzSyntaxHelper_VBEmpty() As SimplyVBUnitType.IBasicExpression
    Set IzSyntaxHelper_VBEmpty = NewExp.vbEmpty
End Function

Private Function IzSyntaxHelper_VBNull() As SimplyVBUnitType.IBasicExpression
    Set IzSyntaxHelper_VBNull = NewExp.vbNull
End Function
