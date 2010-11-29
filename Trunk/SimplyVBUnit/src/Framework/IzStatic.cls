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
Private Function IzSyntaxHelper_Empty() As SimplyVBUnitType.IBasicExpression
    Set IzSyntaxHelper_Empty = NewExp.Empty
End Function

Private Function IzSyntaxHelper_EqualTo(Expected As Variant) As SimplyVBUnitType.IEqualityExpression
    Set IzSyntaxHelper_EqualTo = NewExp.EqualTo(Expected)
End Function

Private Function IzSyntaxHelper_EquivalentTo(Expected As Variant) As SimplyVBUnitType.IEqualityExpression
    Set IzSyntaxHelper_EquivalentTo = NewExp.EquivalentTo(Expected)
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

Private Function IzSyntaxHelper_InRange(FromValue As Variant, ToValue As Variant) As IComparisonConstraint
    Set IzSyntaxHelper_InRange = NewExp.InRange(FromValue, ToValue)
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

Private Function IzSyntaxHelper_Ordered() As SimplyVBUnitType.IOrderedConstraint
    Set IzSyntaxHelper_Ordered = NewExp.Ordered
End Function

Private Function IzSyntaxHelper_Persistable() As SimplyVBUnitType.IBasicExpression
    Set IzSyntaxHelper_Persistable = NewExp.Persistable
End Function

Private Function IzSyntaxHelper_SameAs(ByVal Expected As Object) As SimplyVBUnitType.IBasicExpression
    Set IzSyntaxHelper_SameAs = NewExp.SameAs(Expected)
End Function

Private Function IzSyntaxHelper_SamePath(Expected As String) As SimplyVBUnitType.IPathConstraint
    Set IzSyntaxHelper_SamePath = NewExp.SamePath(Expected)
End Function

Private Function IzSyntaxHelper_SamePathOrUnder(Expected As String) As SimplyVBUnitType.IPathConstraint
    Set IzSyntaxHelper_SamePathOrUnder = NewExp.SamePathOrUnder(Expected)
End Function

Private Function IzSyntaxHelper_StringContaining(Expected As String) As SimplyVBUnitType.ITextConstraint
    Set IzSyntaxHelper_StringContaining = NewExp.StringContaining(Expected)
End Function

Private Function IzSyntaxHelper_StringEnding(Expected As String) As SimplyVBUnitType.ITextConstraint
    Set IzSyntaxHelper_StringEnding = NewExp.StringEnding(Expected)
End Function

Private Function IzSyntaxHelper_StringMatching(Pattern As String) As SimplyVBUnitType.ITextConstraint
    Set IzSyntaxHelper_StringMatching = NewExp.StringMatching(Pattern)
End Function

Private Function IzSyntaxHelper_StringStarting(Expected As String) As SimplyVBUnitType.ITextConstraint
    Set IzSyntaxHelper_StringStarting = NewExp.StringStarting(Expected)
End Function

Private Function IzSyntaxHelper_SubsetOf(Expected As Variant) As SimplyVBUnitType.IEqualityExpression
    Set IzSyntaxHelper_SubsetOf = NewExp.SubsetOf(Expected)
End Function

Private Function IzSyntaxHelper_True() As SimplyVBUnitType.IBasicExpression
    Set IzSyntaxHelper_True = NewExp.True
End Function

Private Function IzSyntaxHelper_Unique() As SimplyVBUnitType.IEqualityExpression
    Set IzSyntaxHelper_Unique = NewExp.Unique
End Function

Private Function IzSyntaxHelper_VBEmpty() As SimplyVBUnitType.IBasicExpression
    Set IzSyntaxHelper_VBEmpty = NewExp.vbEmpty
End Function

Private Function IzSyntaxHelper_VBNull() As SimplyVBUnitType.IBasicExpression
    Set IzSyntaxHelper_VBNull = NewExp.vbNull
End Function

