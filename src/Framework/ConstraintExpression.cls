VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConstraintExpression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Copyright 2009 Kelly Ethridge
'
' Licensed under the Apache License, Version 2.0 (the "License");
' you may not use this file except in compliance with the License.
' You may obtain a copy of the License at
'
'     http://www.apache.org/licenses/LICENSE-2.0
'
' Unless required by applicable law or agreed to in writing, software
' distributed under the License is distributed on an "AS IS" BASIS,
' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
' See the License for the specific language governing permissions and
' limitations under the License.
'
' Module: ConstraintExpression
'
Option Explicit
Implements IBasicExpression
Implements IConstraintExpression

Private mBuilder As New ConstraintBuilder


Public Function Resolve() As IConstraint
    Set Resolve = mBuilder.Resolve
    Set mBuilder = Nothing
End Function

Public Function OrElse() As IConstraintExpression
    Set OrElse = AddOperator(New OrOperator)
End Function

Public Function AndAlso() As IConstraintExpression
    Set AndAlso = AddOperator(New AndOperator)
End Function

Public Function Not_() As IConstraintExpression
    Set Not_ = AddOperator(New NotOperator)
End Function

Public Function EqualTo(ByRef Expected As Variant) As EqualConstraint
    Set EqualTo = AddConstraint(Sim.NewEqualConstraint(Expected, Me))
End Function

Public Function GreaterThan(ByRef Expected As Variant) As IBasicExpression
    Set GreaterThan = AddConstraint(Sim.NewComparisonConstraint(Expected, False, True, False, "greater than", Me))
End Function

Public Function GreaterThanOrEqualTo(ByRef Expected As Variant) As IBasicExpression
    Set GreaterThanOrEqualTo = AddConstraint(Sim.NewComparisonConstraint(Expected, False, True, True, "greater than or equal to", Me))
End Function

Public Function LessThan(ByRef Expected As Variant) As IBasicExpression
    Set LessThan = AddConstraint(Sim.NewComparisonConstraint(Expected, True, False, False, "less than", Me))
End Function

Public Function LessThanOrEqualTo(ByRef Expected As Variant) As IBasicExpression
    Set LessThanOrEqualTo = AddConstraint(Sim.NewComparisonConstraint(Expected, True, False, True, "less than or equal to", Me))
End Function

Public Function InRange(ByRef FromValue As Variant, ByRef ToValue As Variant) As RangeConstraint
    Set InRange = AddConstraint(Sim.NewRangeConstraint(FromValue, ToValue, Me))
End Function

Public Function Empty_() As IBasicExpression
    Set Empty_ = AddConstraint(Sim.NewResolvableConstraint(New EmptyConstraint, Me))
End Function

Public Function vbEmpty() As IBasicExpression
    Set vbEmpty = EqualTo(Empty)
End Function

Public Function vbNull() As IBasicExpression
    Set vbNull = EqualTo(Null)
End Function

Public Function Nothing_() As IBasicExpression
    Set Nothing_ = EqualTo(Nothing)
End Function

Public Function True_() As IBasicExpression
    Set True_ = EqualTo(True)
End Function

Public Function False_() As IBasicExpression
    Set False_ = EqualTo(False)
End Function

Public Function Missing() As IBasicExpression
    Set Missing = EqualTo(MissingVariant)
End Function

Public Function SameAs(ByVal Expected As Object) As IBasicExpression
    Dim Constraint As SameAsConstraint
    Set Constraint = Sim.NewSameAsConstraint(Expected)
    Set SameAs = AddConstraint(Sim.NewResolvableConstraint(Constraint, Me))
End Function

Public Function Unique() As UniqueItemsConstraint
    Set Unique = AddConstraint(Sim.NewUniqueItemsConstraint(Me))
End Function

Public Function EquivalentTo(ByRef Expected As Variant) As EquivalentConstraint
    Set EquivalentTo = AddConstraint(Sim.NewEquivalentConstraint(Expected, Me))
End Function

Public Function SubsetOf(ByRef Expected As Variant) As SubsetOfConstraint
    Set SubsetOf = AddConstraint(Sim.NewSubsetOfConstraint(Expected, Me))
End Function

Public Function Ordered() As IOrderedConstraint
    Set Ordered = AddConstraint(New OrderedConstraint)
End Function

Public Function StringContaining(ByRef Expected As String) As ITextConstraint
    Set StringContaining = AddConstraint(Sim.NewSubstringConstraint(Expected, Me))
End Function

Public Function StringStarting(ByRef Expected As String) As ITextConstraint
    Set StringStarting = AddConstraint(Sim.NewStartsWithConstraint(Expected, Me))
End Function

Public Function StringEnding(ByRef Expected As String) As ITextConstraint
    Set StringEnding = AddConstraint(Sim.NewEndsWithConstraint(Expected, Me))
End Function

Public Function StringMatching(ByRef Pattern As String) As ITextConstraint
    Set StringMatching = AddConstraint(Sim.NewRegexConstraint(Pattern, Me))
End Function

Public Function SamePath(ByRef Expected As String) As IPathConstraint
    Set SamePath = AddConstraint(Sim.NewSamePathConstraint(Expected, Me))
End Function

Public Function SamePathOrUnder(ByRef Expected As String) As IPathConstraint
    Set SamePathOrUnder = AddConstraint(Sim.NewSamePathOrUnderConstraint(Expected, Me))
End Function

Public Function Persistable() As IBasicExpression
    Set Persistable = AddConstraint(Sim.NewResolvableConstraint(New PersistableConstraint, Me))
End Function

Public Function All() As IConstraintExpression
    Set All = AddOperator(New AllItemsOperator)
End Function

Public Function Some() As IConstraintExpression
    Set Some = AddOperator(New SomeItemsOperator)
End Function

Public Function No() As IConstraintExpression
    Set No = Not_
End Function

Public Function None() As IConstraintExpression
    Set None = AddOperator(New NoItemsOperator)
End Function

Public Function Member(ByRef Expected As Variant) As CollectionContainsConstraint
    Set Member = AddConstraint(Sim.NewCollectionContainsConstraint(Expected, Me))
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AddConstraint(ByVal Constraint As IConstraint) As IConstraint
    mBuilder.AddConstraint Constraint
    Set AddConstraint = Constraint
End Function

Private Function AddOperator(ByVal Operator As IOperator) As IConstraintExpression
    mBuilder.AddOperator Operator
    Set AddOperator = Me
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IBasicExpression Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IBasicExpression_AndAlso() As SimplyVBUnitType.IConstraintExpression
    Set IBasicExpression_AndAlso = AndAlso
End Function

Private Function IBasicExpression_OrElse() As SimplyVBUnitType.IConstraintExpression
    Set IBasicExpression_OrElse = OrElse
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IConstraintExpression Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IConstraintExpression_All() As SimplyVBUnitType.IConstraintExpression
    Set IConstraintExpression_All = All
End Function

Private Function IConstraintExpression_Empty() As SimplyVBUnitType.IBasicExpression
    Set IConstraintExpression_Empty = Empty_
End Function

Private Function IConstraintExpression_EqualTo(Expected As Variant) As SimplyVBUnitType.IEqualityExpression
    Set IConstraintExpression_EqualTo = EqualTo(Expected)
End Function

Private Function IConstraintExpression_EquivalentTo(Expected As Variant) As SimplyVBUnitType.IEqualityExpression
    Set IConstraintExpression_EquivalentTo = EquivalentTo(Expected)
End Function

Private Function IConstraintExpression_False() As SimplyVBUnitType.IBasicExpression
    Set IConstraintExpression_False = False_
End Function

Private Function IConstraintExpression_GreaterThan(Expected As Variant) As SimplyVBUnitType.IBasicExpression
    Set IConstraintExpression_GreaterThan = GreaterThan(Expected)
End Function

Private Function IConstraintExpression_GreaterThanOrEqualTo(Expected As Variant) As SimplyVBUnitType.IBasicExpression
    Set IConstraintExpression_GreaterThanOrEqualTo = GreaterThanOrEqualTo(Expected)
End Function

Private Function IConstraintExpression_InRange(FromValue As Variant, ToValue As Variant) As IComparisonConstraint
    Set IConstraintExpression_InRange = InRange(FromValue, ToValue)
End Function

Private Function IConstraintExpression_LessThan(Expected As Variant) As SimplyVBUnitType.IBasicExpression
    Set IConstraintExpression_LessThan = LessThan(Expected)
End Function

Private Function IConstraintExpression_LessThanOrEqualTo(Expected As Variant) As SimplyVBUnitType.IBasicExpression
    Set IConstraintExpression_LessThanOrEqualTo = LessThanOrEqualTo(Expected)
End Function

Private Function IConstraintExpression_Member(Expected As Variant) As SimplyVBUnitType.ICollectionContainsConstraint
    Set IConstraintExpression_Member = Member(Expected)
End Function

Private Function IConstraintExpression_Missing() As SimplyVBUnitType.IBasicExpression
    Set IConstraintExpression_Missing = Missing
End Function

Private Function IConstraintExpression_No() As SimplyVBUnitType.IConstraintExpression
    Set IConstraintExpression_No = No
End Function

Private Function IConstraintExpression_None() As SimplyVBUnitType.IConstraintExpression
    Set IConstraintExpression_None = None
End Function

Private Function IConstraintExpression_Not() As SimplyVBUnitType.IConstraintExpression
    Set IConstraintExpression_Not = Not_
End Function

Private Function IConstraintExpression_Nothing() As SimplyVBUnitType.IBasicExpression
    Set IConstraintExpression_Nothing = Nothing_
End Function

Private Function IConstraintExpression_Ordered() As SimplyVBUnitType.IOrderedConstraint
    Set IConstraintExpression_Ordered = Ordered
End Function

Private Function IConstraintExpression_Persistable() As SimplyVBUnitType.IBasicExpression
    Set IConstraintExpression_Persistable = Persistable
End Function

Private Function IConstraintExpression_SameAs(ByVal Expected As Object) As SimplyVBUnitType.IBasicExpression
    Set IConstraintExpression_SameAs = SameAs(Expected)
End Function

Private Function IConstraintExpression_SamePath(Expected As String) As SimplyVBUnitType.IPathConstraint
    Set IConstraintExpression_SamePath = SamePath(Expected)
End Function

Private Function IConstraintExpression_SamePathOrUnder(Expected As String) As SimplyVBUnitType.IPathConstraint
    Set IConstraintExpression_SamePathOrUnder = SamePathOrUnder(Expected)
End Function

Private Function IConstraintExpression_Some() As SimplyVBUnitType.IConstraintExpression
    Set IConstraintExpression_Some = Some
End Function

Private Function IConstraintExpression_StringContaining(Expected As String) As SimplyVBUnitType.ITextConstraint
    Set IConstraintExpression_StringContaining = StringContaining(Expected)
End Function

Private Function IConstraintExpression_StringEnding(Expected As String) As SimplyVBUnitType.ITextConstraint
    Set IConstraintExpression_StringEnding = StringEnding(Expected)
End Function

Private Function IConstraintExpression_StringMatching(Pattern As String) As SimplyVBUnitType.ITextConstraint
    Set IConstraintExpression_StringMatching = StringMatching(Pattern)
End Function

Private Function IConstraintExpression_StringStarting(Expected As String) As SimplyVBUnitType.ITextConstraint
    Set IConstraintExpression_StringStarting = StringStarting(Expected)
End Function

Private Function IConstraintExpression_SubsetOf(Expected As Variant) As SimplyVBUnitType.IEqualityExpression
    Set IConstraintExpression_SubsetOf = SubsetOf(Expected)
End Function

Private Function IConstraintExpression_True() As SimplyVBUnitType.IBasicExpression
    Set IConstraintExpression_True = True_
End Function

Private Function IConstraintExpression_Unique() As SimplyVBUnitType.IEqualityExpression
    Set IConstraintExpression_Unique = Unique
End Function

Private Function IConstraintExpression_VBEmpty() As SimplyVBUnitType.IBasicExpression
    Set IConstraintExpression_VBEmpty = vbEmpty
End Function

Private Function IConstraintExpression_VBNull() As SimplyVBUnitType.IBasicExpression
    Set IConstraintExpression_VBNull = vbNull
End Function
