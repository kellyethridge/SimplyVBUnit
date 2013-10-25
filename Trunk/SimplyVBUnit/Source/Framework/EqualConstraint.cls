VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EqualConstraint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2012 Kelly Ethridge
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights to
'use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of
'the Software, and to permit persons to whom the Software is furnished to do so,
'subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
'INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
'PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
'FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
'OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
'DEALINGS IN THE SOFTWARE.
'
'
' Module: EqualConstraint
'
Option Explicit
Implements IConstraint
Implements IResolvable
Implements IBasicExpression
Implements IEqualityExpression
Implements IToleranceUnits

Private Const MSG_STRINGSDIFFERBYINDEX          As String = "  String lengths are both {0}. Strings differ at index {1}."
Private Const MSG_STRINGSDIFFERBYLENGTHANDINDEX As String = "  Expected string length {0} but was {1}. Strings differ at index {2}."
Private Const MSG_COLLECTIONTYPE_1              As String = "  Expected and actual are both {0}"
Private Const MSG_COLLECTIONTYPE_2              As String = "  Expected is {0}, actual is {1}"
Private Const MSG_VALUESDIFFER_1                As String = "  Values differ at index [{0}]"
Private Const MSG_VALUESDIFFER_2                As String = "  Values differ at expected index [{0}], actual index [{1}]"

Private Const DEF_CLIPPING                      As Boolean = True

Private mExpression     As ConstraintExpression
Private mExpected       As Variant
Private mActual         As Variant
Private mClipping       As Boolean
Private mComparer       As EqualityComparer


Public Function NoClip() As EqualConstraint
    mClipping = False
    Set NoClip = Me
End Function

Public Function IgnoreCase() As EqualConstraint
    mComparer.IgnoreCase = True
    Set IgnoreCase = Me
End Function

Public Function Within(ByRef Amount As Variant) As IToleranceUnits
    Set Within = SetTolerance(Sim.NewTolerance(Amount))
End Function

Public Function AsCollection() As EqualConstraint
    mComparer.AsCollection = True
    Set AsCollection = Me
End Function

Public Function Using(ByVal Comparer As IEqualityComparer) As EqualConstraint
    Set mComparer.ExternalComparer = Comparer
    Set Using = Me
End Function

Public Function Strict() As EqualConstraint
    mComparer.Strict = True
    Set Strict = Me
End Function

Public Function Matches(ByRef Actual As Variant) As Boolean
    Dim Result As Boolean
    
    If mComparer.Equals(mExpected, Actual) Then
        Result = True
    Else
        VariantCopyInd mActual, Actual
    End If
    
    Matches = Result
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef Expected As Variant, ByVal Expression As ConstraintExpression)
    VariantCopyInd mExpected, Expected
    Set mExpression = Expression
    Set mComparer = New EqualityComparer
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Years() As EqualConstraint
    Set Years = SetTolerance(mComparer.Tolerance.Years)
End Function

Private Function Quarters() As EqualConstraint
    Set Quarters = SetTolerance(mComparer.Tolerance.Quarters)
End Function

Private Function Months() As EqualConstraint
    Set Months = SetTolerance(mComparer.Tolerance.Months)
End Function

Private Function Weeks() As EqualConstraint
    Set Weeks = SetTolerance(mComparer.Tolerance.Weeks)
End Function

Private Function Days() As EqualConstraint
    Set Days = SetTolerance(mComparer.Tolerance.Days)
End Function

Private Function Hours() As EqualConstraint
    Set Hours = SetTolerance(mComparer.Tolerance.Hours)
End Function

Private Function Minutes() As EqualConstraint
    Set Minutes = SetTolerance(mComparer.Tolerance.Minutes)
End Function

Private Function Seconds() As EqualConstraint
    Set Seconds = SetTolerance(mComparer.Tolerance.Seconds)
End Function

Private Function TotalDays() As EqualConstraint
    Set TotalDays = SetTolerance(mComparer.Tolerance.TotalDays)
End Function

Private Function TotalHours() As EqualConstraint
    Set TotalHours = SetTolerance(mComparer.Tolerance.TotalHours)
End Function

Private Function TotalMinutes() As EqualConstraint
    Set TotalMinutes = SetTolerance(mComparer.Tolerance.TotalMinutes)
End Function

Private Function TotalSeconds() As EqualConstraint
    Set TotalSeconds = SetTolerance(mComparer.Tolerance.TotalSeconds)
End Function

Private Function Percent() As EqualConstraint
    Set Percent = SetTolerance(mComparer.Tolerance.Percent)
End Function

Private Function SetTolerance(ByVal Tolerance As Tolerance) As EqualConstraint
    Set mComparer.Tolerance = Tolerance
    Set SetTolerance = Me
End Function

Private Sub DisplayStringDifferences(ByVal Writer As TextMessageWriter)
    Dim MismatchIndex As Long
    MismatchIndex = MsgUtils.FindMismatchIndex(CStr(mExpected), CStr(mActual), mComparer.IgnoreCase)
    
    If Len(mExpected) = Len(mActual) Then
        Writer.WriteLine MSG_STRINGSDIFFERBYINDEX, Len(mExpected), MismatchIndex
    Else
        Writer.WriteLine MSG_STRINGSDIFFERBYLENGTHANDINDEX, Len(mExpected), Len(mActual), MismatchIndex
    End If
    
    Writer.DisplayStringDifferences CStr(mExpected), CStr(mActual), MismatchIndex, mComparer.IgnoreCase, mClipping
End Sub

Private Function GetValue(ByRef Source As Variant, ByVal Index As Long) As Variant
    If IsArray(Source) Then
        VariantCopyInd GetValue, GetArrayElement(Source, Index)
    ElseIf TypeOf Source Is Collection Then
        VariantCopyInd GetValue, Source(Index + 1)
    End If
End Function

Private Sub DisplayCollectionDifferences(ByRef Expected As Variant, ByRef Actual As Variant, ByVal Writer As TextMessageWriter, Optional ByVal Indent As Long)
    DisplayCollectionTypesAndSizes Expected, Actual, Writer, Indent

    If mComparer.FailurePoints.Count > 0 Then
        Dim Index As Long
        Index = mComparer.FailurePoints.Pop

        DisplayCollectionFailurePoint Expected, Actual, Writer, Index, Indent

        If mComparer.FailurePoints.Count > 0 Then
            DisplayCollectionDifferences GetValue(Expected, Index), GetValue(Actual, Index), Writer, Indent + 2
            Exit Sub
        End If
        
        Dim ExpectedCount As Long
        Dim ActualCount As Long
        Dim Indents As String
        
        If TryGetCount(Expected, ExpectedCount) And TryGetCount(Actual, ActualCount) Then
            If ExpectedCount < ActualCount Then
                ' too many
                Writer.WriteText "  Extra: "
                Writer.WriteCollectionElements Actual, Index, 3
                Writer.WriteLine
    
            ElseIf ExpectedCount > ActualCount Then
                ' too few
                Writer.WriteText "  Missing: "
                Writer.WriteCollectionElements Expected, Index, 3
                Writer.WriteLine
            Else
                Writer.DisplayDifferences GetValue(Expected, Index), GetValue(Actual, Index)
            End If
        Else
            Writer.DisplayDifferences GetValue(Expected, Index), GetValue(Actual, Index)
        End If
    End If
End Sub

Private Sub DisplayCollectionTypesAndSizes(ByRef Expected As Variant, ByRef Actual As Variant, ByVal Writer As TextMessageWriter, Optional ByVal Indent As Long)
    Dim sExpected As String
    sExpected = MsgUtils.GetTypeRepresentation(Expected)
    AppendElementCountInfo Expected, sExpected
        
    Dim sActual As String
    sActual = MsgUtils.GetTypeRepresentation(Actual)
    AppendElementCountInfo Actual, sActual
    
    Writer.WriteText Space$(Indent)
    If sExpected = sActual Then
        Writer.WriteLine MSG_COLLECTIONTYPE_1, sExpected
    Else
        Writer.WriteLine MSG_COLLECTIONTYPE_2, sExpected, sActual
    End If
End Sub

Private Sub AppendElementCountInfo(ByRef Value As Variant, ByRef s As String)
    Dim Count As Long
    
    If TryGetCount(Value, Count) Then
        If Count = 1 Then
            s = s & " with " & Count & " element"
        Else
            s = s & " with " & Count & " elements"
        End If
    End If
End Sub

Private Sub DisplayCollectionFailurePoint(ByRef Expected As Variant, ByRef Actual As Variant, ByVal Writer As TextMessageWriter, ByVal Index As Long, ByVal Indent As Long)
    Writer.WriteText Space$(Indent)
    
    Dim ExpectedIndices As String
    ExpectedIndices = MsgUtils.GetArrayIndicesAsString(MsgUtils.GetArrayIndicesFromCollectionIndex(Expected, Index))
    
    Dim ActualIndices As String
    ActualIndices = MsgUtils.GetArrayIndicesAsString(MsgUtils.GetArrayIndicesFromCollectionIndex(Actual, Index))
    
    If ExpectedIndices = ActualIndices Then
        Writer.WriteLine MSG_VALUESDIFFER_1, ExpectedIndices
    Else
        Writer.WriteLine MSG_VALUESDIFFER_2, ExpectedIndices, ActualIndices
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mClipping = DEF_CLIPPING
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IBasicExpression Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IBasicExpression_AndAlso() As SimplyVBUnitType.IConstraintExpression
    Set IBasicExpression_AndAlso = mExpression.AndAlso
End Function

Private Function IBasicExpression_OrElse() As SimplyVBUnitType.IConstraintExpression
    Set IBasicExpression_OrElse = mExpression.OrElse
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IConstraint Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IConstraint_Matches(Actual As Variant) As Boolean
    IConstraint_Matches = Matches(Actual)
End Function

Private Sub IConstraint_WriteActualValueTo(ByVal Writer As TextMessageWriter)
    Writer.WriteActualValue mActual
End Sub

Private Sub IConstraint_WriteDescriptionTo(ByVal Writer As TextMessageWriter)
    Writer.WriteExpectedValue mExpected
End Sub

Private Sub IConstraint_WriteMessageTo(ByVal Writer As TextMessageWriter)
    If IsEnumerable(mExpected) And IsEnumerable(mActual) Then
        DisplayCollectionDifferences mExpected, mActual, Writer, 0
    ElseIf VarType(mActual) = vbString And VarType(mExpected) = vbString Then
        DisplayStringDifferences Writer
    Else
        Writer.DisplayConstraintDifferences Me
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IEqualityExpression Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IEqualityExpression_AndAlso() As SimplyVBUnitType.IConstraintExpression
    Set IEqualityExpression_AndAlso = mExpression.AndAlso
End Function

Private Function IEqualityExpression_AsCollection() As SimplyVBUnitType.IEqualityExpression
    Set IEqualityExpression_AsCollection = AsCollection
End Function

Private Function IEqualityExpression_IgnoreCase() As SimplyVBUnitType.IEqualityExpression
    Set IEqualityExpression_IgnoreCase = IgnoreCase
End Function

Private Function IEqualityExpression_NoClip() As SimplyVBUnitType.IEqualityExpression
    Set IEqualityExpression_NoClip = NoClip
End Function

Private Function IEqualityExpression_OrElse() As SimplyVBUnitType.IConstraintExpression
    Set IEqualityExpression_OrElse = mExpression.OrElse
End Function

Private Function IEqualityExpression_Strict() As SimplyVBUnitType.IEqualityExpression
    Set IEqualityExpression_Strict = Strict
End Function

Private Function IEqualityExpression_Using(ByVal Comparer As SimplyVBUnitType.IEqualityComparer) As SimplyVBUnitType.IEqualityExpression
    Set IEqualityExpression_Using = Using(Comparer)
End Function

Private Function IEqualityExpression_Within(ByVal Tolerance As Variant) As SimplyVBUnitType.IToleranceUnits
    Set IEqualityExpression_Within = Within(Tolerance)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IResolvable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IResolvable_Resolve() As IConstraint
    Set IResolvable_Resolve = Resolve(Me, mExpression)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IToleranceUnits Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IToleranceUnits_Days() As SimplyVBUnitType.IEqualityExpression
    Set IToleranceUnits_Days = Days
End Function

Private Function IToleranceUnits_Hours() As SimplyVBUnitType.IEqualityExpression
    Set IToleranceUnits_Hours = Hours
End Function

Private Function IToleranceUnits_Minutes() As SimplyVBUnitType.IEqualityExpression
    Set IToleranceUnits_Minutes = Minutes
End Function

Private Function IToleranceUnits_Months() As SimplyVBUnitType.IEqualityExpression
    Set IToleranceUnits_Months = Months
End Function

Private Function IToleranceUnits_Percent() As SimplyVBUnitType.IEqualityExpression
    Set IToleranceUnits_Percent = Percent
End Function

Private Function IToleranceUnits_Quarters() As SimplyVBUnitType.IEqualityExpression
    Set IToleranceUnits_Quarters = Quarters
End Function

Private Function IToleranceUnits_Seconds() As SimplyVBUnitType.IEqualityExpression
    Set IToleranceUnits_Seconds = Seconds
End Function

Private Function IToleranceUnits_TotalDays() As SimplyVBUnitType.IEqualityExpression
    Set IToleranceUnits_TotalDays = TotalDays
End Function

Private Function IToleranceUnits_TotalHours() As SimplyVBUnitType.IEqualityExpression
    Set IToleranceUnits_TotalHours = TotalHours
End Function

Private Function IToleranceUnits_TotalMinutes() As SimplyVBUnitType.IEqualityExpression
    Set IToleranceUnits_TotalMinutes = TotalMinutes
End Function

Private Function IToleranceUnits_TotalSeconds() As SimplyVBUnitType.IEqualityExpression
    Set IToleranceUnits_TotalSeconds = TotalSeconds
End Function

Private Function IToleranceUnits_Weeks() As SimplyVBUnitType.IEqualityExpression
    Set IToleranceUnits_Weeks = Weeks
End Function

Private Function IToleranceUnits_Years() As SimplyVBUnitType.IEqualityExpression
    Set IToleranceUnits_Years = Years
End Function
