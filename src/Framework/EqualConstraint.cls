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
' Module: EqualConstraint
'
Option Explicit
Implements IConstraint
Implements IResolvable
Implements IBasicExpression
Implements IEqualConstraint
Implements IToleranceUnits

Private Const MSG_STRINGSDIFFERBYINDEX          As String = "String lengths are both {0}. Strings differ at index {1}."
Private Const MSG_STRINGSDIFFERBYLENGTHANDINDEX As String = "Expected string length {0} but was {1}. Strings differ at index {2}."
Private Const DEF_CLIPPING                      As Boolean = True

Private Enum ToleranceType
    Linear
    InDays
    InHours
    InMinutes
    InSeconds
End Enum

Private mExpression     As ConstraintExpression
Private mExpected       As Variant
Private mActual         As Variant
Private mIgnoreCase     As Boolean
Private mClipping       As Boolean
Private mTolerance      As Variant
Private mToleranceType  As ToleranceType
Private mAsCollection   As Boolean
Private mComparer       As IEqualityComparer
Private mStrict         As Boolean



Public Function NoClip() As EqualConstraint
    mClipping = False
    Set NoClip = Me
End Function

''
' Tells the constraint to perform string comparisons as case-insensitive.
'
' @return Returns a reference to itself to make it easier to chain commands together.
'
Public Function IgnoreCase() As EqualConstraint
    mIgnoreCase = True
    Set IgnoreCase = Me
End Function

Public Function Within(ByRef Tolerance As Variant) As IToleranceUnits
    mTolerance = Tolerance
    Set Within = Me
End Function

Public Function AsCollection() As EqualConstraint
    mAsCollection = True
    Set AsCollection = Me
End Function

Public Function Using(ByVal Comparer As IEqualityComparer) As EqualConstraint
    Set mComparer = Comparer
    Set Using = Me
End Function

Public Function Strict() As EqualConstraint
    mStrict = True
    Set Strict = Me
End Function

''
' Compares the actual value to the expected value, returning the result.
'
' @param Actual The value to be compared to the expected value.
' @return The result of the comparison.
'
Public Function Matches(ByRef Actual As Variant) As Boolean
    Call VariantCopyInd(mActual, Actual)
    Matches = EqualVariants(mExpected, Actual)
End Function

''
' Writes a message to a text writer describing how the constraint failed.
'
' @param Writer The text write that the message will be written to.
'
Public Sub WriteMessageTo(ByVal Writer As TextMessageWriter)
    If VarType(mActual) = vbString And VarType(mExpected) = vbString Then
        DisplayStringDifferences Writer
    Else
        Call Writer.DisplayConstraintDifference(Me)
    End If
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef Expected As Variant, ByVal Expression As ConstraintExpression)
    Call VariantCopyInd(mExpected, Expected)
    Set mExpression = Expression
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Days() As EqualConstraint
    mToleranceType = InDays
    Set Days = Me
End Function

Private Function Hours() As EqualConstraint
    mToleranceType = InHours
    Set Hours = Me
End Function

Private Function Minutes() As EqualConstraint
    mToleranceType = InMinutes
    Set Minutes = Me
End Function

Private Function Seconds() As EqualConstraint
    mToleranceType = InSeconds
    Set Seconds = Me
End Function

Private Function EqualVariants(ByRef Expected As Variant, ByRef Actual As Variant) As Boolean
    Dim Result As Boolean
    
    If IsArray(Expected) And IsArray(Actual) Then
        Result = EqualArrays(Expected, Actual)
    ElseIf IsEnumerable(Expected) And IsEnumerable(Actual) Then
        Result = EqualEnumerables(Expected, Actual)
    Else
        Result = EqualValues(Expected, Actual)
    End If
    
    EqualVariants = Result
End Function

Private Function EqualArrays(ByRef Expected As Variant, ByRef Actual As Variant) As Boolean
    If Not mAsCollection Then
        If Not EqualArraySizes(Expected, Actual) Then
            Exit Function
        End If
    End If
    
    EqualArrays = EqualEnumerables(Expected, Actual)
End Function

Private Function EqualArraySizes(ByRef Expected As Variant, ByRef Actual As Variant) As Boolean
    If EqualArrayRanks(Expected, Actual) Then
        EqualArraySizes = EqualArrayBounds(Expected, Actual)
    End If
End Function

Private Function EqualArrayRanks(ByRef Expected As Variant, ByRef Actual As Variant) As Boolean
    Dim ExpectedRank    As Long
    Dim ActualRank      As Long
    
    ExpectedRank = GetArrayRank(Expected)
    ActualRank = GetArrayRank(Actual)
    
    EqualArrayRanks = (ExpectedRank = ActualRank)
End Function

Private Function EqualArrayBounds(ByRef Expected As Variant, ByRef Actual As Variant) As Boolean
    Dim i As Long
    For i = 1 To GetArrayRank(Expected)
        If LBound(Actual, i) <> LBound(Expected, i) Then
            Exit Function
        End If
        If UBound(Actual, i) <> UBound(Expected, i) Then
            Exit Function
        End If
    Next
    
    EqualArrayBounds = True
End Function

Private Function EqualEnumerables(ByRef Expected As Variant, ByRef Actual As Variant) As Boolean
    Dim ExpectedEnumerator  As IEnumerator
    Dim ActualEnumerator    As IEnumerator
    
    Set ExpectedEnumerator = GetEnumerator(Expected)
    Set ActualEnumerator = GetEnumerator(Actual)
    
    Dim ExpectedHasMore As Boolean
    Dim ActualHasMore   As Boolean
    
    Do
        ExpectedHasMore = ExpectedEnumerator.MoveNext
        ActualHasMore = ActualEnumerator.MoveNext
        
        If ExpectedHasMore = False Or ActualHasMore = False Then
            Exit Do
        End If
        
        If Not EqualVariants(ExpectedEnumerator.Current, ActualEnumerator.Current) Then
            Exit Function
        End If
    Loop
    
    EqualEnumerables = (ExpectedHasMore = False And ActualHasMore = False)
End Function

Private Function EqualValues(ByRef Expected As Variant, ByRef Actual As Variant) As Boolean
    Dim Result As Boolean
    
    If mComparer Is Nothing Then
        Result = InternalEqualValues(Expected, Actual)
    Else
        Result = mComparer.Equals(Expected, Actual)
    End If
    
    EqualValues = Result
End Function

Private Function InternalEqualValues(ByRef Expected As Variant, ByRef Actual As Variant) As Boolean
    Dim Result As Boolean
    
    If CanCompareValues(Expected, Actual) Then
        Select Case VarType(Expected)
            Case vbString
                Result = EqualStrings(Expected, Actual)
                
            Case vbLong, vbInteger, vbByte, vbDouble, vbSingle, vbCurrency, vbDecimal
                Result = EqualNumbers(Expected, Actual)
                
            Case vbObject, vbDataObject
                Result = EqualObjects(Expected, Actual)
                
            Case vbBoolean
                Result = EqualBooleans(Expected, Actual)
                
            Case vbDate
                Result = EqualDates(Expected, Actual)
            
            Case vbEmpty
                Result = IsEmpty(Actual)
                
            Case vbNull
                Result = IsNull(Actual)
                
            Case vbError
                Result = (IsMissing(Expected) And IsMissing(Actual))
                
        End Select
    End If
    
    InternalEqualValues = Result
End Function

Private Function CanCompareValues(ByRef Expected As Variant, ByRef Actual As Variant) As Boolean
    If mStrict Then
        If VarType(Expected) <> VarType(Actual) Then
            Exit Function
        End If
    End If
    
    CanCompareValues = True
End Function

Private Function EqualStrings(ByRef Expected As Variant, ByRef Actual As Variant) As Boolean
    If VarType(Actual) = vbString Then
        EqualStrings = modUtilities.EqualStrings(Expected, Actual, mIgnoreCase)
    End If
End Function

Private Function EqualNumbers(ByRef Expected As Variant, ByRef Actual As Variant) As Boolean
    If IsNumber(Actual) Then
        Dim Diff As Variant
        Diff = Abs(Expected - Actual)
        EqualNumbers = (Diff <= mTolerance)
    End If
End Function

Private Function EqualObjects(ByRef Expected As Variant, ByRef Actual As Variant) As Boolean
    If IsObject(Actual) Then
        EqualObjects = (Actual Is Expected)
    End If
End Function

Private Function GetEnumerator(ByRef Enumerable As Variant) As IEnumerator
    If IsArray(Enumerable) Then
        Set GetEnumerator = Sim.NewArrayEnumerator(Enumerable)
    Else
        Set GetEnumerator = Sim.NewEnumVariantEnumerator(Enumerable)
    End If
End Function

Private Function EqualBooleans(ByRef Expected As Variant, ByRef Actual As Variant) As Boolean
    If VarType(Actual) = vbBoolean Then
        EqualBooleans = (Actual = Expected)
    End If
End Function

Private Function EqualDates(ByRef Expected As Variant, ByRef Actual As Variant) As Boolean
    If VarType(Actual) = vbDate Then
        Dim ActualDifference As Double
        ActualDifference = Abs(DateDiff("s", Expected, Actual))
        
        Dim MaximumDifference As Double
        MaximumDifference = CDbl(mTolerance) * GetDateDifferenceMultiplier
        
        EqualDates = (ActualDifference <= MaximumDifference)
    End If
End Function

Private Function GetDateDifferenceMultiplier() As Long
    Const SECONDS_PER_DAY       As Long = 86400
    Const SECONDS_PER_HOUR      As Long = 3600
    Const SECONDS_PER_MINUTE    As Long = 60
    Const SECONDS_PER_SECOND    As Long = 1
    
    Dim Result As Long
        
    Select Case mToleranceType
        Case InDays:    Result = SECONDS_PER_DAY
        Case InHours:   Result = SECONDS_PER_HOUR
        Case InMinutes: Result = SECONDS_PER_MINUTE
        Case InSeconds: Result = SECONDS_PER_SECOND
    End Select
        
    GetDateDifferenceMultiplier = Result
End Function

Private Function IsNumber(ByRef Value As Variant) As Boolean
    Select Case VarType(Value)
        Case vbLong, vbInteger, vbByte, vbDouble, vbSingle, vbCurrency, vbDecimal
            IsNumber = True
    End Select
End Function

Private Sub DisplayStringDifferences(ByVal Writer As TextMessageWriter)
    Dim MismatchIndex As Long
    MismatchIndex = MsgUtils.FindMismatchIndex(CStr(mExpected), CStr(mActual), mIgnoreCase)
    
    If Len(mExpected) = Len(mActual) Then
        Call Writer.WriteLine(MSG_STRINGSDIFFERBYINDEX, Len(mExpected), MismatchIndex)
    Else
        Call Writer.WriteLine(MSG_STRINGSDIFFERBYLENGTHANDINDEX, Len(mExpected), Len(mActual), MismatchIndex)
    End If
    
    Call Writer.DisplayStringDifferences(CStr(mExpected), CStr(mActual), MismatchIndex, mIgnoreCase, mClipping)
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
    Call Writer.WriteActualValue(mActual)
End Sub

Private Sub IConstraint_WriteDescriptionTo(ByVal Writer As TextMessageWriter)
    Call Writer.WriteExpectedValue(mExpected)
End Sub

Private Sub IConstraint_WriteMessageTo(ByVal Writer As TextMessageWriter)
    Call WriteMessageTo(Writer)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IEqualConstraint
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IEqualConstraint_AndAlso() As SimplyVBUnitType.IConstraintExpression
    Set IEqualConstraint_AndAlso = mExpression.AndAlso
End Function

Private Function IEqualConstraint_AsCollection() As SimplyVBUnitType.IEqualConstraint
    Set IEqualConstraint_AsCollection = AsCollection
End Function

Private Function IEqualConstraint_IgnoreCase() As SimplyVBUnitType.IEqualConstraint
    Set IEqualConstraint_IgnoreCase = Me.IgnoreCase
End Function

Private Function IEqualConstraint_NoClip() As SimplyVBUnitType.IEqualConstraint
    Set IEqualConstraint_NoClip = Me.NoClip
End Function

Private Function IEqualConstraint_OrElse() As SimplyVBUnitType.IConstraintExpression
    Set IEqualConstraint_OrElse = mExpression.OrElse
End Function

Private Function IEqualConstraint_Strict() As SimplyVBUnitType.IEqualConstraint
    Set IEqualConstraint_Strict = Strict
End Function

Private Function IEqualConstraint_Using(ByVal Comparer As SimplyVBUnitType.IEqualityComparer) As SimplyVBUnitType.IEqualConstraint
    Set IEqualConstraint_Using = Using(Comparer)
End Function

Private Function IEqualConstraint_Within(ByVal Tolerance As Variant) As SimplyVBUnitType.IToleranceUnits
    Set IEqualConstraint_Within = Within(Tolerance)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IResolvable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IResolvable_Resolve() As IConstraint
    Dim Result As IConstraint
    
    If mExpression Is Nothing Then
        Set Result = Me
    Else
        Set Result = mExpression.Resolve
    End If
    
    Set IResolvable_Resolve = Result
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IToleranceUnits Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IToleranceUnits_Days() As SimplyVBUnitType.IEqualConstraint
    Set IToleranceUnits_Days = Days
End Function

Private Function IToleranceUnits_Hours() As SimplyVBUnitType.IEqualConstraint
    Set IToleranceUnits_Hours = Hours
End Function

Private Function IToleranceUnits_Minutes() As SimplyVBUnitType.IEqualConstraint
    Set IToleranceUnits_Minutes = Minutes
End Function

Private Function IToleranceUnits_Seconds() As SimplyVBUnitType.IEqualConstraint
    Set IToleranceUnits_Seconds = Seconds
End Function
