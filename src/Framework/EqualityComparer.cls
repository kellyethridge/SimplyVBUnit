VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EqualityComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
' Module: EqualityComparer
'
Option Explicit

Public Enum ToleranceType
    Linear
    InDays
    InHours
    InMinutes
    InSeconds
End Enum

Private mCompareMethod  As VbCompareMethod
Private mTolerance      As Variant
Private mToleranceType  As ToleranceType
Private mAsCollection   As Boolean
Private mComparer       As IEqualityComparer
Private mStrict         As Boolean


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub IgnoreCase()
    mCompareMethod = vbTextCompare
End Sub

Friend Sub Within(ByRef Tolerance As Variant)
    mTolerance = Tolerance
End Sub

Friend Property Let ToleranceType(ByVal RHS As ToleranceType)
    mToleranceType = RHS
End Property

Friend Property Get ToleranceType() As ToleranceType
    ToleranceType = mToleranceType
End Property

Friend Sub AsCollection()
    mAsCollection = True
End Sub

Friend Sub Using(ByVal Comparer As IEqualityComparer)
    Set mComparer = Comparer
End Sub

Friend Sub Strict()
    mStrict = True
End Sub

Friend Function EqualVariants(ByRef Expected As Variant, ByRef Actual As Variant) As Boolean
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


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
        EqualStrings = (StrComp(Expected, Actual, mCompareMethod) = 0)
    End If
End Function

Private Function EqualNumbers(ByRef Expected As Variant, ByRef Actual As Variant) As Boolean
    If IsNumber(VarType(Actual)) Then
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

Private Function IsNumber(ByVal VarType As VbVarType) As Boolean
    Select Case VarType
        Case vbLong, vbInteger, vbByte, vbDouble, vbSingle, vbCurrency, vbDecimal
            IsNumber = True
    End Select
End Function


Private Sub Class_Initialize()
    mCompareMethod = vbBinaryCompare
    mTolerance = 0
End Sub

