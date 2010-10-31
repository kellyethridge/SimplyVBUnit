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
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Copyright 2010 Kelly Ethridge
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
Implements IEqualityContext

Private Const ASC_LOWERT                        As Long = 116
Private Const ASC_LOWERD                        As Long = 100
Private Const ASC_LOWERH                        As Long = 104
Private Const ASC_LOWERN                        As Long = 110

Private Const SECONDS_PER_SECOND                As Long = 1
Private Const SECONDS_PER_MINUTE                As Long = SECONDS_PER_SECOND * 60
Private Const SECONDS_PER_HOUR                  As Long = SECONDS_PER_MINUTE * 60
Private Const SECONDS_PER_DAY                   As Long = SECONDS_PER_HOUR * 24

Private mStrict             As Boolean
Private mIgnoreCase         As Boolean
Private mAsCollection       As Boolean
Private mTolerance          As Tolerance
Private mExternalComparer   As IEqualityComparer
Private mFailurePoints      As New Stack


Public Property Get Strict() As Boolean
    Strict = mStrict
End Property

Public Property Let Strict(ByVal RHS As Boolean)
    mStrict = RHS
End Property

Public Property Get IgnoreCase() As Boolean
    IgnoreCase = mIgnoreCase
End Property

Public Property Let IgnoreCase(ByVal RHS As Boolean)
    mIgnoreCase = RHS
End Property

Public Property Get AsCollection() As Boolean
    AsCollection = mAsCollection
End Property

Public Property Let AsCollection(ByVal RHS As Boolean)
    mAsCollection = RHS
End Property

Public Property Get Tolerance() As Tolerance
    Set Tolerance = mTolerance
End Property

Public Property Set Tolerance(ByVal RHS As Tolerance)
    Set mTolerance = RHS
End Property

Public Property Get ExternalComparer() As IEqualityComparer
    Set ExternalComparer = mExternalComparer
End Property

Public Property Set ExternalComparer(ByVal RHS As IEqualityComparer)
    Set mExternalComparer = RHS
End Property

Public Property Get FailurePoints() As Stack
    Set FailurePoints = mFailurePoints
End Property

Public Function Equals(ByRef X As Variant, ByRef Y As Variant) As Boolean
    Dim Result As Boolean
    
    If IsArray(X) And IsArray(Y) Then
        Result = EqualArrays(X, Y)
    ElseIf IsEnumerable(X) And IsEnumerable(Y) Then
        Result = EqualEnumerables(X, Y)
    ElseIf mExternalComparer Is Nothing Then
        Result = EqualValues(X, Y)
    Else
        Result = mExternalComparer.Equals(X, Y, Me)
    End If
    
    Equals = Result
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function EqualArrays(ByRef X As Variant, ByRef Y As Variant) As Boolean
    If Not mAsCollection Then
        If Not EqualArraySizes(X, Y) Then
            Exit Function
        End If
    End If
    
    EqualArrays = EqualEnumerables(X, Y)
End Function

Private Function EqualArraySizes(ByRef X As Variant, ByRef Y As Variant) As Boolean
    If EqualArrayRanks(X, Y) Then
        EqualArraySizes = EqualArrayBounds(X, Y)
    End If
End Function

Private Function EqualArrayRanks(ByRef X As Variant, ByRef Y As Variant) As Boolean
    Dim XRank    As Long
    Dim YRank      As Long
    
    XRank = GetArrayRank(X)
    YRank = GetArrayRank(Y)
    
    EqualArrayRanks = (XRank = YRank)
End Function

Private Function EqualArrayBounds(ByRef X As Variant, ByRef Y As Variant) As Boolean
    Dim i As Long
    For i = 1 To GetArrayRank(X)
        If LBound(Y, i) <> LBound(X, i) Then
            Exit Function
        End If
        If UBound(Y, i) <> UBound(X, i) Then
            Exit Function
        End If
    Next
    
    EqualArrayBounds = True
End Function

Private Function EqualEnumerables(ByRef X As Variant, ByRef Y As Variant) As Boolean
    Dim XEnumerator  As IEnumerator
    Dim YEnumerator    As IEnumerator
    
    Set XEnumerator = GetEnumerator(X)
    Set YEnumerator = GetEnumerator(Y)
    
    Dim XHasMore As Boolean
    Dim YHasMore   As Boolean
    
    Dim RetVal  As Boolean
    Dim Count   As Long
    Do
        XHasMore = XEnumerator.MoveNext
        YHasMore = YEnumerator.MoveNext
        
        If XHasMore = False And YHasMore = False Then
            RetVal = True
            Exit Do
        End If
        
        If XHasMore <> YHasMore Then
            mFailurePoints.Push Count
            RetVal = False
            Exit Do
        End If
        
        If Not Equals(XEnumerator.Current, YEnumerator.Current) Then
            mFailurePoints.Push Count
            RetVal = False
            Exit Do
        End If
        
        Count = Count + 1
    Loop
    
    EqualEnumerables = RetVal
End Function

Private Function EqualValues(ByRef X As Variant, ByRef Y As Variant) As Boolean
    Dim Result As Boolean
    
    If CanCompareValues(X, Y) Then
        Select Case VarType(X)
            Case vbString
                Result = EqualStrings(X, Y)
                
            Case vbLong, vbInteger, vbByte, vbDouble, vbSingle, vbCurrency, vbDecimal
                Result = EqualNumbers(X, Y)
                
            Case vbObject, vbDataObject
                Result = EqualObjects(X, Y)
                
            Case vbBoolean
                Result = EqualBooleans(X, Y)
                
            Case vbDate
                Result = EqualDates(X, Y)
            
            Case vbEmpty
                Result = IsEmpty(Y)
                
            Case vbNull
                Result = IsNull(Y)
                
            Case vbError
                Result = (IsMissing(X) And IsMissing(Y))
                
        End Select
    End If
    
    EqualValues = Result
End Function

Private Function CanCompareValues(ByRef X As Variant, ByRef Y As Variant) As Boolean
    If mStrict Then
        If VarType(X) <> VarType(Y) Then
            Exit Function
        End If
    End If
    
    CanCompareValues = True
End Function

Private Function EqualStrings(ByRef X As Variant, ByRef Y As Variant) As Boolean
    If VarType(Y) = vbString Then
        EqualStrings = modFunctions.EqualStrings(X, Y, mIgnoreCase)
    End If
End Function

Private Function EqualNumbers(ByRef X As Variant, ByRef Y As Variant) As Boolean
    If IsNumber(Y) Then
        EqualNumbers = modNumerics.EqualNumbers(X, Y, mTolerance)
    End If
End Function

Private Function EqualObjects(ByRef X As Variant, ByRef Y As Variant) As Boolean
    If IsObject(Y) Then
        EqualObjects = (Y Is X)
    End If
End Function

Private Function GetEnumerator(ByRef Enumerable As Variant) As IEnumerator
    If IsArray(Enumerable) Then
        Set GetEnumerator = Sim.NewArrayEnumerator(Enumerable)
    Else
        Set GetEnumerator = Sim.NewEnumVariantEnumerator(Enumerable)
    End If
End Function

Private Function EqualBooleans(ByRef X As Variant, ByRef Y As Variant) As Boolean
    If VarType(Y) = vbBoolean Then
        EqualBooleans = (Y = X)
    End If
End Function

Private Function EqualDates(ByRef X As Variant, ByRef Y As Variant) As Boolean
    If VarType(Y) = vbDate Then
        Dim Interval    As String
        Dim Tolerance   As Double
        
        Tolerance = mTolerance.Amount
        
        If mTolerance.Mode = NoneMode Or Len(mTolerance.Interval) = 0 Then
            Interval = GlobalSettings.DefaultDateToleranceInterval
        Else
            Interval = mTolerance.Interval
            
            If AscW(Interval) = ASC_LOWERT Then
                Select Case AscW(Right$(Interval, 1))
                    Case ASC_LOWERD: Tolerance = Tolerance * SECONDS_PER_DAY
                    Case ASC_LOWERH: Tolerance = Tolerance * SECONDS_PER_HOUR
                    Case ASC_LOWERN: Tolerance = Tolerance * SECONDS_PER_MINUTE
                End Select
                
                Interval = "s"
            End If
        End If
        
        Dim Difference As Double
        Difference = Abs(DateDiff(Interval, X, Y))
        
        EqualDates = (Difference <= Tolerance)
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set mTolerance = modStatics.Tolerance.EmptyTolerance
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IEqualityContext Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get IEqualityContext_AsCollection() As Boolean
    IEqualityContext_AsCollection = AsCollection
End Property

Private Property Get IEqualityContext_IgnoreCase() As Boolean
    IEqualityContext_IgnoreCase = IgnoreCase
End Property

Private Property Get IEqualityContext_Strict() As Boolean
    IEqualityContext_Strict = Strict
End Property

Private Property Get IEqualityContext_Tolerance() As SimplyVBUnitType.ITolerance
    Set IEqualityContext_Tolerance = Tolerance
End Property
