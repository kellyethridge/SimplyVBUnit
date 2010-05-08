Attribute VB_Name = "modNumerics"
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
' Module: modNumerics
'
Option Explicit


Public Function IsNumber(ByRef Value As Variant) As Boolean
    IsNumber = (IsFloatingPointNumber(Value) Or IsFixedPointNumber(Value))
End Function

Public Function IsFloatingPointNumber(ByRef Value As Variant) As Boolean
    Select Case VarType(Value)
        Case vbDouble, vbSingle
            IsFloatingPointNumber = True
    End Select
End Function

Public Function IsFixedPointNumber(ByRef Value As Variant) As Boolean
    Select Case VarType(Value)
        Case vbLong, vbInteger, vbByte, vbCurrency, vbDecimal
            IsFixedPointNumber = True
    End Select
End Function


Public Function EqualNumbers(ByRef Expected As Variant, ByRef Actual As Variant, ByVal Tolerance As Tolerance) As Boolean
    Dim ExpectedType    As VbVarType
    Dim ActualType      As VbVarType
    
    ExpectedType = VarType(Expected)
    ActualType = VarType(Actual)
    
    Dim Result As Boolean
    If ExpectedType = vbDouble Or ActualType = vbDouble Then
        Result = EqualDoubles(Expected, Actual, Tolerance)
    ElseIf ExpectedType = vbSingle Or ActualType = vbSingle Then
        Result = EqualDoubles(Expected, Actual, Tolerance)
    ElseIf ExpectedType = vbDecimal Or ActualType = vbDecimal Then
        Result = EqualDecimals(CDec(Expected), CDec(Actual), Tolerance)
    ElseIf ExpectedType = vbCurrency Or ActualType = vbCurrency Then
        Result = EqualCurrencies(Expected, Actual, Tolerance)
    Else
        Result = EqualLongs(Expected, Actual, Tolerance)
    End If
    
    EqualNumbers = Result
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function EqualDoubles(ByVal Expected As Double, ByVal Actual As Double, ByVal Tolerance As Tolerance) As Boolean
    If Tolerance.Mode = NoneMode Then
        If GlobalSettings.DefaultFloatingPointTolerance > 0# Then
            Set Tolerance = Sim.NewTolerance(GlobalSettings.DefaultFloatingPointTolerance)
        End If
    End If

    Dim Result  As Boolean
    Dim Tol     As Double
    
    Select Case Tolerance.Mode
        Case NoneMode
            Result = (Expected = Actual)
            
        Case LinearMode
            Tol = CDbl(Tolerance.Amount)
            
            Result = (Abs(Expected - Actual) <= Tol)
            
        Case PercentMode
            If Expected = 0# Then
                Result = (Expected = Actual)
            Else
                Dim RelativeDifference As Double
                RelativeDifference = (Abs(Expected - Actual) / Expected)
                
                Tol = (CDbl(Tolerance.Amount) / 100#)
                
                Result = (RelativeDifference <= Tol)
            End If
    End Select
    
    EqualDoubles = Result
End Function

Private Function EqualDecimals(ByRef Expected As Variant, ByRef Actual As Variant, ByVal Tolerance As Tolerance) As Boolean
    Dim Result  As Boolean
    Dim Tol     As Double
    
    Select Case Tolerance.Mode
        Case NoneMode
            Result = (Expected = Actual)
            
        Case LinearMode
            Tol = CDbl(Tolerance.Amount)
            
            Result = (Abs(Expected - Actual) <= Tol)
            
        Case PercentMode
            If Expected = CDec(0) Then
                Result = (Expected = Actual)
            Else
                Dim RelativeDifference As Double
                RelativeDifference = (Abs(Expected - Actual) / Expected)
                
                Tol = (CDbl(Tolerance.Amount) / 100#)
                
                Result = (RelativeDifference <= Tol)
            End If
    End Select
    
    EqualDecimals = Result
End Function

Private Function EqualCurrencies(ByVal Expected As Currency, ByVal Actual As Currency, ByVal Tolerance As Tolerance) As Boolean
    Dim Result  As Boolean
    Dim Tol     As Double
    
    Select Case Tolerance.Mode
        Case NoneMode
            Result = (Expected = Actual)
            
        Case LinearMode
            Tol = CDbl(Tolerance.Amount)
            
            Result = (Abs(Expected - Actual) <= Tol)
            
        Case PercentMode
            If Expected = 0@ Then
                Result = (Expected = Actual)
            Else
                Dim RelativeDifference As Double
                RelativeDifference = (Abs(Expected - Actual) / Expected)
                
                Tol = (CDbl(Tolerance.Amount) / 100#)
                
                Result = (RelativeDifference <= Tol)
            End If
    End Select
    
    EqualCurrencies = Result
End Function

Private Function EqualLongs(ByVal Expected As Long, ByVal Actual As Long, ByVal Tolerance As Tolerance) As Boolean
    Dim Result As Boolean
    
    Select Case Tolerance.Mode
        Case NoneMode
            Result = (Expected = Actual)
            
        Case LinearMode
            Dim TolLong As Long
            TolLong = CLng(Tolerance.Amount)
            
            Result = (Abs(Expected - Actual) <= TolLong)
            
        Case PercentMode
            If Expected = 0 Then
                Result = (Expected = Actual)
            Else
                Dim RelativeDifference As Double
                RelativeDifference = (Abs(Expected - Actual) / Expected)
                
                Dim TolDbl As Double
                TolDbl = (CDbl(Tolerance.Amount) / 100#)
                
                Result = (RelativeDifference <= TolDbl)
            End If
                
    End Select
    
    EqualLongs = Result
End Function
