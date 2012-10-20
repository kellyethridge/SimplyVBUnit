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
    Select Case VarType(Value)
        Case vbLong, vbInteger, vbByte, vbDouble, vbSingle, vbDecimal, vbCurrency
            IsNumber = True
    End Select
End Function

Public Function IsFloatingPointNumber(ByRef Value As Variant) As Boolean
    Select Case VarType(Value)
        Case vbDouble, vbSingle
            IsFloatingPointNumber = True
    End Select
End Function

Public Function EqualNumbers(ByRef X As Variant, ByRef Y As Variant, ByVal Tolerance As Tolerance) As Boolean
    Dim XType   As VbVarType
    Dim YType   As VbVarType
    Dim Result  As Boolean
    
    XType = VarType(X)
    YType = VarType(Y)
    
    If XType = vbDouble Or YType = vbDouble Then
        Result = EqualDoubles(X, Y, Tolerance)
    ElseIf XType = vbSingle Or YType = vbSingle Then
        Result = EqualDoubles(X, Y, Tolerance)
    ElseIf XType = vbDecimal Or YType = vbDecimal Then
        Result = EqualDecimals(CDec(X), CDec(Y), Tolerance)
    ElseIf XType = vbCurrency Or YType = vbCurrency Then
        Result = EqualCurrencies(X, Y, Tolerance)
    Else
        Result = EqualLongs(X, Y, Tolerance)
    End If
        
    EqualNumbers = Result
End Function

Public Function CompareNumbers(ByRef X As Variant, ByRef Y As Variant) As Long
    Dim XType   As VbVarType
    Dim YType   As VbVarType
    Dim Result  As Long
    
    XType = VarType(X)
    YType = VarType(Y)
    
    If XType = vbDouble Or YType = vbDouble Then
        Result = CompareDoubles(X, Y)
    ElseIf XType = vbSingle Or YType = vbSingle Then
        Result = CompareDoubles(X, Y)
    ElseIf XType = vbDecimal Or YType = vbDecimal Then
        Result = CompareDecimals(X, Y)
    ElseIf XType = vbCurrency Or YType = vbCurrency Then
        Result = CompareCurrencies(X, Y)
    Else
        Result = CompareLongs(X, Y)
    End If
    
    CompareNumbers = Result
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function EqualDoubles(ByVal X As Double, ByVal Y As Double, ByVal Tolerance As Tolerance) As Boolean
    If Tolerance.Mode = NoneMode Then
        If GlobalSettings.DefaultFloatingPointTolerance > 0# Then
            Set Tolerance = Sim.NewTolerance(GlobalSettings.DefaultFloatingPointTolerance)
        End If
    End If

    Dim Result  As Boolean
    Dim Tol     As Double
    
    Select Case Tolerance.Mode
        Case NoneMode
            Result = (X = Y)
            
        Case LinearMode
            Tol = CDbl(Tolerance.Amount)
            
            Result = (Abs(X - Y) <= Tol)
            
        Case PercentMode
            If X = 0# Then
                Result = (X = Y)
            Else
                Dim RelativeDifference As Double
                RelativeDifference = (Abs(X - Y) / X)
                
                Tol = (CDbl(Tolerance.Amount) / 100#)
                
                Result = (RelativeDifference <= Tol)
            End If
    End Select
    
    EqualDoubles = Result
End Function

Private Function EqualDecimals(ByRef X As Variant, ByRef Y As Variant, ByVal Tolerance As Tolerance) As Boolean
    Dim Result  As Boolean
    Dim Tol     As Double
    
    Select Case Tolerance.Mode
        Case NoneMode
            Result = (X = Y)
            
        Case LinearMode
            Tol = CDbl(Tolerance.Amount)
            
            Result = (Abs(X - Y) <= Tol)
            
        Case PercentMode
            If X = CDec(0) Then
                Result = (X = Y)
            Else
                Dim RelativeDifference As Double
                RelativeDifference = (Abs(X - Y) / X)
                
                Tol = (CDbl(Tolerance.Amount) / 100#)
                
                Result = (RelativeDifference <= Tol)
            End If
    End Select
    
    EqualDecimals = Result
End Function

Private Function EqualCurrencies(ByVal X As Currency, ByVal Y As Currency, ByVal Tolerance As Tolerance) As Boolean
    Dim Result  As Boolean
    Dim Tol     As Double
    
    Select Case Tolerance.Mode
        Case NoneMode
            Result = (X = Y)
            
        Case LinearMode
            Tol = CDbl(Tolerance.Amount)
            
            Result = (Abs(X - Y) <= Tol)
            
        Case PercentMode
            If X = 0@ Then
                Result = (X = Y)
            Else
                Dim RelativeDifference As Double
                RelativeDifference = (Abs(X - Y) / X)
                
                Tol = (CDbl(Tolerance.Amount) / 100#)
                
                Result = (RelativeDifference <= Tol)
            End If
    End Select
    
    EqualCurrencies = Result
End Function

Private Function EqualLongs(ByVal X As Long, ByVal Y As Long, ByVal Tolerance As Tolerance) As Boolean
    Dim Result As Boolean
    
    Select Case Tolerance.Mode
        Case NoneMode
            Result = (X = Y)
            
        Case LinearMode
            Dim TolLong As Long
            TolLong = CLng(Tolerance.Amount)
            
            Result = (Abs(X - Y) <= TolLong)
            
        Case PercentMode
            If X = 0 Then
                Result = (X = Y)
            Else
                Dim RelativeDifference As Double
                RelativeDifference = (Abs(X - Y) / X)
                
                Dim TolDbl As Double
                TolDbl = (CDbl(Tolerance.Amount) / 100#)
                
                Result = (RelativeDifference <= TolDbl)
            End If
                
    End Select
    
    EqualLongs = Result
End Function

Private Function CompareDoubles(ByVal X As Double, ByVal Y As Double) As Long
    If X < Y Then
        CompareDoubles = LESS_THAN
    ElseIf X > Y Then
        CompareDoubles = GREATER_THAN
    End If
End Function

Private Function CompareDecimals(ByRef X As Variant, ByRef Y As Variant) As Long
    If X < Y Then
        CompareDecimals = LESS_THAN
    ElseIf X > Y Then
        CompareDecimals = GREATER_THAN
    End If
End Function

Private Function CompareCurrencies(ByVal X As Currency, ByVal Y As Currency) As Long
    If X < Y Then
        CompareCurrencies = LESS_THAN
    ElseIf X > Y Then
        CompareCurrencies = GREATER_THAN
    End If
End Function

Private Function CompareLongs(ByVal X As Long, ByVal Y As Long) As Long
    If X < Y Then
        CompareLongs = LESS_THAN
    ElseIf X > Y Then
        CompareLongs = GREATER_THAN
    End If
End Function
