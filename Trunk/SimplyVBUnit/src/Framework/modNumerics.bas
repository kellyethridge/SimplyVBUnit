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

Public Function EqualNumbers(ByRef Expected As Variant, ByRef Actual As Variant, ByVal Tolerance As Tolerance) As Boolean
    Dim ExpectedType    As VbVarType
    Dim ActualType      As VbVarType
    
    ExpectedType = VarType(Expected)
    ActualType = VarType(Actual)
    
    Dim Result As Boolean
    If ExpectedType = vbDouble And ActualType = vbDouble Then
        Result = EqualDoubles(Expected, Actual, Tolerance)
    
    End If
    
    EqualNumbers = Result
End Function

Private Function EqualDoubles(ByVal Expected As Double, ByVal Actual As Double, ByVal Tolerance As Tolerance) As Boolean
    Dim EffectiveTolerance As Double
    
    Dim EffectiveDifference As Double
    EffectiveDifference = Abs(Actual - Expected)
    
End Function

