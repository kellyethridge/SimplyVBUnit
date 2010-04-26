VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tolerance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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
' Module: Tolerance
'
Option Explicit

Public Enum ToleranceMode
    NoneMode
    LinearMode
    PercentMode
End Enum

Private mAmount     As Variant
Private mMode       As ToleranceMode
Private mDateUnits  As String


Public Property Get IsEmpty() As Boolean
    IsEmpty = VBA.IsEmpty(mAmount)
End Property

Public Property Get Amount() As Variant
    Amount = mAmount
End Property

Public Property Get Mode() As ToleranceMode
    Mode = mMode
End Property

'Public Property Get Seconds() As Tolerance
'    Set Seconds = Sim.NewTolerance(mAmount, SecondsMode)
'End Property
'
'Public Property Get Minutes() As Tolerance
'    Set Minutes = Sim.NewTolerance(mValue, MinutesMode)
'End Property
'
'Public Property Get Hours() As Tolerance
'    Set Hours = Sim.NewTolerance(mValue, HoursMode)
'End Property
'
'Public Property Get Days() As Tolerance
'    Set Days = Sim.NewTolerance(mValue, DaysMode)
'End Property
'
'Public Property Get Weeks() As Tolerance
'    Set Weeks = Sim.NewTolerance(mValue, WeeksMode)
'End Property
'
'Public Property Get Months() As Tolerance
'    Set Months = Sim.NewTolerance(mValue, MonthsMode)
'End Property
'
'Public Property Get Quarters() As Tolerance
'    Set Quarters = Sim.NewTolerance(mValue, QuartersMode)
'End Property
'
'Public Property Get Years() As Tolerance
'    Set Years = Sim.NewTolerance(mValue, YearsMode)
'End Property
'
'Public Property Get Percent() As Tolerance
'    Set Percent = Sim.NewTolerance(mValue, PercentMode)
'End Property

Public Function IsWithin(ByRef Expected As Variant, ByRef Actual As Variant) As Boolean
    Dim ExpectedType    As VbVarType
    Dim ActualType      As VbVarType
    
    ExpectedType = VarType(Expected)
    ActualType = VarType(Actual)
    
    IsWithin = IsWithinInt32(Expected, Actual)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef Amount As Variant)
    mAmount = Amount
    mMode = LinearMode
    mDateUnits = "s"
End Sub


Private Function IsWithinInt32(ByVal Expected As Long, ByVal Actual As Long) As Boolean
    Dim Tol As Long
    Dim Result As Boolean
    
    If mMode = NoneMode Then
        Result = (Expected = Actual)
    End If
    
    IsWithinInt32 = Result
End Function
