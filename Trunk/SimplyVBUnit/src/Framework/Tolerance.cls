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
    LinearMode
    PercentMode
    SecondsMode
    MinutesMode
    HoursMode
    DaysMode
    WeeksMode
    MonthsMode
    QuartersMode
    YearsMode
End Enum

Private mValue  As Variant
Private mMode   As ToleranceMode


Public Property Get Value() As Variant
    Value = mValue
End Property

Public Property Get Mode() As ToleranceMode
    Mode = mMode
End Property

Public Property Get Seconds() As Tolerance
    Set Seconds = Sim.NewTolerance(mValue, SecondsMode)
End Property

Public Property Get Minutes() As Tolerance
    Set Minutes = Sim.NewTolerance(mValue, MinutesMode)
End Property

Public Property Get Hours() As Tolerance
    Set Hours = Sim.NewTolerance(mValue, HoursMode)
End Property

Public Property Get Days() As Tolerance
    Set Days = Sim.NewTolerance(mValue, DaysMode)
End Property

Public Property Get Weeks() As Tolerance
    Set Weeks = Sim.NewTolerance(mValue, WeeksMode)
End Property

Public Property Get Months() As Tolerance
    Set Months = Sim.NewTolerance(mValue, MonthsMode)
End Property

Public Property Get Quarters() As Tolerance
    Set Quarters = Sim.NewTolerance(mValue, QuartersMode)
End Property

Public Property Get Years() As Tolerance
    Set Years = Sim.NewTolerance(mValue, YearsMode)
End Property

Public Property Get Percent() As Tolerance
    Set Percent = Sim.NewTolerance(mValue, PercentMode)
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef Value As Variant, ByVal Mode As ToleranceMode)
    mValue = Value
    mMode = Mode
End Sub
