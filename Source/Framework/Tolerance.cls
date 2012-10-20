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
' Module: Tolerance
'
Option Explicit

Private mAmount     As Variant
Private mMode       As ToleranceMode
Private mInterval   As String


Public Property Get Amount() As Variant
    Amount = mAmount
End Property

Public Property Get Mode() As ToleranceMode
    Mode = mMode
End Property

Public Property Get Interval() As String
    Interval = mInterval
End Property

Public Property Get Seconds() As Tolerance
    Set Seconds = NewToleranceWithInterval("s")
End Property

Public Property Get Minutes() As Tolerance
    Set Minutes = NewToleranceWithInterval("n")
End Property

Public Property Get Hours() As Tolerance
    Set Hours = NewToleranceWithInterval("h")
End Property

Public Property Get Days() As Tolerance
    Set Days = NewToleranceWithInterval("d")
End Property

Public Property Get Weeks() As Tolerance
    Set Weeks = NewToleranceWithInterval("ww")
End Property

Public Property Get Months() As Tolerance
    Set Months = NewToleranceWithInterval("m")
End Property

Public Property Get Quarters() As Tolerance
    Set Quarters = NewToleranceWithInterval("q")
End Property

Public Property Get Years() As Tolerance
    Set Years = NewToleranceWithInterval("yyyy")
End Property

Public Property Get TotalDays() As Tolerance
    Set TotalDays = NewToleranceWithInterval("td")
End Property

Public Property Get TotalHours() As Tolerance
    Set TotalHours = NewToleranceWithInterval("th")
End Property

Public Property Get TotalMinutes() As Tolerance
    Set TotalMinutes = NewToleranceWithInterval("tn")
End Property

Public Property Get TotalSeconds() As Tolerance
    Set TotalSeconds = NewToleranceWithInterval("s")
End Property

Public Property Get Percent() As Tolerance
    Set Percent = Sim.NewTolerance(mAmount)
    Percent.Mode = PercentMode
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef Amount As Variant)
    If Not IsNumber(Amount) Then _
        Err.Raise ErrorCode.Argument, , "Tolerance must be numeric."
    
    mAmount = Amount
    mMode = LinearMode
End Sub

Friend Property Let Interval(ByRef RHS As String)
    mInterval = RHS
End Property

Friend Property Let Mode(ByVal RHS As ToleranceMode)
    mMode = RHS
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function NewToleranceWithInterval(ByRef Interval As String) As Tolerance
    Dim Result As Tolerance
    Set Result = Sim.NewTolerance(mAmount)
    Result.Interval = Interval
    
    Set NewToleranceWithInterval = Result
End Function

