VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ToleranceTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
' Module: ToleranceTests
'
Option Explicit

Public Sub Constructor_WhenCalled_StoresValue()
    Dim Tol As Tolerance
    
    Set Tol = Sim.NewTolerance(1)
    
    AssertTolerance Tol, 1, LinearMode, ""
End Sub

Public Sub Constructor_WhenCalled_SetsDefaults()
    Dim Tol As Tolerance

    Set Tol = New Tolerance
    
    AssertTolerance Tol, Empty, NoneMode, ""
End Sub

Public Sub Seconds_WhenCalled_ReturnsToleranceInSeconds()
    Dim Tol As Tolerance
    Set Tol = Sim.NewTolerance(2)

    Set Tol = Tol.Seconds

    AssertTolerance Tol, 2, LinearMode, "s"
End Sub

Public Sub Minutes_WhenCalled_ReturnsToleranceInMinutes()
    Dim Tol As Tolerance
    Set Tol = Sim.NewTolerance(3)

    Set Tol = Tol.Minutes

    AssertTolerance Tol, 3, LinearMode, "n"
End Sub

Public Sub Hours_WhenCalled_ReturnsToleranceInHours()
    Dim Tol As Tolerance
    Set Tol = Sim.NewTolerance(1)

    Set Tol = Tol.Hours

    AssertTolerance Tol, 1, LinearMode, "h"
End Sub

Public Sub Days_WhenCalled_ReturnsToleranceInDays()
    Dim Tol As Tolerance
    Set Tol = Sim.NewTolerance(1)

    Set Tol = Tol.Days

    AssertTolerance Tol, 1, LinearMode, "d"
End Sub

Public Sub Weeks_WhenCalled_ReturnsToleranceInWeeks()
    Dim Tol As Tolerance
    Set Tol = Sim.NewTolerance(1)

    Set Tol = Tol.Weeks

    AssertTolerance Tol, 1, LinearMode, "ww"
End Sub

Public Sub Months_WhenCalled_ReturnsToleranceInMonths()
    Dim Tol As Tolerance
    Set Tol = Sim.NewTolerance(1)
    
    Set Tol = Tol.Months
    
    AssertTolerance Tol, 1, LinearMode, "m"
End Sub

Public Sub Quarters_WhenCalled_ReturnsToleranceInQuarters()
    Dim Tol As Tolerance
    Set Tol = Sim.NewTolerance(1)

    Set Tol = Tol.Quarters

    AssertTolerance Tol, 1, LinearMode, "q"
End Sub

Public Sub Years_WhenCalled_ReturnsToleranceInYears()
    Dim Tol As Tolerance
    Set Tol = Sim.NewTolerance(1)

    Set Tol = Tol.Years

    AssertTolerance Tol, 1, LinearMode, "yyyy"
End Sub

Public Sub Percent_WhenCalled_ReturnsToleranceInPercent()
    Dim Tol As Tolerance
    Set Tol = Sim.NewTolerance(1)

    Set Tol = Tol.Percent

    AssertTolerance Tol, 1, PercentMode, ""
End Sub

Public Sub TotalSeconds_WhenCalled_ReturnsToleranceInSeconds()
    Dim Tol As Tolerance
    Set Tol = Sim.NewTolerance(1)
    
    Set Tol = Tol.TotalSeconds
    
    AssertTolerance Tol, 1, LinearMode, "s"
End Sub

Public Sub TotalMinutes_WhenCalled_ReturnsToleranceScaledToSeconds()
    Dim Tol As Tolerance
    Set Tol = Sim.NewTolerance(1)
    
    Set Tol = Tol.TotalMinutes
    
    AssertTolerance Tol, 1, LinearMode, "tn"
End Sub

Public Sub TotalHours_WhenCalled_ReturnsToleranceScaledToSeconds()
    Dim Tol As Tolerance
    Set Tol = Sim.NewTolerance(1)
    
    Set Tol = Tol.TotalHours
    
    AssertTolerance Tol, 1, LinearMode, "th"
End Sub

Public Sub TotalDays_WhenCalled_ReturnsToleranceScaledToSeconds()
    Dim Tol As Tolerance
    Set Tol = Sim.NewTolerance(1)
    
    Set Tol = Tol.TotalDays
    
    AssertTolerance Tol, 1, LinearMode, "td"
End Sub



Private Sub AssertTolerance(ByVal Tol As Tolerance, ByVal ExpectedAmount As Variant, ByVal ExpectedMode As ToleranceMode, ByVal ExpectedInterval As String)
    Assert.That Tol.Amount, Iz.EqualTo(ExpectedAmount), "Wrong amount"
    Assert.That Tol.Mode, Iz.EqualTo(ExpectedMode), "Wrong mode"
    Assert.That Tol.Interval, Iz.EqualTo(ExpectedInterval), "Wrong interval"
End Sub



