VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DateComparer"
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
' Module: DateComparer
'
Option Explicit
Implements IComparer

Private mInterval       As String
Private mDateInterval   As DateInterval


Public Property Get Interval() As DateInterval
    Interval = mDateInterval
End Property

Public Property Let Interval(ByVal RHS As DateInterval)
    mDateInterval = RHS
    
    Select Case RHS
        Case OfSeconds:     mInterval = "s"
        Case OfMinutes:     mInterval = "n"
        Case OfHours:       mInterval = "h"
        Case OfDays:        mInterval = "d"
        Case OfMonths:      mInterval = "m"
        Case OfYears:       mInterval = "yyyy"
        Case OfWeeks:       mInterval = "ww"
        Case OfQuarters:    mInterval = "q"
        Case OfDayOfYears:  mInterval = "y"
        Case Else
            Err.Raise ErrorCode.Argument, "DateComparer.Interval", "Invalid interval value."
    End Select
End Property

Public Function Compare(ByRef X As Variant, ByRef Y As Variant) As Long
    Dim Result As Long
    
    Select Case DateDiff(mInterval, X, Y)
        Case Is > 0:    Result = LESS_THAN
        Case Is < 0:    Result = GREATER_THAN
        Case Else:      Result = EQUAL_TO
    End Select
    
    Compare = Result
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IComparer Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IComparer_Compare(X As Variant, Y As Variant) As Long
    IComparer_Compare = Compare(X, Y)
End Function
