VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TimingStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
' Module: TimingStatic
'
Option Explicit

Private Const BOOL_FALSE As Long = 0

Private mFrequency As Currency

Public Function GetTicks() As Currency
    QueryPerformanceCounter GetTicks
End Function

Public Function GetTime(ByVal Ticks As Currency) As Currency
    GetTime = (GetTicks - Ticks) * mFrequency
End Function

Private Sub Init()
    Dim Frequency As Currency
    
    If QueryPerformanceFrequency(Frequency) = BOOL_FALSE Then
        Err.Raise ErrorCode.InvalidOperation, , Resource.GetString(HighPerformanceCounterNotSupported)
    End If
    
    mFrequency = GetReciprical(Frequency)
End Sub

Private Function GetReciprical(ByVal Value As Currency) As Currency
    GetReciprical = 1000@ / Value
End Function

Private Sub Class_Initialize()
    Init
End Sub
