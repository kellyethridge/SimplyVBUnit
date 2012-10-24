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
'The MIT License (MIT)
'Copyright (c) 2012 Kelly Ethridge
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights to
'use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of
'the Software, and to permit persons to whom the Software is furnished to do so,
'subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
'INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
'PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
'FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
'OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
'DEALINGS IN THE SOFTWARE.
'
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
