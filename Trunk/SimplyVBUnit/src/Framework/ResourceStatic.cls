VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ResourceStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
' Module: ResourceStatic
'
Option Explicit



Public Enum StringResourceID
    HighPerformanceCounterNotSupported = 101
    Arg_CannotBeNothing = 102
    Arg_CannotBeEmpty = 103
    Arg_CannotBeNegative = 104
    Arg_IndexOutOfRange = 105
    Arg_EnumerationNonSupported = 106
    Msg_FixtureSetupFailed = 122
    
End Enum


Friend Function GetString(ByVal ID As StringResourceID, ParamArray Args() As Variant) As String
    Dim Result As String
    Result = LoadResString(ID)
    
    If UBound(Args) >= 0 Then
        Dim LocalArgs() As Variant
        LocalArgs = Args
        
        Dim sb As New StringBuilder
        sb.AppendFormatArray Result, LocalArgs
        
        Result = sb.ToString
    End If
    
    GetString = Result
End Function


