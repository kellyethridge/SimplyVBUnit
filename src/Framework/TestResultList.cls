VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestResultList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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
' Module: TestResultList
'
Option Explicit

Private mResults As New Collection

Public Sub Add(ByVal Result As TestResult)
    If Result Is Nothing Then _
        Call Error.CannotBeNothing(Me, "Add", "Result")
'        Call Err.Raise(ErrorCode.ArgumentNull, Resource.GetString(Src_TestResultList_Add), Resource.GetString(Arg_CannotBeNothing, "Result"))
        
        
    Call mResults.Add(Result)
End Sub

Public Property Get Item(ByVal Index As Long) As TestResult
Attribute Item.VB_UserMemId = 0
    Set Item = mResults(Index + 1)
End Property

Public Property Get Count() As Long
    Count = mResults.Count
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mResults.[_NewEnum]
End Function
