VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemberQuery"
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
' Module: MemberQuery
'
Option Explicit

Private mResult As QueryResult
Private mInfo   As InterfaceInfo


Public Function Execute(ByVal TestClass As Object) As QueryResult
    Call Init(TestClass)
    Call FindMembers
    Set Execute = Release
End Function

Private Sub Init(ByVal TestClass As Object)
    Set mResult = New QueryResult
    Set mInfo = TLI.InterfaceInfoFromObject(TestClass)
End Sub

Private Sub FindMembers()
    Dim Member As MemberInfo
    
    For Each Member In mInfo.Members
        Call ProcessMember(Member)
    Next
End Sub

Private Sub ProcessMember(ByVal Member As MemberInfo)
    If IncludeMember(Member) Then
        Call mResult.AddMember(Member)
    End If
End Sub

Private Function IncludeMember(ByVal Member As MemberInfo) As Boolean
    IncludeMember = (Member.InvokeKind = INVOKE_FUNC)
End Function

Private Function Release() As QueryResult
    Set Release = mResult
    Set mResult = Nothing
    Set mInfo = Nothing
End Function