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
' Module: SComponent.MemberQuery
'
Option Explicit

Private mResult As QueryResult
Private mInfo   As InterfaceInfo


Public Function Execute(ByVal TestClass As Object) As QueryResult
    If InitedSuccessfully(TestClass) Then
        FindMembers
    End If
    
    Set Execute = Release
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function InitedSuccessfully(ByVal TestClass As Object) As Boolean
    Set mResult = New QueryResult
    
    On Error GoTo errTrap
    Set mInfo = TLI.InterfaceInfoFromObject(TestClass).VTableInterface
    InitedSuccessfully = True
    
errTrap:
End Function

Private Sub FindMembers()
    Dim Member As MemberInfo
    
    For Each Member In mInfo.Members
        ProcessMember Member
    Next
End Sub

Private Sub ProcessMember(ByVal Member As MemberInfo)
    If IncludeMember(Member) Then
        mResult.AddMember Member
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
