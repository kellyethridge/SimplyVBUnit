VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QueryResult"
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
' Module: QueryResult
'
Option Explicit

Private mSetupMember            As Variant
Private mTeardownMember         As Variant
Private mFixtureSetupMember     As Variant
Private mFixtureTeardownMember  As Variant
Private mTestCaseMembers        As New Collection



Public Property Get SetupMember() As Variant
    SetupMember = mSetupMember
End Property

Public Property Get TeardownMember() As Variant
    TeardownMember = mTeardownMember
End Property

Public Property Get FixtureSetupMember() As Variant
    FixtureSetupMember = mFixtureSetupMember
End Property

Public Property Get FixtureTeardownMember() As Variant
    FixtureTeardownMember = mFixtureTeardownMember
End Property

Public Property Get TestCaseMembers() As Collection
    Set TestCaseMembers = mTestCaseMembers
End Property

Public Sub AddMember(ByVal Member As MemberInfo)
    Select Case LCase$(Member.Name)
        Case "setup":           mSetupMember = Member.MemberId
        Case "teardown":        mTeardownMember = Member.MemberId
        Case "fixturesetup":    mFixtureSetupMember = Member.MemberId
        Case "fixtureteardown": mFixtureTeardownMember = Member.MemberId
        Case Else:              mTestCaseMembers.Add Member
    End Select
End Sub
