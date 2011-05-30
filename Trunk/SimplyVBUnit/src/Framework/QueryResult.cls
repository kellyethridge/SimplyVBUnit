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
