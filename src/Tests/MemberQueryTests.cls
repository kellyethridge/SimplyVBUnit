VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemberQueryTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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
' Module: MemberQueryTests
'
Option Explicit

Public Sub Query_WithEmptyTestClass_DoesNotFindSetupMember()
    Dim Fake As New FakeEmpty
    Dim Query As New MemberQuery
    
    Dim Result As QueryResult
    Set Result = Query.Execute(Fake)
    
    Dim Actual As MemberInfo
    Set Actual = Result.SetupMember
    Assert.IsTrue Actual Is Nothing
End Sub

Public Sub Query_WithEmptyTestClass_DoesNotFindTeardownMember()
    Dim Fake As New FakeEmpty
    Dim Query As New MemberQuery
    
    Dim Result As QueryResult
    Set Result = Query.Execute(Fake)
    
    Dim Actual As MemberInfo
    Set Actual = Result.TeardownMember
    Assert.IsTrue Actual Is Nothing
End Sub

Public Sub Query_WithEmptyTestClass_DoesNotFindFixtureSetupMember()
    Dim Fake As New FakeEmpty
    Dim Query As New MemberQuery
    
    Dim Result As QueryResult
    Set Result = Query.Execute(Fake)
    
    Dim Actual As MemberInfo
    Set Actual = Result.FixtureSetupMember
    Assert.IsTrue Actual Is Nothing
End Sub

Public Sub Query_WithEmptyTestClass_DoesNotFindFixtureTeardownMember()
    Dim Fake As New FakeEmpty
    Dim Query As New MemberQuery
    
    Dim Result As QueryResult
    Set Result = Query.Execute(Fake)
    
    Dim Actual As MemberInfo
    Set Actual = Result.FixtureTeardownMember
    Assert.IsTrue Actual Is Nothing
End Sub

Public Sub Query_WithEmptyTestClass_DoesNotFindTestCaseMembers()
    Dim Fake As New FakeEmpty
    Dim Query As New MemberQuery
    
    Dim Result As QueryResult
    Set Result = Query.Execute(Fake)
    
    Dim Actual As Long
    Actual = Result.TestCaseMembers.Count
    Assert.IsTrue Actual = 0
End Sub

Public Sub Query_With1Test_Finds1TestCaseMember()
    Dim Fake As New Fake1Test
    Dim Query As New MemberQuery
    
    Dim Result As QueryResult
    Set Result = Query.Execute(Fake)
    
    Dim Actual As Long
    Actual = Result.TestCaseMembers.Count
    Assert.IsTrue Actual = 1
End Sub

Public Sub Query_With2Tests_Finds2TestCaseMembers()
    Dim Fake As New Fake2Tests
    Dim Query As New MemberQuery
    
    Dim Result As QueryResult
    Set Result = Query.Execute(Fake)
    
    Dim Actual As Long
    Actual = Result.TestCaseMembers.Count
    Assert.IsTrue Actual = 2
End Sub

Public Sub Query_With1TestAndSetup_FindsSetupMember()
    Dim Fake As New Fake1TestWithSetup
    Dim Query As New MemberQuery
    
    Dim Result As QueryResult
    Set Result = Query.Execute(Fake)
    
    Dim Actual As MemberInfo
    Set Actual = Result.SetupMember
    Assert.IsTrue Not Actual Is Nothing
End Sub

Public Sub Query_With1TestAndSetup_Finds1TestCaseMember()
    Dim Fake As New Fake1TestWithSetup
    Dim Query As New MemberQuery
    
    Dim Result As QueryResult
    Set Result = Query.Execute(Fake)
    
    Dim Actual As Long
    Actual = Result.TestCaseMembers.Count
    Assert.IsTrue Actual = 1
End Sub

Public Sub Query_With1TestAndTeardown_FindsTeardownMembmer()
    Dim Fake As New Fake1TestWithTeardown
    Dim Query As New MemberQuery
    
    Dim Result As QueryResult
    Set Result = Query.Execute(Fake)
    
    Dim Actual As MemberInfo
    Set Actual = Result.TeardownMember
    Assert.IsTrue Not Actual Is Nothing
End Sub

Public Sub Query_With1TestAndFixtureSetup_FindsFixtureSetupMethod()
    Dim Fake As New Fake1TestWithFixtureSetup
    Dim Query As New MemberQuery
    
    Dim Result As QueryResult
    Set Result = Query.Execute(Fake)
    
    Dim Actual As MemberInfo
    Set Actual = Result.FixtureSetupMember
    Assert.IsTrue Not Actual Is Nothing
End Sub

Public Sub Query_With1TestAndFixtureTeardown_FindsFixtureTeardownMethod()
    Dim Fake As New Fake1TestWithFixtureTeardown
    Dim Query As New MemberQuery
    
    Dim Result As QueryResult
    Set Result = Query.Execute(Fake)
    
    Dim Actual As MemberInfo
    Set Actual = Result.FixtureTeardownMember
    Assert.IsTrue Not Actual Is Nothing
End Sub

Public Sub Query_With1Property_ExcludesProperty()
    Dim Fake As New Fake1Property
    Dim Query As New MemberQuery
    
    Dim Result As QueryResult
    Set Result = Query.Execute(Fake)
    
    Dim Actual As Long
    Actual = Result.TestCaseMembers.Count
    Assert.IsTrue Actual = 0
End Sub





