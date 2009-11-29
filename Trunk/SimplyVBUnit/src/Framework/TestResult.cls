VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestResult"
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
' Module: TestResult
'
Option Explicit

Private mTest           As ITest
Private mResults        As New TestResultList
Private mMessage        As String
Private mResultState    As ResultState
Private mTime           As Currency
Private mFailureSite    As FailureSite
Private mAssertCount    As Long


Public Property Get AssertCount() As Long
    AssertCount = mAssertCount
End Property

Public Property Let AssertCount(ByVal RHS As Long)
    mAssertCount = RHS
End Property

Public Property Get Time() As Currency
    Time = mTime
End Property

Public Property Let Time(ByVal RHS As Currency)
    mTime = RHS
End Property

Public Property Get Name() As String
    Name = mTest.Name
End Property

Public Property Get FullName() As String
    FullName = mTest.FullName
End Property

Public Property Get HasResults() As Boolean
    HasResults = (mResults.Count > 0)
End Property

Public Property Get ResultState() As ResultState
    ResultState = mResultState
End Property

Public Sub SetResult(ByVal State As ResultState, Optional ByVal Message As String, Optional ByVal Site As FailureSite = TestSite)
    If State = ErrorState And Site = TeardownSite Then
        If Len(mMessage) > 0 Then
            Message = mMessage & vbCrLf & "Teardown: " & Message
        End If
    End If

    mResultState = State
    mMessage = Message
    mFailureSite = Site
End Sub

Public Property Get FailureSite() As FailureSite
    FailureSite = mFailureSite
End Property

Public Property Get Message() As String
    Message = mMessage
End Property

Public Property Get IsSuccess() As Boolean
    IsSuccess = (mResultState = SuccessState)
End Property

Public Sub Success(Optional ByVal Message As String)
    Call SetResult(SuccessState, Message)
End Sub

Public Property Get IsFailure() As Boolean
    IsFailure = (mResultState = FailureState)
End Property

Public Sub Failure(Optional ByVal Message As String, Optional ByVal Site As FailureSite = TestSite)
    Call SetResult(FailureState, Message, Site)
End Sub

Public Property Get IsError() As Boolean
    IsError = (mResultState = ErrorState)
End Property

Public Sub Error(Optional ByVal Message As String, Optional ByVal Site As FailureSite = TestSite)
    Call SetResult(ErrorState, Message, Site)
End Sub

Public Property Get IsIgnored() As Boolean
    IsIgnored = (mResultState = IgnoredState)
End Property

Public Sub Ignore(Optional ByVal Reason As String)
    Call SetResult(IgnoredState, Reason)
End Sub

Public Property Get IsInvalid() As Boolean
    IsInvalid = (mResultState = NotRunnableState)
End Property

Public Sub Invalid(Optional ByVal Reason As String)
    Call SetResult(NotRunnableState, Reason)
End Sub

Public Property Get Test() As ITest
    Set Test = mTest
End Property

Public Property Get Results() As TestResultList
    Set Results = mResults
End Property

Public Sub AddResult(ByVal Result As TestResult)
    If Result Is Nothing Then _
        Call modStatics.Error.CannotBeNothing(Me, "AddResult", "Result")
        
    Call mResults.Add(Result)
    Call SetResultState(Result.ResultState)
End Sub

Private Sub SetResultState(ByVal ChildState As ResultState)
    Select Case ChildState
        Case FailureState, ErrorState:  TrySetFailureState
        Case SuccessState:              TrySetSuccessState
        Case CancelledState:            SetCancelledState
    End Select
End Sub

Private Sub TrySetFailureState()
    If CanSetFailure Then
        Call Me.Failure("Child test failed", ChildSite)
    End If
End Sub

Private Function CanSetFailure() As Boolean
    CanSetFailure = (mResultState <> ErrorState)
End Function

Private Sub TrySetSuccessState()
    If CanSetSuccess Then
        Call Me.Success
    End If
End Sub

Private Function CanSetSuccess() As Boolean
    CanSetSuccess = (mResultState = InconclusiveState)
End Function

Private Sub SetCancelledState()
    Call Me.SetResult(CancelledState)
End Sub


Friend Sub Init(ByVal Test As ITest)
    Set mTest = Test
End Sub
