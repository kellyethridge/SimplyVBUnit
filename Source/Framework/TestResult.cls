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
' Module: SComponent.TestResult
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
    If Not mTest Is Nothing Then
        Name = mTest.Name
    End If
End Property

Public Property Get FullName() As String
    If Not mTest Is Nothing Then
        FullName = mTest.FullName
    End If
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
    SetResult SuccessState, Message
End Sub

Public Property Get IsFailure() As Boolean
    IsFailure = (mResultState = FailureState)
End Property

Public Sub Failure(Optional ByVal Message As String, Optional ByVal Site As FailureSite = TestSite)
    SetResult FailureState, Message, Site
End Sub

Public Property Get IsError() As Boolean
    IsError = (mResultState = ErrorState)
End Property

Public Sub Error(Optional ByVal Message As String, Optional ByVal Site As FailureSite = TestSite)
    SetResult ErrorState, Message, Site
End Sub

Public Property Get IsIgnored() As Boolean
    IsIgnored = (mResultState = IgnoredState)
End Property

Public Sub Ignore(Optional ByVal Reason As String)
    SetResult IgnoredState, Reason
End Sub

Public Property Get IsInvalid() As Boolean
    IsInvalid = (mResultState = NotRunnableState)
End Property

Public Sub Invalid(Optional ByVal Reason As String)
    SetResult NotRunnableState, Reason
End Sub

Public Property Get Test() As ITest
    Set Test = mTest
End Property

Public Property Get Results() As TestResultList
    Set Results = mResults
End Property

Public Sub AddResult(ByVal Result As TestResult)
    If Result Is Nothing Then _
        modStatics.Error.CannotBeNothing Me, "AddResult", "Result"
        
    mResults.Add Result
    SetResultState Result.ResultState
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
        Me.Failure "Child test failed", ChildSite
    End If
End Sub

Private Function CanSetFailure() As Boolean
    CanSetFailure = (mResultState <> ErrorState)
End Function

Private Sub TrySetSuccessState()
    If CanSetSuccess Then
        Me.Success
    End If
End Sub

Private Function CanSetSuccess() As Boolean
    CanSetSuccess = (mResultState = InconclusiveState)
End Function

Private Sub SetCancelledState()
    SetResult CancelledState, "test cancelled by user"
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Test As ITest)
    Set mTest = Test
End Sub
