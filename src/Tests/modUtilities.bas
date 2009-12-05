Attribute VB_Name = "modUtilities"
Option Explicit

Private Const Q As String = """"

Public Function Quote(ByVal Text As String) As String
    Quote = Q & Text & Q
End Function

Public Sub AssertError(ByVal ActualError As ErrObject, ByVal ExpectedNumber As Long, Optional ByVal ExpectedSource As String, Optional ByVal ExpectedDescription As String)
    Dim Actual      As ErrorInfo
    Dim Expected    As ErrorInfo
    
    Set Actual = ErrorInfo.FromErr(ActualError)
    Set Expected = Sim.NewErrorInfo(ExpectedNumber, ExpectedSource, ExpectedDescription)
    
    If Actual.Number = ErrorCode.NoError Then
        Call Err.Raise(AssertCode.FailureCode, , "Expected an error to be raised.")
    ElseIf Actual.Equals(Expected) = False Then
        Call Err.Raise(AssertCode.FailureCode, , "Wrong error raised.")
    End If
End Sub

Public Sub AssertCalls(ByVal ActualCalls As CallTrace, ParamArray ExpectedCalls() As Variant)
    Dim Expected    As New CallTrace
    Dim Name        As Variant
    
    For Each Name In ExpectedCalls
        Call Expected.Add(Name)
    Next
    
    If ActualCalls.Equals(Expected) = False Then
        Call Err.Raise(AssertCode.FailureCode, , "Expected Calls: [" & Expected.ToString & "] - Actual Calls: [" & ActualCalls.ToString & "]")
    End If
End Sub

Public Sub AssertNoCalls(ByVal ActualCalls As CallTrace)
    Call AssertCalls(ActualCalls)
End Sub
