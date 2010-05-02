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

Public Function NewCollection(ParamArray Values() As Variant) As Collection
    Dim Result  As New Collection
    Dim Item    As Variant
    
    For Each Item In Values
        Call Result.Add(Item)
    Next
    
    Set NewCollection = Result
End Function

Public Function MakeLongArray(ByVal LowerBound As Long, ParamArray Args() As Variant) As Long()
    Dim Result() As Long
    ReDim Result(LowerBound To LowerBound + UBound(Args))
    
    Dim i As Long
    For i = 0 To UBound(Args)
        Result(LowerBound + i) = Args(i)
    Next
    
    MakeLongArray = Result
End Function

Public Sub AssertEmptyArray(ByRef Arr As Variant)
    Dim lb As Long
    Dim ub As Long
    
    lb = LBound(Arr)
    lb = UBound(Arr)
    
    If lb <= ub Then
        Err.Raise AssertCode.FailureCode, , "Array should be empty."
    End If
End Sub

Public Function NewLongs(ByVal Size As Long) As Long()
    SAPtrLong(NewLongs) = SafeArrayCreateVector(vbLong, 0, Size)
End Function
