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
        Err.Raise AssertCode.FailureCode, , "Expected an error to be raised."
    ElseIf Actual.Equals(Expected) = False Then
        Err.Raise AssertCode.FailureCode, , "Wrong error raised. Was '" & Err.Description & "'."
    End If
End Sub

Public Sub AssertCalls(ByVal ActualCalls As CallTrace, ParamArray ExpectedCalls() As Variant)
    Dim Expected    As New CallTrace
    Dim Name        As Variant
    
    For Each Name In ExpectedCalls
        Expected.Add Name
    Next
    
    If ActualCalls.Equals(Expected) = False Then
        Err.Raise AssertCode.FailureCode, , "Expected Calls: [" & Expected.ToString & "] - Actual Calls: [" & ActualCalls.ToString & "]"
    End If
End Sub

Public Sub AssertNoCalls(ByVal ActualCalls As CallTrace)
    AssertCalls ActualCalls
End Sub

Public Function NewCollection(ParamArray Values() As Variant) As Collection
    Dim Result  As New Collection
    Dim Item    As Variant
    
    For Each Item In Values
        Result.Add Item
    Next
    
    Set NewCollection = Result
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

Public Function NewLongs(ParamArray Values() As Variant) As Long()
    Dim RetVal() As Long
    PutSAPtr RetVal, SafeArrayCreateVector(vbLong, 0, UBound(Values) + 1)
    
    Dim i As Long
    For i = 0 To UBound(Values)
        RetVal(i) = Values(i)
    Next
    
    NewLongs = RetVal
End Function

Public Function NewLongsLb(ByVal LowerBound As Long, ParamArray Values() As Variant) As Long()
    Dim RetVal() As Long
    PutSAPtr RetVal, SafeArrayCreateVector(vbLong, LowerBound, UBound(Values) + 1)
    
    Dim i As Long
    For i = 0 To UBound(Values)
        RetVal(i + LowerBound) = Values(i)
    Next
    
    NewLongsLb = RetVal
End Function

Public Function NewDoubles(ParamArray Values() As Variant) As Double()
    Dim RetVal() As Double
    PutSAPtr RetVal, SafeArrayCreateVector(vbDouble, 0, UBound(Values) + 1)
    
    Dim i As Long
    For i = 0 To UBound(Values)
        RetVal(i) = Values(i)
    Next
    
    NewDoubles = RetVal
End Function

