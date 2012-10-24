VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SimStatics"
Attribute VB_GlobalNameSpace = True
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
' Module: SimStatics
'
Option Explicit

Public Enum ErrorCode
    NoError = 0
    Argument = 5
    ArgumentNull = 5
    ArgumentOutOfRange = 380
    ArgumentNotOptional = 449
    IndexOutOfRange = 9
    InvalidOperation = 5
    
    #If False Then
        Dim NoError, Argument, ArgumentNull, ArgumentOutOfRange, ArgumentNotOptional, IndexOutOfRange, InvalidOperation, InvalidConstraintExpression
    #End If
End Enum

Public Enum ResultState
    InconclusiveState
    NotRunnableState
    IgnoredState
    SuccessState
    FailureState
    ErrorState
    CancelledState
    
    #If False Then
        Dim InconclusiveState, SuccessState, FailureState, ErrorState, NotRunnableState, IgnoredState, CancelledState
    #End If
End Enum

Public Enum FailureSite
    TestSite
    SetupSite
    TeardownSite
    ParentSite
    ChildSite
    
    #If False Then
        Dim TestSite, SetupSite, TeardownSite, ParentSite, ChildSite
    #End If
End Enum

Public Enum RunState
    RunnableRunState
    NotRunnableRunState
    
    #If False Then
        Dim RunnableRunState, NotRunnableRunState
    #End If
End Enum

Private mTextBuffer As New StringBuilder

Public Function Sim() As SimConstructors
    Set Sim = modStatics.Sim
End Function

Public Function Timing() As TimingStatic
    Set Timing = modStatics.Timing
End Function

Public Function Assert() As Assertions
    Set Assert = modStatics.Assert
End Function

Public Function ErrorInfo() As ErrorInfoStatic
    Set ErrorInfo = modStatics.ErrorInfo
End Function

Public Function Iz() As IzSyntaxHelper
    Set Iz = modStatics.Iz
End Function

Public Function Has() As HasStatic
    Set Has = mHas
End Function

Public Function TestContext() As TestContextStatic
    Set TestContext = modStatics.TestContext
End Function

Public Function TestFilter() As TestFilterStatic
    Set TestFilter = modStatics.TestFilter
End Function

Public Property Get Tolerance() As ToleranceStatic
    Set Tolerance = modStatics.Tolerance
End Property

Public Sub WriteText(ByRef Text As String, ParamArray Args() As Variant)
    Dim LocalArgs() As Variant
    LocalArgs = Args
    mTextBuffer.Length = 0
    mTextBuffer.AppendFormatArray Text, LocalArgs
    TestContext.Out.WriteText mTextBuffer.ToString
End Sub

Public Sub WriteLine(ParamArray TextAndArgs() As Variant)
    Dim LocalArgs() As Variant
    LocalArgs = TextAndArgs
    mTextBuffer.Length = 0
    mTextBuffer.AppendFormatArrayEx LocalArgs
    TestContext.Out.WriteLine mTextBuffer.ToString
End Sub


