VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArgOnlyTestMethod"
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
' Module: SComponent.ArgOnlyTestMethod
'
Option Explicit
Implements ITestMethod

Private mTestClass      As Object
Private mMember         As MemberInfo
Private mTestCaseData   As TestCaseData


Public Sub Run()
    SetupExpectedError
    TLI.InvokeHookArraySub mTestClass, mMember.MemberId, INVOKE_FUNC, mTestCaseData.Arguments
End Sub

Private Sub SetupExpectedError()
    If HasExpectedError Then
        AssertExpectedError
    End If
End Sub

Private Function HasExpectedError() As Boolean
    HasExpectedError = (Not mTestCaseData.ErrorInfo Is Nothing)
End Function

Private Sub AssertExpectedError()
    Dim Info As ErrorInfo
    Set Info = mTestCaseData.ErrorInfo
    
    Assert.Throws Info.Number, Info.Source, Info.Description
End Sub

Public Property Get TestClass() As Object
    Set TestClass = mTestClass
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal TestClass As Object, ByVal Member As MemberInfo, ByVal TestCaseData As TestCaseData)
    Set mTestClass = TestClass
    Set mMember = Member
    Set mTestCaseData = TestCaseData
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ITestMethod Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ITestMethod_Run()
    Run
End Sub

Private Property Get ITestMethod_TestClass() As Object
    Set ITestMethod_TestClass = TestClass
End Property

