VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParameterizedTestMethod"
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
' Module: SComponent.ParameterizedTestMethod
'
Option Explicit
Implements ITestMethod

Private mTestClass      As Object
Private mMemberId       As Variant
Private mTestCaseData   As TestCaseData


Public Sub Run()
    SetupExpectedError
    
    If mTestCaseData.Expected Is Nothing Then
        TLI.InvokeHookArraySub mTestClass, mMemberId, INVOKE_FUNC, mTestCaseData.Arguments
    Else
        Dim Result As Variant
        VariantCopyInd Result, TLI.InvokeHookArray(mTestClass, mMemberId, INVOKE_FUNC, mTestCaseData.Arguments)
        
        Assert.That Result, mTestCaseData.Expected
    End If
End Sub

Public Property Get TestClass() As Object
    Set TestClass = mTestClass
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal TestClass As Object, ByRef MemberId As Variant, ByVal TestCaseData As TestCaseData)
    If TestClass Is Nothing Then _
        Error.CannotBeNothing Me, "Init", "TestClass"
    If TestCaseData Is Nothing Then _
        Error.CannotBeNothing Me, "Init", "TestCaseData"
    
    Set mTestClass = TestClass
    Set mTestCaseData = TestCaseData
    
    Select Case VarType(MemberId)
        Case vbString
            If Len(MemberId) = 0 Then _
                Error.CannotBeEmpty Me, "Init", "MemberId"
                
            mMemberId = MemberId
            
        Case vbLong, vbInteger, vbByte
            mMemberId = MemberId
        
        Case Else
            Error.Argument Me, "Init", "MemberId", "Must be a String or Number."
            
    End Select
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ITestMethod Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ITestMethod_Run()
    Run
End Sub

Private Property Get ITestMethod_TestClass() As Object
    Set ITestMethod_TestClass = TestClass
End Property

