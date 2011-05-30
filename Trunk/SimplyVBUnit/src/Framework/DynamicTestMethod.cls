VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DynamicTestMethod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Copyright 2010 Kelly Ethridge
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
' Module: ArgTestMethod
'
Option Explicit
Implements ITestMethod

Private mTestClass      As Object
Private mMember         As MemberInfo
Private mTestCaseData   As TestCaseData


Public Sub Run()
    SetupExpectedError
    
    If ExpectsResult Then
        Dim Result As Variant
        VariantCopyInd Result, tli.InvokeHookArray(mTestClass, mMember.MemberId, INVOKE_FUNC, mTestCaseData.Arguments)
        
        Assert.That Result, mTestCaseData.Expected
    Else
        tli.InvokeHookArraySub mTestClass, mMember.MemberId, INVOKE_FUNC, mTestCaseData.Arguments
    End If
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

Private Function ExpectsResult() As Boolean
    Dim Result As Boolean
    Select Case mMember.ReturnType
        Case VT_VOID
            Result = False
        Case VT_HRESULT
            Result = LastParameterIsNotNamed
        Case Else
            Result = True
    End Select
    
    ExpectsResult = Result
End Function

Private Function LastParameterIsNotNamed() As Boolean
    LastParameterIsNotNamed = (Len(mMember.Parameters(mMember.Parameters.Count).Name) = 0)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ITestMethod Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ITestMethod_Run()
    Run
End Sub

Private Property Get ITestMethod_TestClass() As Object
    Set ITestMethod_TestClass = TestClass
End Property

