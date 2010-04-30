VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArgWithReturnTestMethod"
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
' Module: ArgWithReturnTestMethod
'
Option Explicit
Implements ITestMethod

Private mTestClass  As Object
Private mMember     As MemberInfo
Private mData       As TestCaseData


Public Sub Run()
    Dim Result As Variant
    
    Call VariantCopyInd(Result, tli.InvokeHookArray(mTestClass, mMember.MemberId, INVOKE_FUNC, mData.Arguments))
    
    Assert.That Result, mData.Expected
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
    Set mData = TestCaseData
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ITestMethod Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ITestMethod_Run()
    Call Run
End Sub

Private Property Get ITestMethod_TestClass() As Object
    Set ITestMethod_TestClass = TestClass
End Property