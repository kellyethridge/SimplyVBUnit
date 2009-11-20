VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SVBUnitConstructors"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
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
' Module: SVBUnitConstructors
'
Option Explicit

Public Function NewTestFixture(ByVal TestClass As Object) As TestFixture
    Set NewTestFixture = New TestFixture
    Call NewTestFixture.Init(TestClass)
End Function

Public Function NewTestResult(ByVal Test As ITest) As TestResult
    Set NewTestResult = New TestResult
    Call NewTestResult.Init(Test)
End Function

Public Function NewTestCase(ByVal Name As String, ByVal TestMethod As ITestMethod) As TestCase
    Set NewTestCase = New TestCase
    Call NewTestCase.Init(Name, TestMethod)
End Function

Public Function NewNoArgTestMethod(ByVal TestClass As Object, ByVal Member As MemberInfo) As NoArgTestMethod
    Set NewNoArgTestMethod = New NoArgTestMethod
    Call NewNoArgTestMethod.Init(TestClass, Member)
End Function

Public Function NewTestSuite(ByVal Name As String) As TestSuite
    Set NewTestSuite = New TestSuite
    Call NewTestSuite.Init(Name)
End Function

Public Function NewErrorInfo(ByVal Number As Long, Optional ByVal Source As String, Optional ByVal Description As String) As ErrorInfo
    Set NewErrorInfo = New ErrorInfo
    Call NewErrorInfo.Init(Number, Source, Description)
End Function

Public Function NewArgOnlyTestMethod(ByVal TestClass As Object, ByVal Member As MemberInfo, ByVal TestCaseData As TestCaseData) As ArgOnlyTestMethod
    Set NewArgOnlyTestMethod = New ArgOnlyTestMethod
    Call NewArgOnlyTestMethod.Init(TestClass, Member, TestCaseData)
End Function

Public Function NewTestListEnumerator(ByVal List As TestList) As TestListEnumerator
    Set NewTestListEnumerator = New TestListEnumerator
    Call NewTestListEnumerator.Init(List)
End Function

Public Function NewTestCaseDataBuilder(ByVal MethodName As String, ByVal DataStore As Collection) As TestCaseDataBuilder
    Set NewTestCaseDataBuilder = New TestCaseDataBuilder
    Call NewTestCaseDataBuilder.Init(MethodName, DataStore)
End Function

Public Function NewEqualConstraint(ByRef Expected As Variant) As EqualConstraint
    Set NewEqualConstraint = New EqualConstraint
    Call NewEqualConstraint.Init(Expected)
End Function

Public Function NewArrayEnumerator(ByRef Arr As Variant) As ArrayEnumerator
    Set NewArrayEnumerator = New ArrayEnumerator
    Call NewArrayEnumerator.Init(Arr)
End Function

Public Function NewEnumVariantEnumerator(ByVal Obj As Object) As EnumVariantEnumerator
    Set NewEnumVariantEnumerator = New EnumVariantEnumerator
    Call NewEnumVariantEnumerator.Init(Obj)
End Function

Public Function NewTestOutput(ByVal Text As String, ByVal OutputType As TestOutputType) As TestOutput
    Set NewTestOutput = New TestOutput
    Call NewTestOutput.Init(Text, OutputType)
End Function

Public Function NewContextMethods(ByVal SetupMethod As ITestMethod, ByVal TeardownMethod As ITestMethod) As ContextMethods
    Set NewContextMethods = New ContextMethods
    Call NewContextMethods.Init(SetupMethod, TeardownMethod)
End Function