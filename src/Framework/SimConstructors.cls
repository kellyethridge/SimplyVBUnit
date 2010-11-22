VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SimConstructors"
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
' Module: SVBUnitConstructors
'
Option Explicit

Public Function NewTestFixture(ByVal TestClass As Object) As TestFixture
    Dim Suite As New TestSuite
    Suite.InitTestFixture TestClass
    Set NewTestFixture = Suite
End Function

Public Function NewTestResult(ByVal Test As ITest) As TestResult
    Set NewTestResult = New TestResult
    NewTestResult.Init Test
End Function

Public Function NewTestCase(ByVal Name As String, ByVal TestMethod As ITestMethod) As TestCase
    Set NewTestCase = New TestCase
    NewTestCase.Init Name, TestMethod
End Function

Public Function NewStaticTestMethod(ByVal TestClass As Object, ByVal Member As MemberInfo) As StaticTestMethod
    Set NewStaticTestMethod = New StaticTestMethod
    NewStaticTestMethod.Init TestClass, Member
End Function

Public Function NewTestSuite(ByVal Name As String) As TestSuite
    Set NewTestSuite = New TestSuite
    NewTestSuite.InitTestSuite Name
End Function

Public Function NewErrorInfo(ByVal Number As Long, Optional ByVal Source As String, Optional ByVal Description As String, Optional ByVal LastDllError As Long) As ErrorInfo
    Set NewErrorInfo = New ErrorInfo
    NewErrorInfo.Init Number, Source, Description, LastDllError
End Function

Public Function NewDynamicTestMethod(ByVal TestClass As Object, ByVal Member As MemberInfo, ByVal TestCaseData As TestCaseData) As DynamicTestMethod
    Set NewDynamicTestMethod = New DynamicTestMethod
    NewDynamicTestMethod.Init TestClass, Member, TestCaseData
End Function

Public Function NewTestListEnumerator(ByVal List As TestList) As TestListEnumerator
    Set NewTestListEnumerator = New TestListEnumerator
    NewTestListEnumerator.Init List
End Function

Public Function NewTestCaseCollector(ByVal MethodName As String, ByVal DataStore As Collection) As TestCaseCollector
    Set NewTestCaseCollector = New TestCaseCollector
    NewTestCaseCollector.Init MethodName, DataStore
End Function

Public Function NewEqualConstraint(ByRef Expected As Variant, Optional ByVal Expression As ConstraintExpression) As EqualConstraint
    Set NewEqualConstraint = New EqualConstraint
    NewEqualConstraint.Init Expected, Expression
End Function

Public Function NewArrayEnumerator(ByRef Arr As Variant) As ArrayEnumerator
    Set NewArrayEnumerator = New ArrayEnumerator
    NewArrayEnumerator.Init Arr
End Function

Public Function NewEnumVariantEnumerator(ByVal Obj As Object) As EnumVariantEnumerator
    Set NewEnumVariantEnumerator = New EnumVariantEnumerator
    NewEnumVariantEnumerator.Init Obj
End Function

Public Function NewTestOutput(ByVal Text As String, ByVal OutputType As TestOutputType) As TestOutput
    Set NewTestOutput = New TestOutput
    NewTestOutput.Init Text, OutputType
End Function

Public Function NewContextMethods(ByVal SetupMethod As ITestMethod, ByVal TeardownMethod As ITestMethod) As ContextMethods
    Set NewContextMethods = New ContextMethods
    NewContextMethods.Init SetupMethod, TeardownMethod
End Function

Public Function NewTestRunner(ByVal Test As ITest) As TestRunner
    Set NewTestRunner = New TestRunner
    NewTestRunner.Init Test
End Function

Public Function NewNotFilter(ByVal BaseFilter As ITestFilter) As NotFilter
    Set NewNotFilter = New NotFilter
    NewNotFilter.Init BaseFilter
End Function

Public Function NewAndConstraint(ByVal LeftSide As IConstraint, ByVal RightSide As IConstraint) As AndConstraint
    Set NewAndConstraint = New AndConstraint
    NewAndConstraint.Init LeftSide, RightSide
End Function

Public Function NewOrConstraint(ByVal LeftSide As IConstraint, ByVal RightSide As IConstraint) As OrConstraint
    Set NewOrConstraint = New OrConstraint
    NewOrConstraint.Init LeftSide, RightSide
End Function

Public Function NewNotConstraint(ByVal BaseConstraint As IConstraint) As NotConstraint
    Set NewNotConstraint = New NotConstraint
    NewNotConstraint.Init BaseConstraint
End Function

Public Function NewThrowsConstraint(ByVal Expected As ErrorInfo) As ThrowsConstraint
    Set NewThrowsConstraint = New ThrowsConstraint
    NewThrowsConstraint.Init Expected
End Function

Public Function NewTestCaseModifier(ByVal Data As TestCaseData) As TestCaseModifier
    Set NewTestCaseModifier = New TestCaseModifier
    NewTestCaseModifier.Init Data
End Function

Public Function NewTolerance(ByRef Amount As Variant) As Tolerance
    Set NewTolerance = New Tolerance
    NewTolerance.Init Amount
End Function

Public Function NewTally(ByRef Value As Variant) As Tally
    Set NewTally = New Tally
    NewTally.Init Value
End Function

Public Function NewComparisonConstraint(ByRef Expected As Variant, ByVal CanBeLess As Boolean, ByVal CanBeGreater As Boolean, ByVal CanBeEqual As Boolean, ByRef Predicate As String, Optional ByVal Expression As ConstraintExpression) As ComparisonConstraint
    Set NewComparisonConstraint = New ComparisonConstraint
    NewComparisonConstraint.Init Expected, CanBeLess, CanBeGreater, CanBeEqual, Predicate, Expression
End Function

Public Function NewRangeConstraint(ByRef FromValue As Variant, ByRef ToValue As Variant, Optional ByVal Expression As ConstraintExpression) As RangeConstraint
    Set NewRangeConstraint = New RangeConstraint
    NewRangeConstraint.Init FromValue, ToValue, Expression
End Function

Public Function NewResolvableConstraint(ByVal InnerConstraint As IConstraint, ByVal Expression As ConstraintExpression) As ResolvableConstraint
    Set NewResolvableConstraint = New ResolvableConstraint
    NewResolvableConstraint.Init InnerConstraint, Expression
End Function

Public Function NewSameAsConstraint(ByVal Expected As Object) As SameAsConstraint
    Set NewSameAsConstraint = New SameAsConstraint
    NewSameAsConstraint.Init Expected
End Function

Public Function NewUniqueItemsConstraint(Optional ByVal Expression As ConstraintExpression) As UniqueItemsConstraint
    Set NewUniqueItemsConstraint = New UniqueItemsConstraint
    NewUniqueItemsConstraint.Init Expression
End Function

Public Function NewEquivalentConstraint(ByRef Expected As Variant, Optional ByVal Expression As ConstraintExpression) As EquivalentConstraint
    Set NewEquivalentConstraint = New EquivalentConstraint
    NewEquivalentConstraint.Init Expected, Expression
End Function

Public Function NewSubsetOfConstraint(ByRef Expected As Variant, Optional ByVal Expression As ConstraintExpression) As SubsetOfConstraint
    Set NewSubsetOfConstraint = New SubsetOfConstraint
    NewSubsetOfConstraint.Init Expected, Expression
End Function

Public Function NewSubstringConstraint(ByRef Expected As String, Optional ByVal Expression As ConstraintExpression) As SubstringConstraint
    Set NewSubstringConstraint = New SubstringConstraint
    NewSubstringConstraint.Init Expected, Expression
End Function

Public Function NewStartsWithConstraint(ByRef Expected As String, Optional ByVal Expression As ConstraintExpression) As StartsWithConstraint
    Set NewStartsWithConstraint = New StartsWithConstraint
    NewStartsWithConstraint.Init Expected, Expression
End Function

Public Function NewEndsWithConstraint(ByRef Expected As String, Optional ByVal Expression As ConstraintExpression) As EndsWithConstraint
    Set NewEndsWithConstraint = New EndsWithConstraint
    NewEndsWithConstraint.Init Expected, Expression
End Function

Public Function NewRegexConstraint(ByRef Pattern As String, Optional ByVal Expression As ConstraintExpression) As RegexConstraint
    Set NewRegexConstraint = New RegexConstraint
    NewRegexConstraint.Init Pattern, Expression
End Function

Public Function NewSamePathConstraint(ByRef Expected As String, Optional ByVal Expression As ConstraintExpression) As SamePathConstraint
    Set NewSamePathConstraint = New SamePathConstraint
    NewSamePathConstraint.Init Expected, Expression
End Function

Public Function NewSamePathOrUnderConstraint(ByRef Expected As String, Optional ByVal Expression As ConstraintExpression) As SamePathOrUnderConstraint
    Set NewSamePathOrUnderConstraint = New SamePathOrUnderConstraint
    NewSamePathOrUnderConstraint.Init Expected, Expression
End Function
