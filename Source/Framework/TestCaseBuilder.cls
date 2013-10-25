VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestCaseBuilder"
Attribute VB_GlobalNameSpace = False
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
' Module: TestCaseBuilder
'
Option Explicit

Private mTestCases      As Collection
Private mMember         As MemberInfo
Private mTestClass      As Object
Private mTestName       As String
Private mNameBuilder    As New StringBuilder
Private mNameSequence   As Long


Public Function BuildFrom(ByVal TestClass As Object, ByVal Member As MemberInfo) As Collection
    Init TestClass, Member
    BuildTestCases
    Set BuildFrom = Release
End Function

Private Sub Init(ByVal TestClass As Object, ByVal Member As MemberInfo)
    Set mTestCases = New Collection
    Set mMember = Member
    Set mTestClass = TestClass
    mNameSequence = 0
    mTestName = Member.Name
End Sub

Private Sub BuildTestCases()
    If HasNoParameters Then
        BuildSimpleTestMethod
    ElseIf IsParameterized Then
        BuildParameterizedTestMethod
    End If
End Sub

Private Function Release() As Collection
    Set mMember = Nothing
    Set mTestClass = Nothing
    Set Release = mTestCases
    Set mTestCases = Nothing
End Function

Private Function IsParameterized() As Boolean
    IsParameterized = HasParameters And CanProvideArguments
End Function

Private Function HasParameters() As Boolean
    HasParameters = (mMember.Parameters.Count > 0)
End Function

Private Function CanProvideArguments() As Boolean
    CanProvideArguments = (TypeOf mTestClass Is ITestCaseSource)
End Function

Private Function HasNoParameters() As Boolean
    HasNoParameters = Not HasParameters
End Function

Private Sub BuildSimpleTestMethod()
    Dim Method As ITestMethod
    Set Method = Sim.NewSimpleTestMethod(mTestClass, mMember.MemberId)
    AddTestCase Method
End Sub

Private Sub BuildParameterizedTestMethod()
    Dim Suite As TestSuite
    Set Suite = Sim.NewTestSuite(mMember.Name)
    mTestCases.Add Suite
    
    Dim Data As TestCaseData
    For Each Data In GetTestCases
        Dim TestCase As TestCase
        Set TestCase = MakeTestCaseFromData(Data)
        Suite.Add TestCase
    Next
End Sub

Private Function MakeTestCaseFromData(ByVal Data As TestCaseData) As TestCase
    Dim Method As ITestMethod
    Set Method = Sim.NewParameterizedTestMethod(mTestClass, mMember.MemberId, Data)
    
    Set MakeTestCaseFromData = Sim.NewTestCase(BuildTestName, BuildTestDisplayName(Data.Arguments), Method)
End Function

Private Function BuildTestDisplayName(ByRef Args() As Variant) As String
    mNameBuilder.Length = 0
    mNameBuilder.Append mTestName
    mNameBuilder.Append " ("
    
    Dim i As Long
    For i = UBound(Args) To LBound(Args) Step -1
        If i < UBound(Args) Then
            mNameBuilder.Append ", "
        End If
        
        Select Case VarType(Args(i))
            Case vbByte
                mNameBuilder.Append "CByte(" & Args(i) & ")"
            Case vbLong
                mNameBuilder.Append Args(i) & "&"
            Case vbSingle
                mNameBuilder.Append Args(i) & "!"
            Case vbDouble
                mNameBuilder.Append Args(i) & "#"
            Case vbCurrency
                mNameBuilder.Append Args(i) & "@"
            Case vbDecimal
                mNameBuilder.Append "CDec(" & Args(i) & ")"
            Case vbString
                mNameBuilder.Append """" & Args(i) & """"
                Case vbDate
                mNameBuilder.Append "#" & Args(i) & "#"
            Case vbInteger, vbBoolean
                mNameBuilder.Append Args(i)
            Case Else
                mNameBuilder.Append TypeName(Args(i))
        End Select
    Next
    
    mNameBuilder.Append ")"
    BuildTestDisplayName = mNameBuilder.ToString
End Function

Private Function BuildTestName() As String
    mNameSequence = mNameSequence + 1
    BuildTestName = mTestName & "_" & Format$(mNameSequence, "000")
End Function

Private Function GetTestCases() As Collection
    Dim Builder As TestCaseCollector
    Set Builder = Sim.NewTestCaseCollector(mMember.Name)
    
    Dim Source As ITestCaseSource
    Set Source = mTestClass
    On Error GoTo errTrap
    Source.GetTestCases Builder
    
errTrap:
    Set GetTestCases = Builder.DataStore
End Function

Private Sub AddTestCase(ByVal Method As ITestMethod)
    Dim TestCase As TestCase
    Set TestCase = Sim.NewTestCase(mTestName, mTestName, Method)
    mTestCases.Add TestCase
End Sub










