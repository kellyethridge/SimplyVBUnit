VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestCaseModifier"
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
' Module: TestCaseModifier
'
Option Explicit
Implements AssertExpression
Implements TestNameExpression

Private mData As TestCaseData


Public Function TestName(ByVal NewTestName As String) As AssertExpression
    mData.TestName = NewTestName
    Set TestName = Me
End Function

Public Function Throws(ByVal ErrorNumber As Long, Optional ByVal ErrorSource As String, Optional ByVal ErrorDescription As String) As TestNameExpression
    Set mData.ErrorInfo = Sim.NewErrorInfo(ErrorNumber, ErrorSource, ErrorDescription)
    Set Throws = Me
End Function

Public Function Expect(ByRef Expected As Variant) As TestNameExpression
    If IsConstraint(Expected) Then
        Set mData.Expected = Expected
    Else
        Set mData.Expected = Sim.NewEqualConstraint(Expected)
    End If
    
    Set Expect = Me
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IsConstraint(ByRef Expected As Variant) As Boolean
    Dim Result As Boolean
    
    If IsObject(Expected) Then
        If Not Expected Is Nothing Then
            Result = TypeOf Expected Is IConstraint
        End If
    End If
    
    IsConstraint = Result
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Data As TestCaseData)
    If Data Is Nothing Then _
        Call Error.CannotBeNothing(Me, "Init", "Data")
        
    Set mData = Data
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   AssertExpression Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AssertExpression_Expect(Expected As Variant)
    Call Expect(Expected)
End Sub

Private Sub AssertExpression_Throws(ByVal ErrorNumber As Long, Optional ByVal ErrorSource As String, Optional ByVal ErrorDescription As String)
    Call Throws(ErrorNumber, ErrorSource, ErrorDescription)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   TestNameExpression Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TestNameExpression_TestName(ByVal NewTestName As String)
    Call TestName(NewTestName)
End Sub
