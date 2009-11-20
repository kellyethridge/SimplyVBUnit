VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Assertions"
Attribute VB_GlobalNameSpace = False
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
' Module: Assertions
'
Option Explicit

Public Enum AssertCode
    FailureCode = vbObjectError + 151000
    SuccessCode
    
    #If False Then
        Dim FailureCode, SuccessCode
    #End If
End Enum

Private mCount As Long

Private Sub IncrementCount()
    mCount = mCount + 1
End Sub

Friend Function SetCount(ByVal NewCount As Long) As Long
    SetCount = mCount
    mCount = NewCount
End Function

Public Sub Fail(Optional ByVal Message As String)
    Call IncrementCount
    Call Err.Raise(AssertCode.FailureCode, , Message)
End Sub

Public Sub Pass(Optional ByVal Message As String)
    Call IncrementCount
    Call Err.Raise(AssertCode.SuccessCode, , Message)
End Sub

Public Sub IsTrue(ByVal Condition As Boolean, Optional ByVal Message As String)
    Call That(Condition, Iz.True, Message)
End Sub

Public Sub IsFalse(ByVal Condition As Boolean, Optional ByVal Message As String)
    Call That(Condition, Iz.False, Message)
End Sub

Public Sub AreEqual(ByRef Expected As Variant, ByRef Actual As Variant, Optional ByVal Message As String)
    Call That(Actual, Iz.EqualTo(Expected), Message)
End Sub

Public Sub That(ByRef Actual As Variant, ByVal Constraint As IConstraint, Optional ByVal Message As String)
    Set Constraint = Constraint.Resolve

    If Constraint.Matches(Actual) = False Then
        Dim Writer As New TextMessageWriter
        
        If Len(Message) > 0 Then
            Call Writer.WriteLine(Message)
        End If
        
        Call Constraint.WriteMessageTo(Writer)
        Call Fail(Writer.ToString)
    End If
    
    Call IncrementCount
End Sub
