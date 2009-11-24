VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestFilterBase"
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
' Module: TestFilterBase
'
Option Explicit

Private mFilter As ITestFilter


Public Function Pass(ByVal Filter As ITestFilter, ByVal Test As ITest) As Boolean
    On Error GoTo errTrap
    
    Call Init(Filter)
    Pass = DoesPass(Test)
    
errTrap:
    Call Release
End Function

Private Sub Init(ByVal Filter As ITestFilter)
    Set mFilter = Filter
End Sub

Private Sub Release()
    Set mFilter = Nothing
End Sub

Private Function DoesPass(ByVal Test As ITest) As Boolean
    Dim Result As Boolean
    
    If ExactMatch(Test) Then
        Result = True
    ElseIf HasMatchingDescendent(Test) Then
        Result = True
    Else
        Result = HasMatchingParent(Test)
    End If

    DoesPass = Result
End Function

Private Function ExactMatch(ByVal Test As ITest) As Boolean
    ExactMatch = mFilter.Match(Test)
End Function

Private Function HasMatchingDescendent(ByVal Test As ITest) As Boolean
    If HasDescendents(Test) Then
        HasMatchingDescendent = TryMatchingDescendent(Test.Tests)
    End If
End Function

Private Function HasDescendents(ByVal Test As ITest) As Boolean
    HasDescendents = Test.IsSuite
End Function

Private Function TryMatchingDescendent(ByVal Descendents As TestList) As Boolean
    Dim Descendent As ITest
    
    For Each Descendent In Descendents
        If DoesPass(Descendent) Then
            TryMatchingDescendent = True
            Exit Function
        End If
    Next
End Function

Private Function HasMatchingParent(ByVal Test As ITest) As Boolean
    Do While Test.HasParent
        If mFilter.Match(Test.Parent) Then
            HasMatchingParent = True
            Exit Function
        End If
        
        Set Test = Test.Parent
    Loop
End Function

