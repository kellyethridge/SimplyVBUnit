VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StartsWithConstraintTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
' Module: StartsWithConstraintTests
'
Option Explicit
Implements ITestCaseSource

Private Sub ITestCaseSource_GetTestCases(ByVal Test As TestCaseCollector)
    Select Case Test.MethodName
        Case "Matches_WithValues_ReturnsExpected"
            Call Test.Use("", "").Expect(True)
            Call Test.Use("a", "a").Expect(True)
            Call Test.Use("abc", "ab").Expect(True)
            Call Test.Use("abc", "AB").Expect(False)
            Call Test.Use("ab", "abc").Expect(False)
        
        Case "Matches_WithValuesIgnoringCase_ReturnsExpected"
            Call Test.Use("", "").Expect(True)
            Call Test.Use("a", "a").Expect(True)
            Call Test.Use("abc", "ab").Expect(True)
            Call Test.Use("abc", "AB").Expect(True)
            Call Test.Use("ab", "abc").Expect(False)
            
    End Select
End Sub


Public Function Matches_WithValues_ReturnsExpected(ByRef Actual As String, ByRef Expected As String) As Boolean
    Dim c As StartsWithConstraint
    Set c = Sim.NewStartsWithConstraint(Expected)
    
    Matches_WithValues_ReturnsExpected = c.Matches(Actual)
End Function

Public Function Matches_WithValuesIgnoringCase_ReturnsExpected(ByRef Actual As String, ByRef Expected As String) As Boolean
    Dim c As StartsWithConstraint
    Set c = Sim.NewStartsWithConstraint(Expected)
    c.IgnoreCase
    
    Matches_WithValuesIgnoringCase_ReturnsExpected = c.Matches(Actual)
End Function
