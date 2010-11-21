VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RegexConstraintTests"
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
' Module: RegexConstraintTests
'
Option Explicit
Implements ITestCaseSource

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SVBUnit.TestCaseCollector)
    Select Case Test.MethodName
        Case "Matches_WithValues_ReturnsExpected"
            Call Test.Use("kelly", "kelly").Expect(True)
            Call Test.Use("KELLY", "kelly").Expect(False)
            Call Test.Use("kelly", "abckellyadsf").Expect(True)
            Call Test.Use("kelly", "abckellads").Expect(False)
            Call Test.Use("^ke", "kflly").Expect(False)
            Call Test.Use("^ke", "kelly").Expect(True)
            Call Test.Use("^ke", "abdkellyddsa").Expect(False)
            Call Test.Use("1234", 1234).Expect(False)
            
        Case "Matches_WithValuesIgnoringCase_ReturnsExpected"
            Call Test.Use("kelly", "kelly").Expect(True)
            Call Test.Use("KELLY", "kelly").Expect(True)
            Call Test.Use("kelly", "abckellyadsf").Expect(True)
            Call Test.Use("kelly", "abckellads").Expect(False)
            Call Test.Use("^ke", "kflly").Expect(False)
            Call Test.Use("^ke", "kelly").Expect(True)
            Call Test.Use("^ke", "abdkellyddsa").Expect(False)
            Call Test.Use("1234", 1234).Expect(False)
        
    End Select
End Sub


Public Function Matches_WithValues_ReturnsExpected(ByRef Pattern As String, ByRef Actual As Variant) As Boolean
    Dim c As RegexConstraint
    Set c = Sim.NewRegexConstraint(Pattern)
    
    Matches_WithValues_ReturnsExpected = c.Matches(Actual)
End Function

Public Function Matches_WithValuesIgnoringCase_ReturnsExpected(ByRef Pattern As String, ByRef Actual As Variant) As Boolean
    Dim c As RegexConstraint
    Set c = Iz.StringMatching(Pattern)
    c.IgnoreCase
    
    Matches_WithValuesIgnoringCase_ReturnsExpected = c.Matches(Actual)
End Function
