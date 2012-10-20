VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderedConstraintTests"
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
' Module: OrderedConstraintTests
'
Option Explicit
Implements ITestCaseSource

Private Sub ITestCaseSource_GetTestCases(ByVal Test As TestCaseCollector)
    Select Case Test.MethodName
        Case "Matches_WithValues_ReturnsExpected"
            Call Test.Use(Array()).Expect(True)
            Call Test.Use(Array(1)).Expect(True)
            Call Test.Use(Array(1, 2)).Expect(True)
            Call Test.Use(Array(2, 1)).Expect(False)
            Call Test.Use(Array(1, 1, 2, 2)).Expect(True)
            Call Test.Use(Array(1, 1)).Expect(True)
            Call Test.Use(Array("a", "b")).Expect(True)
            Call Test.Use(Array("b", "a")).Expect(False)
            Call Test.Use(Array("1", "a")).Expect(True)
            Call Test.Use(Array("a", "1")).Expect(False)
            Call Test.Use(Array("A", "a")).Expect(True)
            Call Test.Use(Array("a", "A")).Expect(False)
            
        Case "Matches_WithCaseInsensitiveValues_ReturnsExpected"
            Call Test.Use(Array("a", "A", "B", "b")).Expect(True)
            Call Test.Use(Array("A", "B", "a", "b")).Expect(False)
            
        Case "Matches_WithDescendingValues_ReturnsExpected"
            Call Test.Use(Array()).Expect(True)
            Call Test.Use(Array(1, 2)).Expect(False)
            Call Test.Use(Array(2, 1)).Expect(True)
            
    End Select
End Sub


Public Function Matches_WithValues_ReturnsExpected(ByRef Actual As Variant) As Boolean
    Dim c As New OrderedConstraint
    Matches_WithValues_ReturnsExpected = c.Matches(Actual)
End Function

Public Function Matches_WithCaseInsensitiveValues_ReturnsExpected(ByRef Actual As Variant) As Boolean
    Dim c As New OrderedConstraint
    c.IgnoreCase
    Matches_WithCaseInsensitiveValues_ReturnsExpected = c.Matches(Actual)
End Function

Public Function Matches_WithDescendingValues_ReturnsExpected(ByRef Actual As Variant) As Boolean
    Dim c As New OrderedConstraint
    c.Descending
    Matches_WithDescendingValues_ReturnsExpected = c.Matches(Actual)
End Function
