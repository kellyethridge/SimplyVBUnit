VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NumberComparerTests"
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
' Module: NumberComparerTests
'
Option Explicit
Implements ITestCaseSource

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SVBUnit.TestCaseCollector)
    Select Case Test.MethodName
        Case "Compare_WithValues_ReturnsExpected"
            Call Test.Use(1, 1).Expect(0)
            Call Test.Use(1, 2).Expect(Iz.LessThan(0))
            Call Test.Use(2, 1).Expect(Iz.GreaterThan(0))
            Call Test.Use(1#, 2#).Expect(Iz.LessThan(0))
            Call Test.Use(2#, 1#).Expect(Iz.GreaterThan(0))
            Call Test.Use(2#, 2#).Expect(0)
            Call Test.Use(1@, 2@).Expect(Iz.LessThan(0))
            Call Test.Use(2@, 1@).Expect(Iz.GreaterThan(0))
            Call Test.Use(2@, 2@).Expect(0)
            Call Test.Use(CDec(1), CDec(2)).Expect(Iz.LessThan(0))
            Call Test.Use(CDec(2), CDec(1)).Expect(Iz.GreaterThan(0))
            Call Test.Use(CDec(1), CDec(1)).Expect(0)
            Call Test.Use(1.1, 1).Expect(Iz.GreaterThan(0))
            Call Test.Use(0.9, 1).Expect(Iz.LessThan(0))
    End Select
End Sub

Public Function Compare_WithValues_ReturnsExpected(ByRef X As Variant, ByRef Y As Variant) As Long
    Dim c As New NumberComparer
    Compare_WithValues_ReturnsExpected = c.Compare(X, Y)
End Function
