VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RangeConstraintTests"
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
' Module: RangeConstraintTests
'
Option Explicit
Implements ITestCaseSource

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SVBUnit.TestCaseDataBuilder)
    Select Case Test.MethodName
        Case "TestRange"
            Call Test.Use(1, 5, 1).Expect(True)
            Call Test.Use(1, 5, 0).Expect(False)
            Call Test.Use(1, 5, 5).Expect(True)
            Call Test.Use(1, 5, 6).Expect(False)
            Call Test.Use(1.5, 2.5, 2).Expect(True)
            Call Test.Use(1.5, 2.5, 2.6).Expect(False)
            Call Test.Use(1.5, 2.5, 1.4).Expect(False)

    
    End Select
End Sub


Public Function TestRange(ByRef FromValue As Variant, ByRef ToValue As Variant, ByRef Actual As Variant) As Boolean
    Dim c As IConstraint
    Set c = Iz.InRange(FromValue, ToValue)
    
    TestRange = c.Matches(Actual)
End Function
