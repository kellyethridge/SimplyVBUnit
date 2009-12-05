VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IzTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
' Module: IzTests
'
Option Explicit
Implements ITestCaseSource

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SVBUnit.TestCaseDataBuilder)
    Select Case Test.MethodName
        Case "Matches_WithConstraint_ReturnsExpected"
            Call Test.Use(Iz.EqualTo(1), 1, True)
            Call Test.Use(Iz.EqualTo(1), 2, False)
            Call Test.Use(Iz.True, True, True)
            Call Test.Use(Iz.True, False, False)
            Call Test.Use(Iz.False, False, True)
            Call Test.Use(Iz.False, True, False)
            Call Test.Use(Iz.Nothing, Nothing, True)
            Call Test.Use(Iz.Nothing, 1, False)
            Call Test.Use(Iz.vbEmpty, Empty, True)
            Call Test.Use(Iz.vbEmpty, 1, False)
            Call Test.Use(Iz.vbNull, Null, True)
            Call Test.Use(Iz.vbNull, 1, False)
            Call Test.Use(Iz.Missing, GetMissingVariant, True)
            Call Test.Use(Iz.Missing, 1, False)
            
        
    End Select
End Sub

Public Sub Matches_WithConstraint_ReturnsExpected(ByVal c As IConstraint, ByRef ConstraintActual As Variant, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = c.Matches(ConstraintActual)
    
    Assert.IsTrue Actual = Expected
End Sub



Private Function GetMissingVariant(Optional ByVal Value As Variant) As Variant
    GetMissingVariant = Value
End Function