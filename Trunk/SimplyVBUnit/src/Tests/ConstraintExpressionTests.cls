VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConstraintExpressionTests"
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
' Module: ConstraintExpressionTests
'
Option Explicit
Implements ITestCaseSource

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SVBUnit.TestCaseDataBuilder)
    Select Case Test.MethodName
        Case "Resolve_WithExpression_ReturnsConstraint"
            Call Test.Use(Exp.EqualTo(1), 1, True)
            Call Test.Use(Exp.EqualTo(1).OrElse.EqualTo(2), 1, True)
            Call Test.Use(Exp.EqualTo(1).OrElse.EqualTo(2), 2, True)
            Call Test.Use(Exp.EqualTo(1).OrElse.EqualTo(2), 3, False)
            Call Test.Use(Exp.[Not].EqualTo(1), 1, False)
            Call Test.Use(Exp.[Not].EqualTo(1), 2, True)
            Call Test.Use(Exp.[Not].EqualTo(1).AndAlso.[Not].EqualTo(2), 1, False)
            Call Test.Use(Exp.[Not].EqualTo(1).AndAlso.[Not].EqualTo(2), 2, False)
            Call Test.Use(Exp.[Not].EqualTo(1).AndAlso.[Not].EqualTo(2), 3, True)
            
        
    End Select
End Sub

Public Sub Resolve_WithExpression_ReturnsConstraint(ByVal Expression As IResolvable, ByRef ConstraintActual As Variant, ByVal Expected As Boolean)
    Dim Constraint As IConstraint
    
    Set Constraint = Expression.Resolve
    
    Dim Actual As Boolean
    Actual = Constraint.Matches(ConstraintActual)
    assert.IsTrue Actual = Expected
End Sub

Private Function Exp() As IConstraintExpression
    Set Exp = New ConstraintExpression
End Function
