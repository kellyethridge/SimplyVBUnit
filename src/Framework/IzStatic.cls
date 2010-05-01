VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IzStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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
' Module: IzStatic
'
Option Explicit
Implements IzSyntaxHelper

Private Function IzSyntaxHelper_EqualTo(Expected As Variant) As SimplyVBUnitType.IEqualConstraint
    Dim Expression As IConstraintExpression
    Set Expression = New ConstraintExpression
    Set IzSyntaxHelper_EqualTo = Expression.EqualTo(Expected)
End Function

Private Function IzSyntaxHelper_False() As SimplyVBUnitType.IBasicExpression
    Set IzSyntaxHelper_False = Sim.NewEqualConstraint(False)
End Function

Private Function IzSyntaxHelper_Missing() As SimplyVBUnitType.IBasicExpression
    Set IzSyntaxHelper_Missing = Sim.NewEqualConstraint(GetMissingVariant)
End Function

Private Function IzSyntaxHelper_Not() As SimplyVBUnitType.IConstraintExpression
    
End Function

Private Function IzSyntaxHelper_Nothing() As SimplyVBUnitType.IBasicExpression
    Set IzSyntaxHelper_Nothing = Sim.NewEqualConstraint(Nothing)
End Function

Private Function IzSyntaxHelper_True() As SimplyVBUnitType.IBasicExpression
    Set IzSyntaxHelper_True = Sim.NewEqualConstraint(True)
End Function

Private Function IzSyntaxHelper_VBEmpty() As SimplyVBUnitType.IBasicExpression
    Set IzSyntaxHelper_VBEmpty = Sim.NewEqualConstraint(Empty)
End Function

Private Function IzSyntaxHelper_VBNull() As SimplyVBUnitType.IBasicExpression
    Set IzSyntaxHelper_VBNull = Sim.NewEqualConstraint(Null)
End Function
