VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NotConstraint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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
' Module: NotConstraint
'
Option Explicit
Implements IConstraint

Private mBaseConstraint As IConstraint


Public Function Matches(ByRef Actual As Variant) As Boolean
    Matches = Not mBaseConstraint.Matches(Actual)
End Function

Public Sub WriteActualValueTo(ByVal Writer As TextMessageWriter)
    Call mBaseConstraint.WriteActualValueTo(Writer)
End Sub

Public Sub WriteDescriptionTo(ByVal Writer As TextMessageWriter)
    Call Writer.WritePredicate("not")
    Call mBaseConstraint.WriteDescriptionTo(Writer)
End Sub

Public Sub WriteMessageTo(ByVal Writer As TextMessageWriter)
    Call Writer.DisplayConstraintDifference(Me)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal BaseConstraint As IConstraint)
    If BaseConstraint Is Nothing Then _
        Call Error.CannotBeNothing(Me, "Init", "BaseConstraint")
        
    Set mBaseConstraint = BaseConstraint
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IConstraint Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IConstraint_Matches(Actual As Variant) As Boolean
    IConstraint_Matches = Matches(Actual)
End Function

Private Sub IConstraint_WriteActualValueTo(ByVal Writer As TextMessageWriter)
    Call WriteActualValueTo(Writer)
End Sub

Private Sub IConstraint_WriteDescriptionTo(ByVal Writer As TextMessageWriter)
    Call WriteDescriptionTo(Writer)
End Sub

Private Sub IConstraint_WriteMessageTo(ByVal Writer As TextMessageWriter)
    Call WriteMessageTo(Writer)
End Sub
