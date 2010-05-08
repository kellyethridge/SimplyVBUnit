VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AndConstraint"
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
' Module: AndConstraint
'
Option Explicit
Implements IConstraint

Private mActual     As Variant
Private mLeftSide   As IConstraint
Private mRightSide  As IConstraint


Public Property Get LeftSide() As IConstraint
    Set LeftSide = mLeftSide
End Property

Public Property Get RightSide() As IConstraint
    Set RightSide = mRightSide
End Property

Public Function Matches(ByRef Actual As Variant) As Boolean
    Call VariantCopyInd(mActual, Actual)
    
    If mLeftSide.Matches(Actual) Then
        Matches = mRightSide.Matches(Actual)
    End If
End Function

Public Sub WriteActualValueTo(ByVal Writer As TextMessageWriter)
    Call Writer.WriteActualValue(mActual)
End Sub

Public Sub WriteDescriptionTo(ByVal Writer As TextMessageWriter)
    Call mLeftSide.WriteDescriptionTo(Writer)
    Call Writer.WriteConnector("and")
    Call mRightSide.WriteDescriptionTo(Writer)
End Sub

Public Sub WriteMessageTo(ByVal Writer As TextMessageWriter)
    Call Writer.DisplayConstraintDifferences(Me)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal LeftSide As IConstraint, ByVal RightSide As IConstraint)
    If LeftSide Is Nothing Then _
        Call Error.CannotBeNothing(Me, "Init", "LeftSide")
    If RightSide Is Nothing Then _
        Call Error.CannotBeNothing(Me, "Init", "RightSide")
        
    Set mLeftSide = LeftSide
    Set mRightSide = RightSide
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IConstraint Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IConstraint_Matches(Actual As Variant) As Boolean
    IConstraint_Matches = Matches(Actual)
End Function

Private Function IConstraint_Resolve() As IConstraint

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
