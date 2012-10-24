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
'The MIT License (MIT)
'Copyright (c) 2012 Kelly Ethridge
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights to
'use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of
'the Software, and to permit persons to whom the Software is furnished to do so,
'subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
'INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
'PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
'FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
'OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
'DEALINGS IN THE SOFTWARE.
'
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
    VariantCopyInd mActual, Actual
    
    If mLeftSide.Matches(Actual) Then
        Matches = mRightSide.Matches(Actual)
    End If
End Function

Public Sub WriteActualValueTo(ByVal Writer As TextMessageWriter)
    Writer.WriteActualValue mActual
End Sub

Public Sub WriteDescriptionTo(ByVal Writer As TextMessageWriter)
    mLeftSide.WriteDescriptionTo Writer
    Writer.WriteConnector "and"
    mRightSide.WriteDescriptionTo Writer
End Sub

Public Sub WriteMessageTo(ByVal Writer As TextMessageWriter)
    Writer.DisplayConstraintDifferences Me
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal LeftSide As IConstraint, ByVal RightSide As IConstraint)
    If LeftSide Is Nothing Then _
        Error.CannotBeNothing Me, "Init", "LeftSide"
    If RightSide Is Nothing Then _
        Error.CannotBeNothing Me, "Init", "RightSide"
        
    Set mLeftSide = LeftSide
    Set mRightSide = RightSide
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IConstraint Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IConstraint_Matches(Actual As Variant) As Boolean
    IConstraint_Matches = Matches(Actual)
End Function

Private Sub IConstraint_WriteActualValueTo(ByVal Writer As TextMessageWriter)
    WriteActualValueTo Writer
End Sub

Private Sub IConstraint_WriteDescriptionTo(ByVal Writer As TextMessageWriter)
    WriteDescriptionTo Writer
End Sub

Private Sub IConstraint_WriteMessageTo(ByVal Writer As TextMessageWriter)
    WriteMessageTo Writer
End Sub
