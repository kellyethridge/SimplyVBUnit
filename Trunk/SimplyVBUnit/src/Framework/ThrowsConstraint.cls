VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ThrowsConstraint"
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
' Module: ThrowsConstraint
'
Option Explicit
Implements IConstraint


Private mExpected   As ErrorInfo
Private mActual     As ErrorInfo
Private mIgnoreCase As Boolean


Public Function Matches(ByVal Actual As ErrorInfo) As Boolean
    Set mActual = Actual
    Matches = mActual.Equals(mExpected, mIgnoreCase)
End Function

Public Function IgnoreCase() As ThrowsConstraint
    mIgnoreCase = True
    Set IgnoreCase = Me
End Function

Public Sub WriteMessageTo(ByVal Writer As TextMessageWriter)
    Writer.DisplayConstraintDifferences Me
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Expected As ErrorInfo)
    If Expected Is Nothing Then _
        Error.CannotBeNothing Me, "Init", "Expected"
        
    Set mExpected = Expected
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IConstraint Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IConstraint_Matches(Actual As Variant) As Boolean
    IConstraint_Matches = Matches(Actual)
End Function

Private Function IConstraint_Resolve() As IConstraint
    Set IConstraint_Resolve = Me
End Function

Private Sub IConstraint_WriteActualValueTo(ByVal Writer As TextMessageWriter)
    If mActual Is Nothing Then
        Writer.WriteText "no error raised"
    Else
        Writer.WriteErrorInfo mActual
    End If
End Sub

Private Sub IConstraint_WriteDescriptionTo(ByVal Writer As TextMessageWriter)
    If mActual Is Nothing Then
        Writer.WriteText "an error"
    Else
        Writer.WriteErrorInfo mExpected
    End If
End Sub

Private Sub IConstraint_WriteMessageTo(ByVal Writer As TextMessageWriter)
    WriteMessageTo Writer
End Sub
