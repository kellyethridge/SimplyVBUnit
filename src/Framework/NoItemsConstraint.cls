VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NoItemsConstraint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
' Module: NoItemsConstraint
'
Option Explicit
Implements IConstraint

Private mConstraint As IConstraint
Private mActual     As Variant


''
' Applies the constraint to a collection or array.
'
' @param Actual The collection or array to apply the constraint against.
' @return Returns true if value passes the constraint, false otherwise.
'
Public Function Matches(ByRef Actual As Variant) As Boolean
    Dim Result As Boolean
    If IsEnumerable(Actual) Then
        Result = True
        Dim Item As Variant
        For Each Item In Actual
            If mConstraint.Matches(Item) Then
                Result = False
                Exit For
            End If
        Next Item
    End If
    
    If Not Result Then
        VariantCopyInd mActual, Actual
    End If
    
    Matches = Result
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Constraint As IConstraint)
    If Constraint Is Nothing Then _
        Error.CannotBeNothing Me, "Init", "Constraint"
    
    Set mConstraint = Constraint
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IConstraint Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IConstraint_Matches(Actual As Variant) As Boolean
    IConstraint_Matches = Matches(Actual)
End Function

Private Sub IConstraint_WriteActualValueTo(ByVal Writer As TextMessageWriter)
    Writer.WriteValue mActual
End Sub

Private Sub IConstraint_WriteDescriptionTo(ByVal Writer As TextMessageWriter)
    Writer.WritePredicate "no items of"
    mConstraint.WriteDescriptionTo Writer
End Sub

Private Sub IConstraint_WriteMessageTo(ByVal Writer As TextMessageWriter)
    Writer.DisplayConstraintDifferences Me
End Sub
