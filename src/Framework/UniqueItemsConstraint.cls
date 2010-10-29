VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UniqueItemsConstraint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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
' Module: UniqueConstraint
'
Option Explicit
Implements IConstraint

Private mActual As Variant


Public Function Matches(ByRef Actual As Variant) As Boolean
    If Not IsEnumerable(Actual) Then _
        Err.Raise ErrorCode.Argument, "UniqueItemsConstraint.Matches", "The actual value must be enumerable."
    
    Matches = IsUnique(Actual)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IsUnique(ByRef Actual As Variant) As Boolean
    Call VariantCopyInd(mActual, Actual)
    
    Dim Tallier As New ValueTallier
    Dim Element As Variant
    For Each Element In Actual
        If Tallier.IncrementTally(Element) > 1 Then Exit Function
    Next Element
    
    IsUnique = True
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IConstraint Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IConstraint_Matches(Actual As Variant) As Boolean
    IConstraint_Matches = Matches(Actual)
End Function

Private Sub IConstraint_WriteActualValueTo(ByVal Writer As TextMessageWriter)

End Sub

Private Sub IConstraint_WriteDescriptionTo(ByVal Writer As TextMessageWriter)

End Sub

Private Sub IConstraint_WriteMessageTo(ByVal Writer As TextMessageWriter)

End Sub
