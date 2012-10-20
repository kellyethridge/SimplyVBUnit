VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PersistableConstraint"
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
' Module: PersistableConstraint
'
Option Explicit
Implements IConstraint

Private Const ERR_CANNOTPERSISTARRAYS   As Long = 328
Private Const ERR_OBJECTNOTPERSISTABLE  As Long = 330


Private mActual As Variant


Public Function Matches(ByRef Actual As Variant) As Boolean
    Dim Result As Boolean
    If IsObject(Actual) Then
        On Error Resume Next
        Dim Bag As New PropertyBag
        Bag.WriteProperty "Item", Actual, Nothing
        
        If Err.Number = 0 Then
            Dim Item As Object
            Set Item = Bag.ReadProperty("Item", Nothing)
            Result = (Not Item Is Nothing)
        End If
    End If
    
    If Not Result Then
        VariantCopyInd mActual, Actual
        
        If Err.Number = ERR_CANNOTPERSISTARRAYS Or Err.Number = ERR_OBJECTNOTPERSISTABLE Then
            Err.Clear
        End If
    End If
    
    Matches = Result
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IConstraint Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IConstraint_Matches(Actual As Variant) As Boolean
    IConstraint_Matches = Matches(Actual)
End Function

Private Sub IConstraint_WriteActualValueTo(ByVal Writer As TextMessageWriter)
    Writer.WriteText "<{0}>", TypeName(mActual)
End Sub

Private Sub IConstraint_WriteDescriptionTo(ByVal Writer As TextMessageWriter)
    Writer.WriteText "PropertyBag persistable"
End Sub

Private Sub IConstraint_WriteMessageTo(ByVal Writer As TextMessageWriter)
    Writer.DisplayConstraintDifferences Me
End Sub
