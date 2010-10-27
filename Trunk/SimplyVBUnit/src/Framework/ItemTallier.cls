VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ItemTallier"
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
' Module: ItemTallier
'
Option Explicit

Private mTallies As New Collection


Public Function GetTally(ByRef Value As Variant) As Long
    Tally = GetTally(Value).Count
End Function

Public Sub ClearTally(ByRef Value As Variant)
    GetTally(Value).Clear
End Sub

Public Function IncrementTally(ByRef Value As Variant) As Long
    IncrementTally = GetTally(Value).Increment
End Function

Public Function DecrementTally(ByRef Value As Variant) As Long
    DecrementTally = GetTally(Value).Decrement
End Function

Public Property Get TotalTallyCount() As Long
    Dim Result  As Long
    Dim t       As Tally
    
    For Each t In mTallies
        Result = Result + t.Count
    Next t
    
    TotalTallyCount = Result
End Property

Public Function Remove(ByRef Value As Variant) As Boolean
    With GetTally(Value)
        If .Count > 0 Then
            .Decrement
            Remove = True
        End If
    End With
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef Values As Variant)
    If Not IsEnumerable(Values) Then _
        Err.Raise ErrorCode.Argument, "ItemTallier", "An enumerable object is required."
        
    Dim Element As Variant
    For Each Element In Values
        If (Not IsArray(Element)) And (VarType(Element) <> vbUserDefinedType) Then
            IncrementTally Element
        End If
    Next Element
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetTally(ByRef Value As Variant) As Tally
    Dim Key As String
    Key = CreateKey(Value)
    
    If Len(Key) > 0 Then
        On Error GoTo errTrap
        Set GetTally = mTallies(Key)
    Else
        Set GetTally = NullTally.Instance
    End If
    
    Exit Function
    
errTrap:
    Set GetTally = New Tally
    mTallies.Add GetTally, Key
End Function

Private Function CreateKey(ByRef Value As Variant) As String
    Dim Result As String
    
    Select Case VarType(Value)
        Case vbByte:        Result = "I1_" & Value
        Case vbInteger:     Result = "I2_" & Value
        Case vbLong:        Result = "I4_" & Value
        Case vbSingle:      Result = "R4_" & Value
        Case vbDouble:      Result = "R8_" & Value
        Case vbDecimal:     Result = "Dec_" & Value
        Case vbCurrency:    Result = "Cy_" & Value
        Case vbBoolean:     Result = "Bool_" & Value
        Case vbDate:        Result = "Date_" & Value
        Case vbString:      Result = "Str_" & Value
        Case vbObject:      Result = "Obj_ " & ObjPtr(Value)
        Case vbNull:        Result = "Null"
        Case vbEmpty:       Result = "Empty"
    End Select
    
    CreateKey = Result
End Function


