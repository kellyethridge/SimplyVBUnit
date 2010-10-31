VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ValueTallier"
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

Private mComparer   As EqualityComparer
Private mTallies    As New Collection


Public Sub Initialize(ByRef Values As Variant, ByVal Comparer As EqualityComparer)
    If Not IsEnumerable(Values) Then _
        Error.EnumerationNotSupported Me, "Initialize", "Source"
    If Comparer Is Nothing Then _
        Error.CannotBeNothing Me, "Initialize", "Comparer"
    
    Set mComparer = Comparer
    
    Dim Value As Variant
    Dim Tally As Tally
    
    For Each Value In Values
        Set Tally = GetTally(Value)
        Tally.Increment
    Next
End Sub

Public Function TryRemove(ByRef Value As Variant) As Boolean
    Dim Tally As Tally
    Set Tally = GetTally(Value)
    
    If Tally.Count = 0 Then
        Exit Function
    End If
    
    Tally.Decrement
    TryRemove = True
End Function

'Public Property Get Count() As Long
'    Dim Tally As Tally
'    Dim Result As Long
'
'    For Each Tally In mTallies
'        Result = Result + Tally.Count
'    Next
'
'    Count = Result
'End Property

Public Property Get IsEmpty() As Boolean
    Dim Tally As Tally
    For Each Tally In mTallies
        If Tally.Count > 0 Then
            Exit Property
        End If
    Next
    
    IsEmpty = True
End Property


Private Function GetTally(ByRef Value As Variant) As Tally
    Dim Tally As Tally
    For Each Tally In mTallies
        If Tally.Equals(Value, mComparer) Then
            Set GetTally = Tally
            Exit Function
        End If
    Next
    
    Set GetTally = New Tally
    GetTally.Init Value
    mTallies.Add GetTally
End Function
