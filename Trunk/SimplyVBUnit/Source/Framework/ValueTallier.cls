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
' Module: ValueTallier
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
    
    Dim Result As Boolean
    If Tally.Count > 0 Then
        Tally.Decrement
        Result = True
    End If
    
    TryRemove = Result
End Function

Public Property Get IsEmpty() As Boolean
    Dim Tally As Tally
    For Each Tally In mTallies
        If Tally.Count > 0 Then
            Exit Property
        End If
    Next
    
    IsEmpty = True
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetTally(ByRef Value As Variant) As Tally
    Dim Tally As Tally
    For Each Tally In mTallies
        If Tally.Equals(Value, mComparer) Then
            Set GetTally = Tally
            Exit Function
        End If
    Next
    
    Set GetTally = Sim.NewTally(Value)
    mTallies.Add GetTally
End Function
