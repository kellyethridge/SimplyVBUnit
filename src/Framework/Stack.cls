VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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
' Module: Stack
'
Option Explicit

Private Const TOP_ITEM_INDEX As Long = 1


Private mItems As Collection


Public Property Get Count() As Long
    Count = mItems.Count
End Property

Public Sub Push(ByRef Value As Variant)
    If Me.IsEmpty Then
        Call mItems.Add(Value)
    Else
        Call mItems.Add(Value, before:=TOP_ITEM_INDEX)
    End If
End Sub

Public Property Get IsEmpty() As Boolean
    IsEmpty = (Me.Count = 0)
End Property
    
Public Function Pop() As Variant
    Call VariantCopyInd(Pop, mItems(TOP_ITEM_INDEX))
    Call mItems.Remove(TOP_ITEM_INDEX)
End Function

Public Sub Clear()
    Call Initialize
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Initialize()
    Set mItems = New Collection
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Call Initialize
End Sub
