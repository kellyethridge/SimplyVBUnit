VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestListEnumerator"
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
' Module: TestListEnumerator
'
Option Explicit
Implements IEnumerator

Private mList   As TestList
Private mIndex  As Long

Public Function MoveNext() As Boolean
    mIndex = mIndex + 1
    MoveNext = (mIndex < mList.Count)
End Function

Public Property Get Current() As ITest
    Set Current = mList(mIndex)
End Property

Public Sub Reset()
    mIndex = -1
End Sub

Friend Sub Init(ByVal List As TestList)
    Set mList = List
    Reset
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IEnumerator Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get IEnumerator_Current() As Variant
    Set IEnumerator_Current = Current
End Property

Private Function IEnumerator_MoveNext() As Boolean
    IEnumerator_MoveNext = MoveNext
End Function

Private Sub IEnumerator_Reset()
    Reset
End Sub
