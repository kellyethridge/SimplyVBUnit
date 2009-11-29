VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CategoryList"
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
' Module: CategoryList
'
Option Explicit

Private mItems As New Collection

    
Public Property Get Count() As Long
    Count = mItems.Count
End Property

Public Sub Add(ByVal Name As String)
    On Error GoTo errTrap
    Call mItems.Add(Name, Name)
errTrap:
End Sub

Public Property Get Item(ByVal Index As Long) As String
Attribute Item.VB_UserMemId = 0
    If Index < 0 Or Index >= Me.Count Then _
        Call Error.IndexOutOfRange(Me, "Item")
'        Call Err.Raise(ErrorCode.IndexOutOfRange, Resource.GetString(Src_CategoryList_Item), Resource.GetString(Arg_IndexOutOfRange))
        
    Item = mItems(Index + 1)
End Property

Public Function Contains(ByVal Name As String) As Boolean
    On Error GoTo errTrap
    Name = mItems(Name)
    Contains = True
errTrap:
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mItems.[_NewEnum]
End Function
