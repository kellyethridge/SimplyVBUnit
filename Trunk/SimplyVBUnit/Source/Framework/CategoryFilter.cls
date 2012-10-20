VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CategoryFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Copyright 2011 Kelly Ethridge
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
' Module: CategoryFilter
'
Option Explicit
Implements ITestFilter

Private mBase       As New TestFilterBase
Private mCategories As New Collection


Public Sub Add(ByRef Name As String)
    If Len(Name) = 0 Then _
        Error.CannotBeEmpty Me, "Add", "Name"
        
    mCategories.Add Name
End Sub

Public Function Match(ByVal Test As ITest) As Boolean
    If Test.HasCategories Then
        Dim Category As Variant
        For Each Category In mCategories
            If Test.Categories.Contains(CStr(Category)) Then
                Match = True
                Exit Function
            End If
        Next
    End If
End Function

Public Function Pass(ByVal Test As ITest) As Boolean
    Pass = mBase.Pass(Me, Test)
End Function

Public Property Get Categories() As Collection
    Set Categories = mCategories
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ITestFilter Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ITestFilter_Match(ByVal Test As ITest) As Boolean
    ITestFilter_Match = Match(Test)
End Function

Private Function ITestFilter_Pass(ByVal Test As ITest) As Boolean
    ITestFilter_Pass = Pass(Test)
End Function
