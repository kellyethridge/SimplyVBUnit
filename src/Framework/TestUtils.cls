VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
' Module: TestUtils
'
Option Explicit

Friend Sub ClearParents(ByVal List As TestList)
    Dim Test As ITest
    
    If List Is Nothing Then
        Exit Sub
    End If
    
    For Each Test In List
        Set Test.Parent = Nothing
    Next
End Sub

Public Sub SetParents(ByVal List As TestList, ByVal Parent As ITest)
    Dim Test As ITest
    
    For Each Test In List
        Set Test.Parent = Parent
    Next
End Sub

Friend Function GetObjectFromPointer(ByVal Ptr As Long) As Object
    If Ptr <> vbNullPtr Then
        Dim Unk As IUnknown
        
        ObjectPtr(Unk) = Ptr
        Set GetObjectFromPointer = Unk
        ObjectPtr(Unk) = vbNullPtr
    End If
End Function

Friend Function BuildFullName(ByVal Test As ITest) As String
    Dim Parent As ITest
    Dim Prefix As String
    
    Set Parent = Test.Parent
    If Not Parent Is Nothing Then
        Prefix = Parent.FullName & "."
    End If
    
    BuildFullName = Prefix & Test.Name
End Function

