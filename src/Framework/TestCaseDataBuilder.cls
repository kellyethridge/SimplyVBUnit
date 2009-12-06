VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestCaseDataBuilder"
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
' Module: TestCaseDataBuilder
'
Option Explicit

Private mDataStore  As Collection
Private mMethodName As String


Public Property Get MethodName() As String
    MethodName = mMethodName
End Property

Public Function Use(ParamArray Args() As Variant) As TestCaseModifier
    Dim Data As New TestCaseData
    
    If HasElements(UBound(Args)) Then
        Dim Locals() As Variant
        Locals = Args
        Call Data.CopyArguments(Locals)
        Call mDataStore.Add(Data)
    End If
    
    Set Use = Sim.NewTestCaseModifier(Data)
End Function

Private Function HasElements(ByVal UpperBound As Long) As Boolean
    HasElements = (UpperBound >= 0)
End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal MethodName As String, ByVal DataStore As Collection)
    If Len(MethodName) = 0 Then _
        Call Error.CannotBeEmpty(Me, "Init", "MethodName")
    If DataStore Is Nothing Then _
        Call Error.CannotBeNothing(Me, "Init", "DataStore")
        
    mMethodName = MethodName
    Set mDataStore = DataStore
End Sub
