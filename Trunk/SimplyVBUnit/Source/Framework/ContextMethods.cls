VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContextMethods"
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
' Module: ContextMethods
'
Option Explicit

Private mSetup          As ITestMethod
Private mTeardown       As ITestMethod
Private mSkipSetup      As Boolean
Private mSkipTeardown   As Boolean


Public Sub RunSetup()
    RunMethod mSetup, mSkipSetup
End Sub

Public Sub RunTeardown()
    RunMethod mTeardown, mSkipTeardown
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal SetupMethod As ITestMethod, ByVal TeardownMethod As ITestMethod)
    Set mSetup = SetupMethod
    Set mTeardown = TeardownMethod
End Sub


Private Sub RunMethod(ByRef Method As ITestMethod, ByRef Skip As Boolean)
    On Error GoTo errTrap
    If Skip = False Then
        Method.Run
    End If
    Exit Sub
    
errTrap:
    If Err.Number = 438 Then
        Skip = True
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub
