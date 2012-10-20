VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestFixtureBuilder"
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
' Module: TestFixtureBuilder
'
Option Explicit

Private mCollectors         As New Collection
Private mHasSetup           As Boolean
Private mHasFixtureSetup    As Boolean
Private mHasTeardown        As Boolean
Private mHasFixtureTeardown As Boolean



Public Property Get HasSetup() As Boolean
    HasSetup = mHasSetup
End Property

Public Property Let HasSetup(ByVal RHS As Boolean)
    mHasSetup = RHS
End Property

Public Property Get HasFixtureSetup() As Boolean
    HasFixtureSetup = mHasFixtureSetup
End Property

Public Property Let HasFixtureSetup(ByVal RHS As Boolean)
    mHasFixtureSetup = RHS
End Property

Public Property Get HasTeardown() As Boolean
    HasTeardown = mHasTeardown
End Property

Public Property Let HasTeardown(ByVal RHS As Boolean)
    mHasTeardown = RHS
End Property

Public Property Get HasFixtureTeardown() As Boolean
    HasFixtureTeardown = mHasFixtureTeardown
End Property

Public Property Let HasFixtureTeardown(ByVal RHS As Boolean)
    mHasFixtureTeardown = RHS
End Property

Public Function Add(ByRef MethodName As String) As TestCaseCollector
    On Error GoTo errTrap
    
    Set Add = mCollectors(MethodName)
    Exit Function
    
errTrap:
    Set Add = Sim.NewTestCaseCollector(MethodName)
    mCollectors.Add Add, MethodName
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCollectors.[_NewEnum]
End Function
