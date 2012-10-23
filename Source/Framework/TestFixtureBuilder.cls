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
' Module: SComponent.TestFixtureBuilder
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
