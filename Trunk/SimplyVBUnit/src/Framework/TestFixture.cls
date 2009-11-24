VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestFixture"
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
' Module: TestFixture
'
Option Explicit


Public Property Get RunState() As RunState: End Property

Public Property Let RunState(ByVal RHS As RunState): End Property

Public Sub Sort(Optional ByVal Comparer As ITestComparer): End Sub

Public Property Get Parent() As ITest: End Property

Public Property Set Parent(ByVal RHS As ITest): End Property

Public Property Get TestCount() As Long: End Property

Public Function Run(Optional ByVal Listener As IEventListener, Optional ByVal Filter As ITestFilter) As TestResult: End Function

Public Property Get Name() As String: End Property

Public Property Get FullName() As String: End Property

Public Property Get Tests() As TestList: End Property

Public Property Get IsSuite() As Boolean: End Property

Public Property Get HasParent() As Boolean: End Property

Public Property Get TestClass() As Object: End Property

