VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'    CopyRight (c) 2008 Kelly Ethridge
'
'    This file is part of SimplyVBUnitUI.
'
'    SimplyVBUnitUI is free software; you can redistribute it and/or modify
'    it under the terms of the GNU Library General Public License as published by
'    the Free Software Foundation; either version 2.1 of the License, or
'    (at your option) any later version.
'
'    SimplyVBUnitUI is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Library General Public License for more details.
'
'    You should have received a copy of the GNU Library General Public License
'    along with Foobar; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'    Module: UserEvents
'

Option Explicit

Public Event AddTest(ByVal Fixture As Object)
Public Event AddListener(ByVal Listener As IEventListener)
Public Event SetFilter(ByVal Filter As ITestFilter)


Public Sub OnSetFilter(ByVal Filter As ITestFilter)
    RaiseEvent SetFilter(Filter)
End Sub

Public Sub OnAddListener(ByVal Listener As IEventListener)
    RaiseEvent AddListener(Listener)
End Sub

Public Sub OnAddTest(ByVal Fixture As Object)
    RaiseEvent AddTest(Fixture)
End Sub
