VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UIConstructors"
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
'    Module: UIConstructors
'

Option Explicit

Public Function NewTestTreeController(ByVal TreeView As Object, ByVal RootTest As ITest, ByVal Listener As EventCastListener) As TestTreeController
    Set NewTestTreeController = New TestTreeController
    Call NewTestTreeController.Init(TreeView, RootTest, Listener)
End Function

Public Function NewResultsTabController(ByVal Output As UIListBox, ByVal NotRun As Object, ByVal Out As Object, ByVal Log As Object, ByVal Error As Object, ByVal Listener As EventCastListener) As ResultsTabController
    Set NewResultsTabController = New ResultsTabController
    Call NewResultsTabController.Init(Output, NotRun, Out, Log, Error, Listener)
End Function

Public Function NewTestProgressController(ByVal Bar As Object, ByVal Display As Object, ByVal Listener As EventCastListener) As TestProgressController
    Set NewTestProgressController = New TestProgressController
    Call NewTestProgressController.Init(Bar, Display, Listener)
End Function

Public Function NewStatusBarController(ByVal Bar As Object, ByVal Listener As EventCastListener) As StatusBarController
    Set NewStatusBarController = New StatusBarController
    Call NewStatusBarController.Init(Bar, Listener)
End Function

Public Function NewUISetting(ByVal Name As String, ByRef Value As Variant) As UISetting
    Set NewUISetting = New UISetting
    Call NewUISetting.Init(Name, Value)
End Function

Public Function NewTestResultCollector(ByVal Listener As EventCastListener) As TestResultCollector
    Set NewTestResultCollector = New TestResultCollector
    Call NewTestResultCollector.Init(Listener)
End Function

Public Function NewClientInfo(ByVal Info As Object) As ClientInfo
    Set NewClientInfo = New ClientInfo
    Call NewClientInfo.Init(Info)
End Function

Public Function NewCategoryListController(ByVal Tests As ITest, ByVal List As Object, ByVal Enabled As Object, ByVal Exclude As Object, ByVal SelectAll As Object) As CategoryListController
    Set NewCategoryListController = New CategoryListController
    Call NewCategoryListController.Init(Tests, List, Enabled, Exclude, SelectAll)
End Function


