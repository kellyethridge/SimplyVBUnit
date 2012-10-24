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
' Module: SComponent.UIConstructors
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


