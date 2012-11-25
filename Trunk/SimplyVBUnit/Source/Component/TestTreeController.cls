VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestTreeController"
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
' Module: SComponent.TestTreeController
'
Option Explicit

Private Const IMG_PASSED        As String = "Passed"
Private Const IMG_FAILED        As String = "Failed"
Private Const IMG_NOTRUN        As String = "NotRun"
Private Const IMG_IGNORED       As String = "Ignored"
Private Const IMG_INCONCLUSIVE  As String = "Inconclusive"

Private WithEvents mListener As EventCastListener
Attribute mListener.VB_VarHelpID = -1
Private mTreeView   As TreeView
Private mRootKey    As String
Private mRootNode   As Nodes


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Public Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddTest(ByVal Test As ITest)
    Dim Parent As Node
    Set Parent = mTreeView.Nodes(mRootKey)

    Call AddTestNode(Parent, Test)
    Parent.Sorted = True
End Sub

Public Property Get SelectedTest() As ITest
    If Not mTreeView.SelectedItem Is Nothing Then
        Set SelectedTest = mTreeView.SelectedItem.Tag
    End If
End Property

Public Sub ExpandAllNodes()
    Call SetNodeExpanded(True)
End Sub

Public Sub CollapseAllNodes()
    Call SetNodeExpanded(False)
End Sub

Public Sub CollapseToTopLevel()
    Call CollapseAllNodes
    mTreeView.Nodes(mRootKey).Expanded = True
End Sub

Public Sub RestoreTreeViewState(ByVal Config As UIConfiguration)
    Select Case Config.TreeViewStartUpState
        Case "Auto"
            Dim Settings As Collection
            Set Settings = Config.GetSettings("TreeViewExpanded")

            On Error Resume Next
            Dim Setting As UISetting
            For Each Setting In Settings
                mTreeView.Nodes(Setting.Name).Expanded = Setting.Value
            Next Setting

        Case "Collapsed"
            Call CollapseAllNodes

        Case "Expanded"
            Call ExpandAllNodes

        Case "Hide Tests"
            Dim Node As Node
            For Each Node In mTreeView.Nodes
                Dim Test As ITest
                Set Test = Node.Tag
                Node.Expanded = Test.IsSuite
            Next Node
    End Select

    On Error GoTo errTrap
    mTreeView.Nodes(Config.GetValue("SelectedTest", mRootKey)).Selected = True
errTrap:
    mTreeView.Visible = True
End Sub

Public Sub SaveTreeViewState(ByVal Config As UIConfiguration)
    Dim Settings    As New Collection
    Dim Node        As Node
    
    For Each Node In mTreeView.Nodes
        If Node.Children > 0 Then
            Call Settings.Add(UI.NewUISetting(Node.Key, Node.Expanded))
        End If
    Next Node
    
    Call Config.SetSettings("TreeViewExpanded", Settings)
    
    If Not mTreeView.SelectedItem Is Nothing Then
        Call Config.SetValue("SelectedTest", mTreeView.SelectedItem.Key)
    Else
        Call Config.SetValue("SelectedTest", mRootKey)
    End If
End Sub

Public Sub Refresh()
    Dim Root As ITest
    Set Root = mTreeView.Nodes(mRootKey).Tag
    Call mTreeView.Nodes.Clear
    Call AddRootNode(Root)
End Sub

Public Sub SelectRoot()
    mTreeView.Nodes(mRootKey).Selected = True
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal TreeView As TreeView, ByVal Root As ITest, ByVal Listener As EventCastListener)
    Set mTreeView = TreeView
    Set mListener = Listener
    mTreeView.Visible = False
    Set mRootNode = TreeView.Nodes
    Call AddRootNode(Root)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetNodeExpanded(ByVal Expanded As Boolean)
    Dim Node As Node
    For Each Node In mTreeView.Nodes
        Node.Expanded = Expanded
    Next Node
End Sub

Private Sub HideTests()
    Dim Node As Node
    For Each Node In mTreeView.Nodes
        Dim Test As ITest
        Set Test = Node.Tag
        Node.Expanded = Test.IsSuite
    Next Node
End Sub

Private Sub AddTestNode(ByVal Parent As Node, ByVal Test As ITest)
    Dim NewNode As Node
    Set NewNode = mRootNode.Add(Parent, tvwChild, Test.FullName, Test.DisplayName, IMG_NOTRUN)
    Set NewNode.Tag = Test

    Call AddSubTests(NewNode, Test)
    NewNode.Sorted = True
End Sub

Private Sub AddRootNode(ByVal Root As ITest)
    mRootKey = Root.FullName

    Dim RootNode As Node
    Set RootNode = mRootNode.Add(, , mRootKey, ClientInfo.Path & "\" & Root.DisplayName, IMG_NOTRUN)
    Set RootNode.Tag = Root
    
    Call AddSubTests(RootNode, Root)
End Sub

Private Sub AddSubTests(ByVal ParentNode As Node, ByVal ParentTest As ITest)
    If Not ParentTest.Tests Is Nothing Then
        Dim ChildTest As ITest
        For Each ChildTest In ParentTest.Tests
            Call AddTestNode(ParentNode, ChildTest)
        Next ChildTest
    End If
End Sub

Private Sub SetImageUpTree(ByVal Node As Node, ByVal Image As String)
    Do While CanReplaceImage(Node, Image)
        Node.Image = Image
        Set Node = Node.Parent
    Loop
End Sub

Private Function CanReplaceImage(ByVal Node As Node, ByVal NewImage As String) As Boolean
    Dim Result As Boolean
    
    If Not Node Is Nothing Then
        If Node.Image <> NewImage Then
            Select Case Node.Image
                Case IMG_FAILED:    Result = False
                Case IMG_IGNORED:   Result = (NewImage = IMG_FAILED)
                Case Else:          Result = True
            End Select
        End If
    End If
    
    CanReplaceImage = Result
End Function

Private Sub UpdateTreeNode(ByVal Result As TestResult)
    Dim Node As Node
    Set Node = mTreeView.Nodes(Result.Test.FullName)
    
    If Result.IsSuccess Then
        Call SetImageUpTree(Node, IMG_PASSED)
    ElseIf Result.IsFailure Or Result.IsError Then
        Call SetImageUpTree(Node, IMG_FAILED)
    ElseIf Result.IsIgnored Then
        Call SetImageUpTree(Node, IMG_IGNORED)
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   mListener Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mListener_RunStarted(ByVal Name As String, ByVal TestCount As Long)
    Dim Node As Node
    For Each Node In mTreeView.Nodes
        Node.Image = IMG_NOTRUN
    Next Node
End Sub

Private Sub mListener_TestCaseFinished(ByVal Result As TestResult)
    UpdateTreeNode Result
End Sub
