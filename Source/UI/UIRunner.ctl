VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.UserControl UIRunner 
   ClientHeight    =   5130
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9060
   ScaleHeight     =   5130
   ScaleWidth      =   9060
   Begin VB.PictureBox picSplitter 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   3480
      MousePointer    =   9  'Size W E
      ScaleHeight     =   4695
      ScaleWidth      =   135
      TabIndex        =   7
      Top             =   0
      Width           =   135
   End
   Begin VB.PictureBox picRightPanel 
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   3720
      ScaleHeight     =   4695
      ScaleWidth      =   5295
      TabIndex        =   6
      Top             =   0
      Width           =   5295
      Begin SimplyVBUnitUI.UIListBox lstFailureOutput 
         Height          =   2415
         Left            =   120
         TabIndex        =   15
         Top             =   2130
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   4260
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtErrorsOutput 
         Height          =   2415
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   14
         Top             =   2130
         Width           =   4935
      End
      Begin VB.TextBox txtConsoleOutput 
         Height          =   2415
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   10
         Top             =   2130
         Width           =   4935
      End
      Begin VB.TextBox txtLogOutput 
         Height          =   2415
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   9
         Top             =   2130
         Width           =   4935
      End
      Begin MSComctlLib.TreeView tvwTestsNotRun 
         Height          =   2415
         Left            =   120
         TabIndex        =   8
         Top             =   2130
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   4260
         _Version        =   393217
         Indentation     =   529
         LineStyle       =   1
         Style           =   6
         Appearance      =   1
      End
      Begin MSComctlLib.TabStrip tabOutputs 
         Height          =   2895
         Left            =   0
         TabIndex        =   3
         Top             =   1770
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   5106
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   5
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Failures && Errors"
               Key             =   "Failures"
               Object.ToolTipText     =   "Displays failures and errors generated from assertions"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Tests Not Run"
               Key             =   "Ignored"
               Object.ToolTipText     =   "List of tests that were not run"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Console Out"
               Key             =   "Console"
               Object.ToolTipText     =   "Displays text that is output by the user using TestContext.Out"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Log"
               Key             =   "Log"
               Object.ToolTipText     =   "Displays logging text "
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Errors"
               Key             =   "Errors"
               Object.ToolTipText     =   "Displays error text"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.Frame framTestRunnerControls 
         Height          =   1695
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   5175
         Begin VB.CommandButton cmdRun 
            Caption         =   "Run"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   0
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdStop 
            Caption         =   "Stop"
            Height          =   375
            Left            =   1320
            TabIndex        =   1
            Top             =   240
            Width           =   1095
         End
         Begin MSComctlLib.ProgressBar barTestProgress 
            Height          =   375
            Left            =   120
            TabIndex        =   12
            Top             =   1200
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   661
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
            Scrolling       =   1
         End
         Begin VB.Label lblCurrentTest 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   450
            Left            =   120
            TabIndex        =   13
            Top             =   720
            Width           =   4935
         End
      End
   End
   Begin VB.PictureBox picLeftPanel 
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   120
      ScaleHeight     =   4575
      ScaleWidth      =   3255
      TabIndex        =   5
      Top             =   120
      Width           =   3255
      Begin MSComctlLib.TreeView tvwTests 
         Height          =   4575
         Left            =   0
         TabIndex        =   2
         Tag             =   "skip"
         Top             =   0
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   8070
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   441
         LabelEdit       =   1
         LineStyle       =   1
         PathSeparator   =   "."
         Sorted          =   -1  'True
         Style           =   7
         ImageList       =   "imglTestTree"
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imglTestTree 
      Left            =   600
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   12
      ImageHeight     =   12
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UIRunner.ctx":0000
            Key             =   "Inconclusive"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UIRunner.ctx":0057
            Key             =   "Ignored"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UIRunner.ctx":00AE
            Key             =   "Passed"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UIRunner.ctx":0105
            Key             =   "NotRun"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UIRunner.ctx":015C
            Key             =   "Failed"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar statProgress 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   4755
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   5662
            Key             =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "TestCases"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "TestsRun"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "Failures"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "Time"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuRun 
         Caption         =   "Run"
      End
      Begin VB.Menu mnuRunAll 
         Caption         =   "Run All"
      End
      Begin VB.Menu Split2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExpandAllNodesPopUp 
         Caption         =   "Expand All Nodes"
      End
      Begin VB.Menu mnuCollapseAllNodesPopUp 
         Caption         =   "Collapse All Nodes"
      End
      Begin VB.Menu mnuCollapseToTopLevelPopUp 
         Caption         =   "Collapse To Top Level"
      End
      Begin VB.Menu Split3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuResultDetailsPopUp 
         Caption         =   "Result Details"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      NegotiatePosition=   1  'Left
      Begin VB.Menu mnuExpandAllNodes 
         Caption         =   "E&xpand All Nodes"
      End
      Begin VB.Menu mnuCollapseAllNodes 
         Caption         =   "&Collapse All Nodes"
      End
      Begin VB.Menu mnuCollapseToTopLevel 
         Caption         =   "Collapse To &Top Level"
      End
      Begin VB.Menu Split1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options"
      End
   End
End
Attribute VB_Name = "UIRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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
'    Module: ctlSimplyVBUnitUI
'

Option Explicit

'Private WithEvents mUserEvents As UserEvents
Private WithEvents mContainer  As Form
Attribute mContainer.VB_VarHelpID = -1

Private mConfig                 As New UIConfiguration
Private mAnchor                 As Anchor
Private mLeftPanelContent       As Anchor
Private mRightPanelContent      As Anchor
Private mMouseDownDX            As Long
Private mDragSplitter           As Boolean
Private mSplitterLeftMargin     As Long
Private mSplitterRightMargin    As Long

Private mTests          As TestSuite
Private mListeners      As New MultiCastListener
Private mFilter         As ITestFilter
Private mRunner         As TestRunner
Private mListener       As New EventCastListener
Private mTestTree       As TestTreeController
Private mResultsTab     As ResultsTabController
Private mProgress       As TestProgressController
Private mStatus         As StatusBarController
Private mResults        As TestResultCollector


Public Property Get Width() As Single
    Width = UserControl.Width
End Property

Public Property Get Height() As Single
    Height = UserControl.Height
End Property

Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    Call PropertyChanged("Font")
End Property

Public Property Get SplitterPosition() As Long
    SplitterPosition = picSplitter.Left
End Property

Public Property Let SplitterPosition(ByVal RHS As Long)
    picSplitter.Left = RHS
    Set mAnchor = Nothing
    Call PositionControls
End Property

Public Sub AddListener(ByVal Listener As IEventListener)
    Call mListeners.Add(Listener)
End Sub

Public Sub AddTest(ByVal Fixture As Object)
    mTests.Add Fixture
    Call mStatus.Reset(mTests.TestCount)
End Sub

Public Sub SetFilter(ByVal Filter As ITestFilter)
    Set mFilter = Filter
End Sub

Public Sub Init(ByVal Info As Object)
    Set ClientInfo = UI.NewClientInfo(Info)
    mSplitterLeftMargin = 205
    mSplitterRightMargin = 210

    Set mTests = Sim.NewTestSuite(ClientInfo.EXEName)
    
    Dim Item As Object
    For Each Item In modStaticClasses.Tests
        Call mTests.Add(Item)
    Next

    Call InitControllers
    Call DisplayTabPages
    Call InitApp
    Call InitTitle
    Call mStatus.Reset(mTests.TestCount)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitControllers()
    Set mTestTree = UI.NewTestTreeController(UserControl.tvwTests, mTests, mListener)
    Set mResultsTab = UI.NewResultsTabController(UserControl.lstFailureOutput, UserControl.tvwTestsNotRun, UserControl.txtConsoleOutput, UserControl.txtLogOutput, UserControl.txtErrorsOutput, mListener)
    Set mProgress = UI.NewTestProgressController(UserControl.barTestProgress, UserControl.lblCurrentTest, mListener)
    Set mStatus = UI.NewStatusBarController(UserControl.statProgress, mListener)
    Set mResults = UI.NewTestResultCollector(mListener)
End Sub

Private Sub InitApp()
    Set mContainer = UserControl.Parent
    Call mConfig.Load(ClientInfo.Path & "\" & ClientInfo.EXEName & ".config")
    Call RestoreFormConfiguration
    Call mTestTree.RestoreTreeViewState(mConfig)
    Call InitContextWriters
    Call Me.AddListener(mListener)
End Sub

Private Sub InitTitle()
    Dim sb As New StringBuilder
    Call sb.AppendFormat("SimplyVBUnit {0}.{1} - [{2}]", App.Major, App.Minor, ClientInfo.EXEName)
    mContainer.Caption = sb.ToString
End Sub

Private Sub PositionControls()
    If Not mContainer Is Nothing Then
        If mContainer.WindowState = vbMinimized Then
            Exit Sub
        End If
    End If
    
    If mLeftPanelContent Is Nothing Then
        Set mLeftPanelContent = New Anchor
        Call mLeftPanelContent.Add(UserControl.tvwTests, ToAllSides)
    End If
    Call mLeftPanelContent.ReAnchor
    
    If mRightPanelContent Is Nothing Then
        Set mRightPanelContent = New Anchor
        Call mRightPanelContent.Add(UserControl.lstFailureOutput, ToAllSides)
        Call mRightPanelContent.Add(UserControl.tabOutputs, ToAllSides)
        Call mRightPanelContent.Add(UserControl.txtConsoleOutput, ToAllSides)
        Call mRightPanelContent.Add(UserControl.txtLogOutput, ToAllSides)
        Call mRightPanelContent.Add(UserControl.txtErrorsOutput, ToAllSides)
        Call mRightPanelContent.Add(UserControl.tvwTestsNotRun, ToAllSides)
        Call mRightPanelContent.Add(UserControl.framTestRunnerControls, ToLeft Or ToRight)
        Call mRightPanelContent.Add(UserControl.barTestProgress, ToLeft Or ToRight)
        Call mRightPanelContent.Add(UserControl.lblCurrentTest, ToLeft Or ToRight)
    End If
    Call mRightPanelContent.ReAnchor

    If mAnchor Is Nothing Then
        picLeftPanel.Width = picSplitter.Left - mSplitterLeftMargin
        Dim NewLeft As Long
        Dim NewWidth As Long

        NewLeft = picSplitter.Left + mSplitterRightMargin
        NewWidth = UserControl.Width - picSplitter.Left - mSplitterRightMargin

        Call picRightPanel.Move(NewLeft, picRightPanel.Top, NewWidth, picRightPanel.Height)
        
        Set mAnchor = New Anchor
        Call mAnchor.Add(picSplitter, ToTop Or ToBottom)
        Call mAnchor.Add(picRightPanel, ToTop Or ToBottom Or ToRight Or ToLeft)
        Call mAnchor.Add(picLeftPanel, ToTop Or ToBottom)
        Call PositionControls
    Else
        Call mAnchor.ReAnchor
    End If
End Sub

Private Sub DisplayTabPages()
    UserControl.lstFailureOutput.Visible = UserControl.tabOutputs.Tabs("Failures").Selected
    UserControl.tvwTestsNotRun.Visible = UserControl.tabOutputs.Tabs("Ignored").Selected
    UserControl.txtConsoleOutput.Visible = UserControl.tabOutputs.Tabs("Console").Selected
    UserControl.txtLogOutput.Visible = UserControl.tabOutputs.Tabs("Log").Selected
    UserControl.txtErrorsOutput.Visible = UserControl.tabOutputs.Tabs("Errors").Selected
End Sub

Private Sub InitContextWriters()
    Set TestContext.Out = Sim.NewTestOutputWriter(mListener, TestOutputType.StandardOutput)
    Set TestContext.Log = Sim.NewTestOutputWriter(mListener, TestOutputType.LogOutput)
    Set TestContext.Error = Sim.NewTestOutputWriter(mListener, TestOutputType.ErrorOutput)
End Sub

Private Sub SaveFormConfiguration()
    Dim Settings As New Collection
    Call Settings.Add(UI.NewUISetting("WindowState", mContainer.WindowState))
    Call Settings.Add(UI.NewUISetting("Left", mContainer.Left))
    Call Settings.Add(UI.NewUISetting("Top", mContainer.Top))
    Call Settings.Add(UI.NewUISetting("Width", mContainer.Width))
    Call Settings.Add(UI.NewUISetting("Height", mContainer.Height))
    Call Settings.Add(UI.NewUISetting("SplitterPosition", Me.SplitterPosition))
    
    Call mConfig.SetSettings("Form", Settings)
End Sub

Private Sub RestoreFormConfiguration()
    Dim Settings As Collection
    Set Settings = mConfig.GetSettings("Form")
    
    If Settings.Count > 0 Then
        Dim Left    As Long: Left = Settings("Left").Value
        Dim Top     As Long: Top = Settings("Top").Value
        Dim Width   As Long: Width = Settings("Width").Value
        Dim Height  As Long: Height = Settings("Height").Value

        Call mContainer.Move(Left, Top, Width, Height)
        mContainer.WindowState = Settings("WindowState").Value
        Me.SplitterPosition = Settings("SplitterPosition").Value
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Control Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdRun_Click()
' TODO: Understand this
'    Call mTests.Reset
    
    Dim StartingTest As ITest
    Set StartingTest = mTestTree.SelectedTest
    
    If StartingTest Is Nothing Then
        Set StartingTest = mTests
    End If
    
    Call mResultsTab.SetOutputSupport(mConfig)
    Set mRunner = Sim.NewTestRunner(StartingTest)
    mRunner.AllowCancel = mConfig.AllowStop
    cmdStop.Enabled = mConfig.AllowStop
    Call mRunner.Run(mListeners, mFilter)
    cmdStop.Enabled = True
End Sub

Private Sub cmdStop_Click()
    Call mRunner.Cancel
End Sub

Private Sub mContainer_Resize()
    Call PositionControls
    Call UserControl.Extender.Move(0, 0, mContainer.ScaleWidth, mContainer.ScaleHeight)
End Sub

Private Sub mnuCollapseAllNodes_Click()
    Call mTestTree.CollapseAllNodes
End Sub

Private Sub mnuCollapseAllNodesPopUp_Click()
    Call mTestTree.CollapseAllNodes
End Sub

Private Sub mnuCollapseToTopLevel_Click()
    Call mTestTree.CollapseToTopLevel
End Sub

Private Sub mnuCollapseToTopLevelPopUp_Click()
    Call mTestTree.CollapseToTopLevel
End Sub

Private Sub mnuExpandAllNodes_Click()
    Call mTestTree.ExpandAllNodes
End Sub

Private Sub mnuExpandAllNodesPopUp_Click()
    Call mTestTree.ExpandAllNodes
End Sub

Private Sub mnuOptions_Click()
    Dim Editor As New frmOptions
    Call Editor.Edit(mConfig, Me)
    Call Unload(Editor)
End Sub

Private Sub mnuResultDetailsPopUp_Click()
    Call frmResultDetails.ShowResult(UserControl.tvwTests, mResults, Me)
End Sub

Private Sub mnuRun_Click()
    Call cmdRun_Click
End Sub

Private Sub mnuRunAll_Click()
    Call mTestTree.SelectRoot
    Call cmdRun_Click
End Sub

Private Sub picLeftPanel_Resize()
    Call PositionControls
End Sub

Private Sub picSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        mMouseDownDX = x
        picSplitter.BackColor = &H8080FF
        mDragSplitter = True
    End If
End Sub

Private Sub picSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If mDragSplitter Then
        Dim NewX As Long
        NewX = picSplitter.Left + x - mMouseDownDX
        If NewX < 1000 Then NewX = 1000
        If NewX > UserControl.Width - 1000 Then NewX = UserControl.Width - 1000
        picSplitter.Left = NewX
    End If
End Sub

Private Sub picSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If mDragSplitter Then
        Set mAnchor = Nothing
        picSplitter.BackColor = vbButtonFace
        mDragSplitter = False
        Call PositionControls
        Call picSplitter.Refresh
    End If
End Sub

Private Sub picSplitter_Paint()
    picSplitter.Cls
    
    Dim y As Long
    y = picSplitter.Height / 2 - 200
    
    Dim i As Long
    For i = y To y + 270 Step 90
        picSplitter.CurrentX = -70
        picSplitter.CurrentY = i
        picSplitter.Print "w"
        picSplitter.CurrentX = 20
        picSplitter.CurrentY = i
        picSplitter.Print "8"
    Next i
End Sub

Private Sub picSplitter_Resize()
    Call picSplitter.Refresh
End Sub

Private Sub tabOutputs_Click()
    Call DisplayTabPages
End Sub

Private Sub tvwTests_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        Call PopupMenu(mnuPopUp)
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   UserControl Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UserControl_Hide()
    Set mAnchor = Nothing
    Set mLeftPanelContent = Nothing
    Set mRightPanelContent = Nothing
    
    If Ambient.UserMode Then
        Call SaveFormConfiguration
        Call mTestTree.SaveTreeViewState(mConfig)
        Call mConfig.Save
    End If
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
End Sub

Private Sub UserControl_Resize()
    If Not mContainer Is Nothing Then
        If mContainer.WindowState = vbMinimized Then
            Exit Sub
        End If
    End If

    If UserControl.Width - picSplitter.Left < 1000 Then
        picSplitter.Left = UserControl.Width - 1000
        Set mAnchor = Nothing
    End If
    Call PositionControls
End Sub

Private Sub UserControl_Show()
    If Ambient.UserMode Then
        Call mTestTree.Refresh
        Call mTestTree.RestoreTreeViewState(mConfig)
    End If
    
    Set mAnchor = Nothing
    Call PositionControls
        
    If Ambient.UserMode Then
        If mConfig.AutoRun Then
            Call cmdRun_Click
        End If
    End If
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
End Sub
