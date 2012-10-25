VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SimplyVBUnit Options"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtUpdateFrequency 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2760
      TabIndex        =   9
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
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
      Left            =   4080
      TabIndex        =   8
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
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
      Left            =   5280
      TabIndex        =   7
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CheckBox chkAutoRunTests 
      Caption         =   "Autorun tests on start"
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
      Left            =   480
      TabIndex        =   5
      Top             =   1080
      Width           =   2655
   End
   Begin VB.CheckBox chkOutputToLogConsole 
      Caption         =   "Output to log console"
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
      Left            =   480
      TabIndex        =   4
      Top             =   2040
      Width           =   3615
   End
   Begin VB.CheckBox chkOutputToErrorConsole 
      Caption         =   "Output to error console"
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
      Left            =   480
      TabIndex        =   3
      Top             =   2400
      Width           =   3615
   End
   Begin VB.CheckBox chkOutputToTextConsole 
      Caption         =   "Output to text console"
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
      Left            =   480
      TabIndex        =   2
      Top             =   1680
      Width           =   3615
   End
   Begin VB.ComboBox cboTreeViewStates 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmOptions.frx":0000
      Left            =   3960
      List            =   "frmOptions.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "tests completed."
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
      Left            =   3600
      TabIndex        =   11
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Update display after every"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   10
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Height          =   45
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   6255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Test TreeView state when starting test:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   3360
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
' Module: SComponent.frmOptions
'
Option Explicit

Private mOK         As Boolean
Private mOptions    As UIConfiguration

Public Function Edit(ByVal Options As UIConfiguration, ByVal Owner As Object) As Boolean
    Set mOptions = Options
    Call DisplayOptions
    Call Me.Show(vbModal, Owner)

    Edit = mOK
End Function

Private Sub DisplayOptions()
    Me.txtUpdateFrequency.Text = mOptions.DoEventsFrequency
    Me.cboTreeViewStates.Text = mOptions.TreeViewStartUpState

    Call SetChecked(Me.chkAutoRunTests, mOptions.AutoRun)
    Call SetChecked(Me.chkOutputToErrorConsole, mOptions.OutputToErrorConsole)
    Call SetChecked(Me.chkOutputToLogConsole, mOptions.OutputToLogConsole)
    Call SetChecked(Me.chkOutputToTextConsole, mOptions.OutputToTextConsole)
End Sub

Private Sub cmdCancel_Click()
    Call Unload(Me)
End Sub

Private Sub cmdOK_Click()
    mOptions.DoEventsFrequency = Val(Me.txtUpdateFrequency.Text)
    mOptions.TreeViewStartUpState = Me.cboTreeViewStates.Text
    mOptions.AutoRun = GetChecked(Me.chkAutoRunTests)
    mOptions.OutputToErrorConsole = GetChecked(Me.chkOutputToErrorConsole)
    mOptions.OutputToLogConsole = GetChecked(Me.chkOutputToLogConsole)
    mOptions.OutputToTextConsole = GetChecked(Me.chkOutputToTextConsole)
    
    mOK = True
    Call Unload(Me)
End Sub

Private Function ToCheckMark(ByVal Condition As Boolean) As CheckBoxConstants
    If Condition Then
        ToCheckMark = vbChecked
    Else
        ToCheckMark = vbUnchecked
    End If
End Function

Private Function GetChecked(ByVal CheckBox As CheckBox) As Boolean
    GetChecked = (CheckBox.Value = vbChecked)
End Function

Private Sub SetChecked(ByVal CheckBox As CheckBox, ByVal Checked As Boolean)
    If Checked Then
        CheckBox.Value = vbChecked
    Else
        CheckBox.Value = vbUnchecked
    End If
End Sub
