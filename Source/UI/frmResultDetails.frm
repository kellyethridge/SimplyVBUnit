VERSION 5.00
Begin VB.Form frmResultDetails 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Test Result"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblType 
      AutoSize        =   -1  'True
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
      Left            =   2040
      TabIndex        =   11
      Top             =   2520
      Width           =   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Type:"
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
      Index           =   5
      Left            =   480
      TabIndex        =   10
      Top             =   2520
      Width           =   480
   End
   Begin VB.Label lblResult 
      AutoSize        =   -1  'True
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
      Left            =   2040
      TabIndex        =   9
      Top             =   1080
      Width           =   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Result:"
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
      Index           =   4
      Left            =   480
      TabIndex        =   8
      Top             =   1080
      Width           =   555
   End
   Begin VB.Label lblMessage 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2040
      TabIndex        =   7
      Top             =   3000
      Width           =   105
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Message:"
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
      Index           =   3
      Left            =   480
      TabIndex        =   6
      Top             =   3000
      Width           =   750
   End
   Begin VB.Label lblAssertCount 
      AutoSize        =   -1  'True
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
      Left            =   2040
      TabIndex        =   5
      Top             =   2040
      Width           =   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Assert Count:"
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
      Index           =   2
      Left            =   480
      TabIndex        =   4
      Top             =   2040
      Width           =   1125
   End
   Begin VB.Label lblExecutionTime 
      AutoSize        =   -1  'True
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
      Left            =   2040
      TabIndex        =   3
      Top             =   1560
      Width           =   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Execution Time:"
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
      Index           =   1
      Left            =   480
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
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
      Left            =   2040
      TabIndex        =   1
      Top             =   600
      Width           =   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Full Name:"
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
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   840
   End
End
Attribute VB_Name = "frmResultDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'    CopyRight (c) 2009 Kelly Ethridge
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
'    Module: frmResultDetails
'

Option Explicit

Private WithEvents mTreeView As TreeView
Attribute mTreeView.VB_VarHelpID = -1
Private mResults As TestResultCollector


Public Sub ShowResult(ByVal View As Object, ByVal Results As TestResultCollector, ByVal Owner As Object)
    Set mTreeView = View
    Set mResults = Results
    
    If Not mTreeView.SelectedItem Is Nothing Then
        Call DisplayResults(mTreeView.SelectedItem.Key)
    End If
    
    If Not Me.Visible Then
        Call Me.Show(vbModeless, Owner)
    End If
End Sub

Private Sub DisplayResults(ByVal FullName As String)
    Dim Result As TestResult
    Set Result = mResults(FullName)
    
    Me.lblName.Caption = FullName
    
    If Not Result Is Nothing Then
        Me.lblExecutionTime.Caption = Result.Time
        Me.lblAssertCount.Caption = Result.AssertCount
        Me.lblMessage.Caption = Result.Message
        Me.lblResult.Caption = GetStatusText(Result)
        Me.lblType.Caption = GetTypeText(Result)
    Else
        Me.lblExecutionTime.Caption = ""
        Me.lblAssertCount.Caption = ""
        Me.lblMessage.Caption = ""
        Me.lblResult.Caption = "Not Run"
        Me.lblType.Caption = ""
    End If
End Sub

Private Function GetStatusText(ByVal Source As TestResult) As String
    Dim Result As String
    
    If Source.IsSuccess Then
        Result = "Success"
    ElseIf Source.IsFailure Then
        Result = "Failure"
    ElseIf Source.IsError Then
        Result = "Error"
    Else
        Result = "Ignored"
    End If
    
    GetStatusText = Result
End Function

Private Function GetTypeText(ByVal Source As TestResult) As String
    Dim Result As String
    
    If Source.Test.IsSuite Then
        Result = "Test Suite"
    Else
        Result = "Test Case"
    End If
    
    GetTypeText = Result
End Function

Private Sub mTreeView_NodeClick(ByVal Node As MSComctlLib.Node)
    Call DisplayResults(Node.Key)
End Sub
