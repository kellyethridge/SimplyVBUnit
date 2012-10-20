VERSION 5.00
Object = "*\A..\SimplyVBUnitUI.vbp"
Begin VB.Form frmSample 
   Caption         =   "Form1"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9420
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6555
   ScaleWidth      =   9420
   StartUpPosition =   3  'Windows Default
   Begin SimplyVBUnitUI.UIRunner UIRunner1 
      Height          =   5175
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   9128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Initialize()
    Call Me.UIRunner1.Init(App)
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Call Unload(Me)
End Sub

Private Sub Form_Load()
    AddTest New FixtureWithTests
    AddTest New FixtureWithTestCases
End Sub
