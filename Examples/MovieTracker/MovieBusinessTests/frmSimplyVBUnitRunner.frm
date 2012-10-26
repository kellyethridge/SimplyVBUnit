VERSION 5.00
Object = "{7983BD3B-752A-43EA-9BFF-444BBA1FC293}#2.1#0"; "SimplyVBUnit.Component.ocx"
Begin VB.Form frmSimplyVBRunner 
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9585
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   9585
   Begin SimplyVBComp.UIRunner UIRunner1 
      Height          =   5175
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   9135
      _ExtentX        =   16113
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
Attribute VB_Name = "frmSimplyVBRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' frmSimplyVBUnit3Runner
'
' ** NOTE **
' Please set <Tools\Options\General\Error Trapping> to <Break on Unhandled Errors>
'
' ** HELP - Please view the SimplyVBUnit3.HLP file **
'
Option Explicit
' Namespaces Available:
'       Assert.*            ie. Assert.That Value, Iz.EqualTo(5)
'
' Public Functions Availabe:
'       AddTest <TestObject>
'       AddListener <ITestListener Object>
'       AddFilter <ITestFilter Object>
'       WriteText "Message"
'       WriteLine "Message"
'
' Adding a test fixture:
'   Use AddTest <object>
'
' Steps to create a TestCase:
'
'   1. Add a new class
'   2. Name it as desired
'   3. (Optionally) Add a Setup/Teardown method to be run before and after every test.
'   4. (Optionally) Add a FixtureSetup/FixtureTeardown method to be run at the
'      before the first test and after the last test.
'   5. Add public Subs of the tests you want run.
'
'      Public Sub MyTest()
'          Assert.That a, Iz.EqualTo(b)
'      End Sub
'
Private Sub Form_Load()
    ' Add tests here
    '
    ' AddTest <TestObject>
    '
    
    AddTest New TestMovie
    AddTest New TestEmptyMovieList
    AddTest New TestMovieListWithOneMovie
    AddTest New TestMovieListWithTwoMovies
    AddTest New TestMovieListWithPopulatedList
    AddTest New TestMovieListCategorySublist
    AddTest New TestFileWriter
    AddTest New TestFileReader
    AddTest New TestSaveToFile
    AddTest New TestStringTokenizer
    
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Don't change anything below
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Initialize()
    Call Me.UIRunner1.Init(App)
End Sub

' This enables the escape key to exit the tests for rapid testing.
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Call Unload(Me)
End Sub
