VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestFileWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestFileWriter
'
Option Explicit
Implements ICategorizable

Private Writer As FileWriter


Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "File Access"
End Sub


Public Sub Setup()
    DeleteFile
    Set Writer = New FileWriter
    Writer.OpenFile Filename
End Sub

Public Sub Teardown()
    Writer.CloseFile
End Sub

Public Sub TestCreateEmptyFile()
    Assert.That Writer.IsOpen, Iz.True, "Writer should be open."
    Writer.CloseFile
    Assert.That Writer.IsOpen, Iz.False, "Writer should be closed."
    
    Assert.That Dir$(Filename, vbNormal), Iz.Not.Empty, "TestFile.txt should exist."
    Assert.That FileLen(Filename), Iz.EqualTo(0), "File length should be 0."
End Sub

Public Sub TestWriteText()
    Writer.WriteText "Test data"
    Writer.CloseFile
    
    Dim Text As String
    Text = ReadFile
    
    Assert.That Text, Iz.EqualTo("Test data"), "Text should be 'Test data'."
End Sub

Public Sub TestOpenAnOpenWriterException()
    Assert.Throws 55
    
    Writer.OpenFile Filename
End Sub






Private Function ReadFile() As String
    Dim FF As Long
    FF = FreeFile
    Open Filename For Input As FF
    ReadFile = Input(LOF(FF), #FF)
    Close FF
End Function

Private Property Get Filename() As String
    Filename = App.Path & "\TestFile.txt"
End Property

Private Sub DeleteFile()
    On Error GoTo errTrap
    Kill Filename
errTrap:
End Sub

