VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestFileReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestFileReader
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "File Access"
End Sub

Public Sub TestEmptyFile()
    Dim FF As Long
    FF = FreeFile
    Open Filename For Output As #FF
    Close #FF
    
    Dim Reader As New FileReader
    Reader.OpenFile Filename
    
    Assert.That Reader.HasMore, Iz.False, "File shoule not have more data."
End Sub

Public Sub TestOneLineFile()
    Dim Text As String
    Text = "This is line one"
    
    Dim FF As Long
    FF = FreeFile
    Open Filename For Output As #FF
    Print #FF, Text
    Close #FF
    
    Dim Reader As New FileReader
    Reader.OpenFile Filename
    
    Dim Line As String
    Line = Reader.ReadLine
    Assert.That Reader.HasMore, Iz.False, "Reader should not have more text."
    Assert.That Line, Iz.EqualTo(Text), "Line contains wrong text."
    
    Reader.CloseFile
End Sub

Public Sub TestOpenAnOpenReaderException()
    Assert.Throws 55
    
    Dim Reader As New FileReader
    
    Reader.OpenFile Filename
    Reader.OpenFile Filename
End Sub


Private Property Get Filename() As String
    Filename = App.Path & "\TestFile.txt"
End Property

