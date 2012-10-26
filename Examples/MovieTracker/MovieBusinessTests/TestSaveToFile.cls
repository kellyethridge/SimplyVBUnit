VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestSaveToFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestSaveToFile
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "File Access"
End Sub

Public Sub TestWriteEmptyList()
    Dim List As New MovieList
    
    Dim Writer As New FileWriter
    Writer.OpenFile Filename
    List.WriteTo Writer
    Writer.CloseFile
    
    Assert.That FileLen(App.Path & "\MovieData.txt"), Iz.EqualTo(0)
End Sub

Public Sub TestWriteOneMovieList()
    Dim List As New MovieList
    List.Add MB.NewMovie("A Movie")
    
    Dim Writer As New FileWriter
    Writer.OpenFile Filename
    List.WriteTo Writer
    Writer.CloseFile
    
    Open Filename For Input As #1
    Dim s As String
    Line Input #1, s
    Close #1
    
    Assert.That s, Iz.EqualTo("A Movie|Uncategorized|-1")
End Sub




Private Property Get Filename() As String
    Filename = App.Path & "\MovieData.txt"
End Property

