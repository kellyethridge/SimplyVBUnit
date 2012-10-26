VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestMovieListWithOneMovie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestMovieListWithOneMovie
'
Option Explicit

Private MovieList   As MovieList
Private StarWars    As Movie


Public Sub Setup()
    Set StarWars = NewMovie("Star Wars")
    Set MovieList = New MovieList
    MovieList.Add StarWars
End Sub


Public Sub TestCount()
    Assert.That MovieList.Count, Iz.EqualTo(1), "Count of one item list should be 1."
End Sub

