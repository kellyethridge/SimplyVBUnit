VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestMovieListWithTwoMovies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestMovieListWithTwoMovies
'
Option Explicit

Private MovieList   As MovieList
Private StarWars    As Movie
Private StarTrek    As Movie
Private Stargate    As Movie



Public Sub Setup()
    Set StarWars = NewMovie("Star Wars")
    Set StarTrek = NewMovie("Star Trek")
    Set Stargate = NewMovie("Stargate")
    Set MovieList = New MovieList
    MovieList.Add StarWars
    MovieList.Add StarTrek
End Sub


Public Sub TestCount()
    Assert.That MovieList.Count, Iz.EqualTo(2), "Count of two item list should be 2."
End Sub

Public Sub TestContents()
    Assert.That MovieList.Contains(StarWars), Iz.True, "List should contain Star Wars."
    Assert.That MovieList.Contains(StarTrek), Iz.True, "List should contain Star Trek."
    Assert.That MovieList.Contains(Stargate), Iz.False, "List should not contain Stargate."
End Sub

