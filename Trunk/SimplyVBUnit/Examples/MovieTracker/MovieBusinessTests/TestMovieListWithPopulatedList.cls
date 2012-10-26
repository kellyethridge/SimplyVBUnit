VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestMovieListWithPopulatedList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestMovieListWithPopulatedList
'
Option Explicit

Private MovieList   As MovieList
Private StarWars    As Movie
Private StarTrek    As Movie
Private StarGate    As Movie


Public Sub Setup()
    Set StarWars = NewMovie("Star Wars")
    Set StarTrek = NewMovie("Star Trek")
    Set StarGate = NewMovie("Stargate")
    Set MovieList = New MovieList
    MovieList.Add StarWars
    MovieList.Add StarTrek
    MovieList.Add StarGate
End Sub


Public Sub TestAddingDuplicate()
    Dim Duplicate As Movie
    Set Duplicate = NewMovie(StarWars.Name)
    
    Assert.Throws 5, , "Movie already exists."
    MovieList.Add Duplicate
End Sub

Public Sub TestRenaming()
    Dim NewName As String
    NewName = "StarTrek I"
    MovieList.Rename StarTrek, NewName
    
    Assert.That StarTrek.Name, Iz.EqualTo(NewName), "Name should be different after renaming."
End Sub

Public Sub TestRenamingDuplicate()
    Assert.Throws 5, , "Movie already exists."
    
    On Error GoTo errTrap
    MovieList.Rename StarTrek, "Star Wars"
    Exit Sub
    
errTrap:
    Dim Ex As ErrorInfo
    Set Ex = ErrorInfo.FromErr(Err)
    
    Assert.That MovieList.Count, Iz.EqualTo(3), "Failed rename shouldn't change list count."
    Assert.That StarTrek.Name, Iz.EqualTo("Star Trek"), "Failed rename shouldn't change the name."
    
    Ex.Throw
End Sub

Public Sub TestTwoNotEqualLists()
    Dim Other As New MovieList
    Other.Add StarWars
    Other.Add StarGate
    
    Assert.That MovieList.Equals(Other), Iz.False, "Other list should not be equal."
End Sub

Public Sub TestTwoEqualLists()
    Dim Other As New MovieList
    Other.Add StarWars
    Other.Add StarGate
    Other.Add StarTrek
    
    Assert.That MovieList.Equals(Other), Iz.True, "Other list should be equal."
End Sub
