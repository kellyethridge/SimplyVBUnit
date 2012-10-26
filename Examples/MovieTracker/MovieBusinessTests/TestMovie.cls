VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestMovie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private StarWars As Movie


Public Sub Setup()
    Set StarWars = NewMovie("Star Wars")
End Sub


Public Sub TestMovieName()
    Assert.That StarWars.Name, Iz.EqualTo("Star Wars")
End Sub

Public Sub TestEmptyNameFails()
    Assert.Throws 5, , "Movie name cannot be empty."
    
    Dim Movie As Movie
    Set Movie = NewMovie("")
End Sub

Public Sub TestRenaming()
    Dim NewName As String
    NewName = "Star Trek"
    
    StarWars.Name = NewName
    Assert.That StarWars.Name, Iz.EqualTo(NewName), "Renaming should change the name."
End Sub

Public Sub TestEmptyRename()
    Assert.Throws 5, , "Movie name cannot be empty."
    
    StarWars.Name = ""
End Sub

Public Sub TestClone()
    Dim CloneOfStarWars As Movie
    Set CloneOfStarWars = StarWars.Clone
    
    Assert.That CloneOfStarWars, Iz.Not.SameAs(StarWars), "Clone should not be the same as the original."
    Assert.That CloneOfStarWars, Iz.EqualTo(StarWars).Using(New MovieComparer), "Clone should be equal to the original."
End Sub

Public Sub TestUnrated()
    Assert.That StarWars.HasRating, Iz.False, "StarWars should be unrated."
End Sub

Public Sub TestRatedMovie()
    Dim Fotr As Movie
    Set Fotr = NewMovie("Fellowship of the Ring", , 5)
    
    Assert.That Fotr.HasRating, Iz.True, "fotr should be rated."
    Assert.That Fotr.Rating, Iz.EqualTo(5), "fotr should be rated 5."
End Sub

Public Sub TestUnratedException()
    Assert.Throws 5, StarWars.Name, "Movie is not rated."
    
    Dim Rating As Long
    Rating = StarWars.Rating
End Sub

Public Sub TestUncategorized()
    Assert.That StarWars.Category, Iz.EqualTo(Category.Uncategorized), "StarWars should be uncategorized."
End Sub

Public Sub TestScienceFiction()
    Dim Alien As Movie
    Set Alien = MB.NewMovie("Alien", Category.SciFi, -1)
    
    Assert.That Alien.Category, Iz.EqualTo(Category.SciFi), "Alien should be science fiction."
End Sub
