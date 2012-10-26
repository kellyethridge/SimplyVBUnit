VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestMovieListCategorySublist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestMovieListCategorySublist
'
Option Explicit

Private MovieList       As MovieList
Private FantasyList     As MovieList
Private SciFiList       As MovieList
Private HorrorList      As MovieList
Private ThrillerList    As MovieList

Private StarWars        As Movie
Private StarTrek        As Movie
Private StarGate        As Movie
Private TheShining      As Movie
Private Carrie          As Movie
Private Fotr            As Movie
Private RedOctober      As Movie
Private Congo           As Movie
Private PrincessBride   As Movie


Public Sub Setup()
    Set StarWars = MB.NewMovie("Star Wars", Category.SciFi, 5)
    Set StarTrek = MB.NewMovie("Star Trek", Category.SciFi, 3)
    Set StarGate = MB.NewMovie("Stargate", Category.SciFi, -1)
    Set TheShining = MB.NewMovie("The Shining", Category.Horror, 2)
    Set Carrie = MB.NewMovie("Carrie", Category.Horror, 3)
    Set Fotr = MB.NewMovie("The Fellowship of The Ring", Category.Fantasy, 5)
    Set RedOctober = MB.NewMovie("The Hunt for Red October", Category.Thriller, 3)
    Set Congo = MB.NewMovie("Congo", Category.Thriller, 3)
    Set PrincessBride = MB.NewMovie("The Princess Bride", Category.Fantasy, 5)
    
    Set MovieList = New MovieList
    MovieList.Add StarWars
    MovieList.Add StarTrek
    MovieList.Add StarGate
    MovieList.Add TheShining
    MovieList.Add Carrie
    MovieList.Add Fotr
    MovieList.Add RedOctober
    MovieList.Add Congo
    MovieList.Add PrincessBride
    
    Set FantasyList = New MovieList
    FantasyList.Add Fotr
    FantasyList.Add PrincessBride
    
    Set SciFiList = New MovieList
    SciFiList.Add StarWars
    SciFiList.Add StarTrek
    SciFiList.Add StarGate
    
    Set HorrorList = New MovieList
    HorrorList.Add TheShining
    HorrorList.Add Carrie
    
    Set ThrillerList = New MovieList
    ThrillerList.Add RedOctober
    ThrillerList.Add Congo
End Sub


Public Sub TestSubsets()
    Assert.That FantasyList, Iz.EqualTo(MovieList.CategorySublist(Category.Fantasy)).Using(New MovieListComparer), "wrong Fantasy sublist."
    Assert.That SciFiList, Iz.EqualTo(MovieList.CategorySublist(Category.SciFi)).Using(New MovieListComparer), "wrong SciFi sublist."
    Assert.That HorrorList, Iz.EqualTo(MovieList.CategorySublist(Category.Horror)).Using(New MovieListComparer), "wrong Horror sublist."
    Assert.That ThrillerList, Iz.EqualTo(MovieList.CategorySublist(Category.Thriller)).Using(New MovieListComparer), "wrong Thriller sublist."
End Sub

Public Sub TestAllSubset()
    Assert.That MovieList, Iz.EqualTo(MovieList.CategorySublist(Category.All)).Using(New MovieListComparer), "Should be all movies."
End Sub
