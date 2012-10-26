VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MovieList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
' MovieList
'
Option Explicit

Private mMovies As New Collection


' Gets the number of movies in the list.
'
Public Property Get Count() As Long
    Count = mMovies.Count
End Property

' Adds a new movie to the list.
'
Public Sub Add(ByVal MovieToAdd As Movie)
    If Me.Contains(MovieToAdd) Then _
        Call Err.Raise(5, , "Movie already exists.")
        
    Call mMovies.Add(MovieToAdd)
End Sub

' Checks if a movie already exists in the list.
'
Public Function Contains(ByVal MovieToCheckFor As Movie) As Boolean
    Dim Movie As Movie
    
    For Each Movie In mMovies
        If Movie.Equals(MovieToCheckFor) Then
            Contains = True
            Exit Function
        End If
    Next Movie
End Function

' Renames a movie already in the list.
'
Public Sub Rename(ByVal Movie As Movie, ByVal NewName As String)
    Dim PotentialMovie As Movie
    Set PotentialMovie = Movie.Clone
    PotentialMovie.Name = NewName
    
    If Me.Contains(PotentialMovie) Then _
        Call Err.Raise(5, , "Movie already exists.")
        
    Movie.Name = NewName
End Sub

' Compares another list to this list for equality.
'
Public Function Equals(ByVal Other As MovieList) As Boolean
    If Me.Count = Other.Count Then
        Dim Movie As Movie
        
        For Each Movie In mMovies
            If Not Other.Contains(Movie) Then
                Exit Function
            End If
        Next Movie
        
        Equals = True
    End If
End Function

' Returns a sublist of movies in a specific category.
'
Public Function CategorySublist(ByVal aCategory As Category) As MovieList
    If aCategory.Name = Category.All.Name Then
        Set CategorySublist = Me
        Exit Function
    End If
    
    Dim FilteredList    As New MovieList
    Dim Movie           As Movie
    
    For Each Movie In mMovies
        If Movie.IsOfCategory(aCategory) Then
            Call FilteredList.Add(Movie)
        End If
    Next Movie
    
    Set CategorySublist = FilteredList
End Function

' Writes the list of movies to the data source.
'
Public Sub WriteTo(ByVal Writer As FileWriter)
    Dim Movie As Movie
    
    For Each Movie In mMovies
        Call Movie.WriteTo(Writer)
    Next Movie
End Sub

' Reads a list of movies from a data source.
'
Public Sub ReadFrom(ByVal Reader As FileReader)
    Do While Reader.HasMore
        Dim newMovie As Movie
        Set newMovie = Movie.ReadFrom(Reader)
        
        If Not newMovie Is Nothing Then
            Call mMovies.Add(newMovie)
        End If
    Loop
End Sub

