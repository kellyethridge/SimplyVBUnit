VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Movie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
' Movie
'
Option Explicit

Private mName       As String
Private mRating     As Long
Private mCategory   As Category


' Gets the name of the movie.
'
Public Property Get Name() As String
    Name = mName
End Property

' Sets the name of the movie.
'
Public Property Let Name(ByVal RHS As String)
    Call CheckEmpty(RHS)
    mName = RHS
End Property

' Returns a clone of the movie object.
'
Public Function Clone() As Movie
    Set Clone = MB.newMovie(mName, mCategory, mRating)
End Function

' Checks another movie to this movie for equality.
'
Public Function Equals(ByVal Other As Movie) As Boolean
    Equals = (Me.Name = Other.Name)
End Function

' Gets if the movies has a rating or not.
'
Public Property Get HasRating() As Boolean
    HasRating = (mRating > NO_RATING)
End Property

' Gets the movie rating or -1 if the movie is not rated yet.
'
Public Property Get Rating() As Long
    If Not Me.HasRating Then _
        Call Err.Raise(5, Me.Name, "Movie is not rated.")
        
    Rating = mRating
End Property

' Sets the movie rating.
'
Public Property Let Rating(ByVal RHS As Long)
    mRating = RHS
End Property

' Gets the movie category.
'
Public Property Get Category() As Category
    Set Category = mCategory
End Property

' Sets the movie category.
'
Public Property Set Category(ByVal RHS As Category)
    Set mCategory = RHS
End Property

' Checks if this movie is of the specific category.
'
Public Function IsOfCategory(ByVal Category As Category) As Boolean
    IsOfCategory = (mCategory Is Category)
End Function

' Writes the movie to a data source.
'
Public Sub WriteTo(ByVal Writer As FileWriter)
    Call Writer.WriteText(mName)
    Call Writer.WriteText(RECORD_DELIMETER)
    Call Writer.WriteText(mCategory.Name)
    Call Writer.WriteText(RECORD_DELIMETER)
    Call Writer.WriteText(mRating)
    Call Writer.WriteText(vbCrLf)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Name As String, ByVal Category As Category, ByVal Rating As Long)
    Call CheckEmpty(Name)
    
    mName = Name
    mRating = Rating
    Set mCategory = IIf(Not Category Is Nothing, Category, modStaticClasses.Category.Uncategorized)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CheckEmpty(ByVal NewName As String)
    If Len(NewName) = 0 Then
        Call Err.Raise(5, , "Movie name cannot be empty.")
    End If
End Sub

