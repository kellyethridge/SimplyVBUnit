VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CategoryStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
' CategoryStatic
'
Option Explicit

Private mCategories As New Collection


Public Property Get All() As Category
    Set All = mCategories("All")
End Property

Public Property Get Uncategorized() As Category
    Set Uncategorized = mCategories("Uncategorized")
End Property

Public Property Get SciFi() As Category
    Set SciFi = mCategories("Science Fiction")
End Property

Public Property Get Horror() As Category
    Set Horror = mCategories("Horror")
End Property

Public Property Get Comedy() As Category
    Set Comedy = mCategories("Comedy")
End Property

Public Property Get Western() As Category
    Set Western = mCategories("Western")
End Property

Public Property Get Drama() As Category
    Set Drama = mCategories("Drama")
End Property

Public Property Get Fantasy() As Category
    Set Fantasy = mCategories("Fantasy")
End Property

Public Property Get Kids() As Category
    Set Kids = mCategories("Kids")
End Property

Public Property Get Adult() As Category
    Set Adult = mCategories("Adult")
End Property

Public Property Get Mystery() As Category
    Set Mystery = mCategories("Mystery")
End Property

Public Property Get Thriller() As Category
    Set Thriller = mCategories("Thriller")
End Property

Public Function GetCategoryByName(ByVal Name As String) As Category
    On Error GoTo errTrap
    Set GetCategoryByName = mCategories(Name)
    Exit Function
    
errTrap:
    Set GetCategoryByName = Me.Uncategorized
End Function


Private Sub Class_Initialize()
    Call mCategories.Add(MB.NewCategory("All"), "All")
    Call mCategories.Add(MB.NewCategory("Uncategorized"), "Uncategorized")
    Call mCategories.Add(MB.NewCategory("Science Fiction"), "Science Fiction")
    Call mCategories.Add(MB.NewCategory("Horror"), "Horror")
    Call mCategories.Add(MB.NewCategory("Comedy"), "Comedy")
    Call mCategories.Add(MB.NewCategory("Western"), "Western")
    Call mCategories.Add(MB.NewCategory("Drama"), "Drama")
    Call mCategories.Add(MB.NewCategory("Fantasy"), "Fantasy")
    Call mCategories.Add(MB.NewCategory("Kids"), "Kids")
    Call mCategories.Add(MB.NewCategory("Adult"), "Adult")
    Call mCategories.Add(MB.NewCategory("Mystery"), "Mystery")
    Call mCategories.Add(MB.NewCategory("Thriller"), "Thriller")
End Sub
