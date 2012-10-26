VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MBConstructors"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
' MBConstructors
'
Option Explicit


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Public Constructor Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NewMovie(ByVal Name As String, Optional ByVal Category As Category, Optional ByVal Rating As Long = -1) As Movie
    Set NewMovie = New Movie
    Call NewMovie.Init(Name, Category, Rating)
End Function

Public Function NewStringTokenizer(ByVal Text As String, Optional ByVal DELIMETER As String = ",") As StringTokenizer
    Set NewStringTokenizer = New StringTokenizer
    Call NewStringTokenizer.Init(Text, DELIMETER)
End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Constructor Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function NewCategory(ByVal Name As String) As Category
    Set NewCategory = New Category
    Call NewCategory.Init(Name)
End Function
