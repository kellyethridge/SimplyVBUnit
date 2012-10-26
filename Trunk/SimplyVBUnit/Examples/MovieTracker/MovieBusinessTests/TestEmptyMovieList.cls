VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestEmptyMovieList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestEmptyMovieList
'
Option Explicit

Private MovieList As MovieList


Public Sub Setup()
    Set MovieList = New MovieList
End Sub


Public Sub TestCount()
    Assert.That MovieList.Count, Iz.EqualTo(0), "Count of empty movie list should be 0."
End Sub

