Attribute VB_Name = "modUtilities"
Option Explicit

Private Const Q As String = """"

Public Function Quote(ByVal Text As String) As String
    Quote = Q & Text & Q
End Function

