VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringTokenizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
' StringTokenizer
'
Option Explicit

Private mText       As String
Private mDelimeter  As String



' Gets if there are any more tokens remaining.
'
Public Property Get HasTokens() As Boolean
    HasTokens = (Len(mText) > 0)
End Property

' Returns the next string token if any remain.
'
Public Function NextToken() As String
    If Not HasTokens Then _
        Call Err.Raise(5, , "No more tokens.")
        
    Dim Index As Long
    Index = InStr(mText, mDelimeter)
    
    Dim Result As String
    If Index > 0 Then
        Result = Left$(mText, Index - 1)
        mText = Mid$(mText, Index + 1)
    Else
        Result = mText
        mText = ""
    End If
    
    NextToken = Result
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef Text As String, ByRef Delimeter As String)
    If Len(Delimeter) = 0 Then _
        Call Err.Raise(5, "StringTokenizer", "Delimeter cannot be empty.")
        
    mText = Text
    mDelimeter = Delimeter
End Sub


