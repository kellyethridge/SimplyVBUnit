VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MovieStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
' MovieStatic
'
Option Explicit


' Reads a record from the data source and creates a movie object from it.
'
Public Function ReadFrom(ByVal Reader As FileReader) As Movie
    Dim Result As Movie
    
    If Reader.HasMore Then
        Dim Name        As String
        Dim Category    As Category
        Dim Rating      As Long
        Dim Tokenizer   As StringTokenizer
        
        Set Tokenizer = MB.NewStringTokenizer(Reader.ReadLine, RECORD_DELIMETER)
        Name = Tokenizer.NextToken
        Set Category = modStaticClasses.Category.GetCategoryByName(Tokenizer.NextToken)
        Rating = CLng(Tokenizer.NextToken)
                
        Set Result = MB.newMovie(Name, Category, Rating)
    End If
    
    Set ReadFrom = Result
End Function
