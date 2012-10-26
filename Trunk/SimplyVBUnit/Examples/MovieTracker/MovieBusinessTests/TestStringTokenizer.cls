VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestStringTokenizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestStringTokenizer
'
Option Explicit

Public Sub TestEmptyString()
    Dim t As StringTokenizer
    Set t = MB.NewStringTokenizer("")
    
    Assert.That t.HasTokens, Iz.False, "t should have no tokens."
End Sub

Public Sub TestStringWithOneToken()
    Dim t As StringTokenizer
    Set t = MB.NewStringTokenizer("Field1")
    
    Assert.That t.HasTokens, Iz.True, "t should have first token."
    Assert.That t.NextToken, Iz.EqualTo("Field1"), "token should be 'Field1'."
    Assert.That t.HasTokens, Iz.False, "t should not have token."
End Sub

Public Sub TestStringWithTwoTokens()
    Dim t As StringTokenizer
    Set t = MB.NewStringTokenizer("Field1,Field2")
    
    Assert.That t.NextToken, Iz.EqualTo("Field1"), "first token should be 'Field1'"
    Assert.That t.NextToken, Iz.EqualTo("Field2"), "second token should be 'Field2'"
    Assert.That t.HasTokens, Iz.False, "t should not have any tokens."
End Sub

Public Sub TestWithDifferentDelimeter()
    Dim t As StringTokenizer
    Set t = MB.NewStringTokenizer("FieldOne|FieldTwo", "|")
    
    Assert.That t.NextToken, Iz.EqualTo("FieldOne"), "first token should be 'FieldOne'"
    Assert.That t.NextToken, Iz.EqualTo("FieldTwo"), "second token should be 'FieldTwo'"
    Assert.That t.HasTokens, Iz.False, "t should not have any tokens."
End Sub
