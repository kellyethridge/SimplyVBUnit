VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CallTrace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text

Private mMethodNames As New Collection

Public Property Get Count() As Long
    Count = mMethodNames.Count
End Property

Public Sub Add(ByVal MethodName As String)
    Call mMethodNames.Add(MethodName)
End Sub

Public Property Get Item(ByVal Index As Long) As String
Attribute Item.VB_UserMemId = 0
    Item = mMethodNames(Index)
End Property

Public Function Equals(ByVal Other As CallTrace) As Boolean
    If Me.Count <> Other.Count Then Exit Function
    
    Dim i As Long
    For i = 1 To Me.Count
        If Not (Me(i) Like Other(i)) Then Exit Function
    Next
    
    Equals = True
End Function

Public Property Get IsEmpty() As Boolean
    IsEmpty = Me.Count = 0
End Property

Public Function ToString() As String
    Dim sb  As New StringBuilder
    Dim v   As Variant
    
    Dim AddComma As Boolean
    For Each v In mMethodNames
        If AddComma Then
            Call sb.Append(",")
        End If
        
        Call sb.Append(v)
        AddComma = True
    Next
    
    ToString = sb.ToString
End Function
