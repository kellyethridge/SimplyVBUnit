VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MsgUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'// ****************************************************************
'// Copyright 2007, Charlie Poole
'// This is free software licensed under the NUnit license. You may
'// obtain a copy of the license at http://nunit.org/?p=license&r=2.4
'// ****************************************************************

'
' Converted MsgUtils.cs to VB6.
'
Option Explicit

Private Const MSG_ELLIPSIS  As String = "..."
Private Const MSG_BOUNDS    As String = "{0} To {1}"


Public Function FindMismatchIndex(ByVal Expected As String, ByVal Actual As String, ByVal IgnoreCase As Boolean) As Long
    Dim MaxIndex As Long
    MaxIndex = GetShorterStringLength(Expected, Actual)
    
    Dim i As Long
    For i = 1 To MaxIndex
        If StringCompare(Mid$(Expected, i, 1), Mid$(Actual, i, 1), IgnoreCase) <> 0 Then
            FindMismatchIndex = i
            Exit Function
        End If
    Next i
    
    If AreDifferentLengths(Expected, Actual) Then
        FindMismatchIndex = MaxIndex + 1
    End If
End Function

Public Function ClipString(ByVal s As String, ByVal StartIndex As Long, ByVal MaxStringLength As Long) As String
    Dim sb As New StringBuilder

    If StartIndex > 1 Then
        MaxStringLength = MaxStringLength - Len(MSG_ELLIPSIS)
        Call sb.Append(MSG_ELLIPSIS)
    End If

    If Len(s) - StartIndex + 1 > MaxStringLength Then
        MaxStringLength = MaxStringLength - Len(MSG_ELLIPSIS)
        Call sb.Append(Mid$(s, StartIndex, MaxStringLength))
        Call sb.Append(MSG_ELLIPSIS)
    ElseIf StartIndex > 1 Then
        Call sb.Append(Mid$(s, StartIndex))
    Else
        Call sb.Append(s)
    End If

    ClipString = sb.ToString
End Function

Public Sub ClipExpectedAndActual(ByRef Expected As String, ByRef Actual As String, ByVal MaxDisplayLength As Long, ByVal MismatchIndex As Long)
    Dim MaxStringLength As Long
    MaxStringLength = Max(Len(Expected), Len(Actual))
    
    If MaxStringLength <= MaxDisplayLength Then
        Exit Sub
    End If

    Dim ClipLength  As Long
    Dim ClipStart   As Long
    
    ClipLength = MaxDisplayLength - Len(MSG_ELLIPSIS)
    ClipStart = MaxStringLength - ClipLength + 1

    If ClipStart > MismatchIndex Then
        ClipStart = Max(1, MismatchIndex - ClipLength \ 2)
    End If

    Expected = ClipString(Expected, ClipStart, MaxDisplayLength)
    Actual = ClipString(Actual, ClipStart, MaxDisplayLength)
End Sub

'Public Function GetTypeRepresentation(ByRef Value As Variant) As String
'    Dim sb As New StringBuilder
'
'    Call sb.Append("<")
'
'    If IsArray(Value) Then
'        Call sb.Append(TypeName(Value))
'        sb.Length = sb.Length - 1
'
'        Dim i As Long
'        For i = 1 To GetArrayRank(Value)
'            If i > 1 Then
'                Call sb.Append(", ")
'            End If
'
'            Call sb.AppendFormat(MSG_BOUNDS, LBound(Value, i), UBound(Value, i))
'        Next i
'
'        Call sb.Append(")")
'    Else
'        Call sb.Append(TypeName(Value))
'    End If
'
'    Call sb.Append(">")
'    GetTypeRepresentation = sb.ToString
'End Function
'
'Public Function GetArrayIndicesFromCollectionIndex(ByRef Value As Variant, ByVal Index As Long) As Long()
'    Dim Result() As Long
'
'    If IsArray(Value) Then
'        Dim Rank As Long
'        Rank = GetArrayRank(Value)
'
'        ReDim Result(0 To Rank - 1)
'
'        Dim i As Long
'        For i = 1 To Rank
'            Dim j As Long
'            j = UBound(Value, i) - LBound(Value, i) + 1
'            Result(i - 1) = LBound(Value, i) + (Index Mod j)
'            Index = Index \ j
'        Next i
'    Else
'        ReDim Result(0)
'        Result(0) = Index
'    End If
'
'    GetArrayIndicesFromCollectionIndex = Result
'End Function
'
'Public Function GetArrayIndicesAsString(ByRef Indices() As Long) As String
'    Dim sb As New StringBuilder
'    Dim i As Long
'    For i = 0 To UBound(Indices)
'        If i > 0 Then
'            Call sb.Append(", ")
'        End If
'
'        Call sb.Append(Indices(i))
'    Next i
'
'    GetArrayIndicesAsString = sb.ToString
'End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetShorterStringLength(ByRef String1 As String, ByRef String2 As String) As Long
    Dim String1Length As Long
    Dim String2Length As Long
    
    String1Length = Len(String1)
    String2Length = Len(String2)
    
    Dim Result As Long
    If String1Length < String2Length Then
        Result = String1Length
    Else
        Result = String2Length
    End If
    
    GetShorterStringLength = Result
End Function

Private Function AreDifferentLengths(ByRef String1 As String, ByRef String2 As String) As Boolean
    AreDifferentLengths = (Len(String1) <> Len(String2))
End Function

Private Function Max(ByVal X As Long, ByVal Y As Long) As Long
    If X > Y Then
        Max = X
    Else
        Max = Y
    End If
End Function
