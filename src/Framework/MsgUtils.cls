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

Private Const MSG_ELLIPSIS          As String = "..."
Private Const MSG_BOUNDS            As String = "{0} To {1}"

Private Const ASC_NULL              As Long = 0
Private Const ASC_BELL              As Long = 4
Private Const ASC_FORMFEED          As Long = 7
Private Const ASC_BACKSPACE         As Long = 8
Private Const ASC_LINEFEED          As Long = 10
Private Const ASC_CARRIAGERETURN    As Long = 13
Private Const ASC_HORIZONTALTAB     As Long = 9
Private Const ASC_VERTICALTAB       As Long = 11
Private Const ESC_NULL              As String = "\0"
Private Const ESC_BELL              As String = "\a"
Private Const ESC_FORMFEED          As String = "\f"
Private Const ESC_BACKSPACE         As String = "\b"
Private Const ESC_LINEFEED          As String = "\n"
Private Const ESC_CARRIAGERETURN    As String = "\r"
Private Const ESC_HORIZONTALTAB     As String = "\t"
Private Const ESC_VERTICALTAB       As String = "\v"
Private Const ESC_HEXPREFIX         As String = "\x"



Public Function FindMismatchIndex(ByRef Expected As String, ByRef Actual As String, ByVal IgnoreCase As Boolean) As Long
    Dim MaxIndex As Long
    MaxIndex = Min(Len(Expected), Len(Actual))
    
    Dim i As Long
    For i = 1 To MaxIndex
        If Not EqualStrings(Mid$(Expected, i, 1), Mid$(Actual, i, 1), IgnoreCase) Then
            FindMismatchIndex = i
            Exit Function
        End If
    Next i
    
    If Len(Expected) <> Len(Actual) Then
        FindMismatchIndex = MaxIndex + 1
    End If
End Function

Public Function ClipString(ByRef s As String, ByVal StartIndex As Long, ByVal MaxStringLength As Long) As String
    Dim sb As New StringBuilder

    If StartIndex > 1 Then
        MaxStringLength = MaxStringLength - Len(MSG_ELLIPSIS)
        sb.Append MSG_ELLIPSIS
    End If

    If Len(s) - StartIndex + 1 > MaxStringLength Then
        MaxStringLength = MaxStringLength - Len(MSG_ELLIPSIS)
        sb.Append Mid$(s, StartIndex, MaxStringLength)
        sb.Append MSG_ELLIPSIS
    ElseIf StartIndex > 1 Then
        sb.Append Mid$(s, StartIndex)
    Else
        sb.Append s
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

Public Function EscapeControlChars(ByRef s As String) As String
    Dim sb As New StringBuilder
    
    Dim i As Long
    For i = 1 To Len(s)
        Dim Ch As String
        Ch = Mid$(s, i, 1)
        
        Select Case AscW(Ch)
            Case ASC_NULL:              sb.Append ESC_NULL
            Case ASC_LINEFEED:          sb.Append ESC_LINEFEED
            Case ASC_CARRIAGERETURN:    sb.Append ESC_CARRIAGERETURN
            Case ASC_BACKSPACE:         sb.Append ESC_BACKSPACE
            Case ASC_HORIZONTALTAB:     sb.Append ESC_HORIZONTALTAB
            Case ASC_VERTICALTAB:       sb.Append ESC_VERTICALTAB
            Case ASC_FORMFEED:          sb.Append ESC_FORMFEED
            Case ASC_BELL:              sb.Append ESC_BELL
            Case Is < 32, Is > 127
                                        sb.Append ESC_HEXPREFIX
                                        sb.Append Right$("0000" & Hex$(AscW(Ch)), 4)
            Case Else:                  sb.Append Ch
        End Select
    Next
        
    Dim Result As String
    If sb.Length > 0 Then
        Result = sb.ToString
    Else
        Result = s
    End If
    
    EscapeControlChars = Result
End Function

Public Function GetTypeRepresentation(ByRef Value As Variant) As String
    Dim sb As New StringBuilder

    sb.Append "<"

    If IsArray(Value) Then
        sb.Append TypeName(Value)
        sb.Length = sb.Length - 1

        Dim i As Long
        For i = 1 To GetArrayRank(Value)
            If i > 1 Then
                sb.Append ", "
            End If

            sb.AppendFormat MSG_BOUNDS, LBound(Value, i), UBound(Value, i)
        Next i

        sb.Append ")"
    Else
        On Error GoTo errTrap
        
        If IsObject(Value) Then
            Dim Info As TypeInfo
            Set Info = tli.ClassInfoFromObject(Value)
            sb.Append Info.Parent.Name
            sb.Append "."
        End If
    
errTrap:
        sb.Append TypeName(Value)
        Err.Clear
    End If

    sb.Append ">"
    GetTypeRepresentation = sb.ToString
End Function

Public Function GetArrayIndicesFromCollectionIndex(ByRef Value As Variant, ByVal Index As Long) As Long()
    Dim Result() As Long

    If IsArray(Value) Then
        If Not IsEmptyOrNullArray(Value) Then
            Dim Rank As Long
            Rank = GetArrayRank(Value)
    
            ReDim Result(0 To Rank - 1)
    
            Dim i As Long
            For i = 1 To Rank
                Dim j As Long
                j = UBound(Value, i) - LBound(Value, i) + 1
                
                Result(i - 1) = LBound(Value, i) + (Index Mod j)
                Index = Index \ j
            Next i
        Else
            Result = NewLongs(0)
        End If
    Else
        Dim Count As Long
        TryGetCount Value, Count
        
        If Count > 0 Then
            ReDim Result(0)
            Result(0) = Index
        Else
            Result = NewLongs(0)
        End If
    End If

    GetArrayIndicesFromCollectionIndex = Result
End Function

Public Function GetArrayIndicesAsString(ByRef Indices() As Long) As String
    
    On Error GoTo errTrap
    
    Dim sb  As New StringBuilder
    Dim i   As Long
    For i = LBound(Indices) To UBound(Indices)
        If i > LBound(Indices) Then
            sb.Append ","
        End If

        sb.Append Indices(i)
    Next i

    GetArrayIndicesAsString = sb.ToString
    Exit Function
    
errTrap:
    GetArrayIndicesAsString = vbNullString
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Max(ByVal X As Long, ByVal Y As Long) As Long
    If X > Y Then
        Max = X
    Else
        Max = Y
    End If
End Function

Private Function Min(ByVal X As Long, ByVal Y As Long) As Long
    If X < Y Then
        Min = X
    Else
        Min = Y
    End If
End Function
