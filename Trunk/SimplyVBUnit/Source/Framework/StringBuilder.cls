VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2012 Kelly Ethridge
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights to
'use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of
'the Software, and to permit persons to whom the Software is furnished to do so,
'subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
'INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
'PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
'FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
'OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
'DEALINGS IN THE SOFTWARE.
'
'
' Module: StringBuilder
'
Option Explicit

Private Const DEF_CAPACITY      As Long = 128
Private Const MIN_LENGTH        As Long = 0
Private Const FIRST_ARGUMENT    As Long = 0
Private Const SECOND_ARGUMENT   As Long = FIRST_ARGUMENT + 1


Private mBuffer             As String
Private mLength             As Long
Private mRequiredCapacity   As Long
Private mNewCapacity        As Long


Public Sub AppendFormat(ByRef Format As String, ParamArray Args() As Variant)
    Dim LocalArgs() As Variant
    LocalArgs = Args
    AppendFormatArray Format, LocalArgs
End Sub

Public Sub AppendFormatArray(ByRef Format As String, ByRef Args() As Variant)
    Dim Segments() As String
    Segments = Split(Format, "{{")

    Dim j As Long
    For j = 0 To UBound(Segments)
        Dim i As Long
        For i = 0 To UBound(Args)
            Segments(j) = Replace$(Segments(j), "{" & i & "}", GetStringValue(Args(i)))
        Next
    Next

    Dim Result As String
    Result = Join(Segments, "{")
    Result = Replace$(Result, "}}", "}")
    
    AppendText Result
End Sub

Public Sub AppendFormatArrayEx(ByRef TextAndArgs() As Variant)
    Dim UpperBound As Long
    UpperBound = UBound(TextAndArgs)
    
    If HasOneArgument(UpperBound) Then
        AppendText CStr(TextAndArgs(FIRST_ARGUMENT))
    ElseIf HasMultipleArguments(UpperBound) Then
        Dim LocalArgs() As Variant
        LocalArgs = TextAndArgs
        WriteFormattedText CStr(LocalArgs(FIRST_ARGUMENT)), LocalArgs(SECOND_ARGUMENT), UpperBound
    End If
End Sub

Private Function HasOneArgument(ByVal UpperBound As Long) As Boolean
    HasOneArgument = (UpperBound = 0)
End Function

Private Function HasMultipleArguments(ByVal UpperBound As Long) As Boolean
    HasMultipleArguments = (UpperBound > 0)
End Function

Private Function HasNoArguments(ByVal UpperBound As Long) As Boolean
    HasNoArguments = (UpperBound < 0)
End Function

Private Sub WriteFormattedText(ByRef Text As String, ByRef FirstArgument As Variant, ByVal ArgumentCount As Long)
    Dim Proxy As ArrayProxy
    InitArrayProxy Proxy, FirstArgument, ArgumentCount
    AppendFormatArray Text, Proxy.Data
End Sub

Public Sub Append(ByVal Value As String)
    AppendText Value
End Sub

Public Property Get Length() As Long
    Length = mLength
End Property

Public Property Let Length(ByVal RHS As Long)
    If RHS < MIN_LENGTH Then _
        Error.CannotBeNegative Me, "Length", "Length"
    
    If RHS > mLength Then
        Dim FillSpace As String
        FillSpace = Space$(RHS - mLength)
        
        Append FillSpace
    Else
        mLength = RHS
    End If
End Property

Public Function ToString() As String
    ToString = Left$(mBuffer, mLength)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetStringValue(ByRef Value As Variant) As String
    Dim Result As String
    
    If IsArray(Value) Or IsObject(Value) Then
        Result = MsgUtils.GetTypeRepresentation(Value)
    Else
        Result = CStr(Value)
    End If
    
    GetStringValue = Result
End Function

Private Sub AppendText(ByRef Text As String)
    EnsureCapacity Length + Len(Text)
    AppendToBuffer Text
End Sub

Private Sub EnsureCapacity(ByVal RequiredCapacity As Long)
    mRequiredCapacity = RequiredCapacity
    
    If NeedsMoreCapacity Then
        CalculateNewCapacity
        ExpandCapacity
    End If
End Sub

Private Function NeedsMoreCapacity() As Boolean
    NeedsMoreCapacity = (mRequiredCapacity > CurrentCapacity)
End Function

Private Sub CalculateNewCapacity()
    Dim NewCapacity As Long
    NewCapacity = CurrentCapacity
    
    Do While mRequiredCapacity > NewCapacity
        NewCapacity = NewCapacity * 2
    Loop
    
    mNewCapacity = NewCapacity
End Sub

Private Function CurrentCapacity() As Long
    CurrentCapacity = Len(mBuffer)
End Function

Private Sub ExpandCapacity()
    Dim NewBuffer As String
    NewBuffer = String$(mNewCapacity, vbNullChar)
    
    CopyMemory ByVal StrPtr(NewBuffer), ByVal StrPtr(mBuffer), Length * 2
    mBuffer = vbNullString
    
    StringPtr(mBuffer) = StringPtr(NewBuffer)
    StringPtr(NewBuffer) = vbNullPtr
End Sub

Private Sub AppendToBuffer(ByRef Text As String)
    CopyMemory ByVal StrPtr(mBuffer) + (Length * 2), ByVal StrPtr(Text), LenB(Text)
    SetLength Text
End Sub

Private Sub SetLength(ByRef Text As String)
    mLength = mLength + Len(Text)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mBuffer = String$(DEF_CAPACITY, vbNullChar)
End Sub
