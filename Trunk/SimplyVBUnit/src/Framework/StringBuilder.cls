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
' Copyright 2009 Kelly Ethridge
'
' Licensed under the Apache License, Version 2.0 (the "License");
' you may not use this file except in compliance with the License.
' You may obtain a copy of the License at
'
'     http://www.apache.org/licenses/LICENSE-2.0
'
' Unless required by applicable law or agreed to in writing, software
' distributed under the License is distributed on an "AS IS" BASIS,
' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
' See the License for the specific language governing permissions and
' limitations under the License.
'
' Module: StringBuilder
'

''
' A class to help build complex strings.
'
' @remarks This is a helper class that provides services to build complex strings
' used during the testing process.
'
Option Explicit

Private Const DEF_CAPACITY  As Long = 32
Private Const MIN_LENGTH    As Long = 0


Private mBuffer             As String
Private mLength             As Long
Private mRequiredCapacity   As Long
Private mNewCapacity        As Long



''
' Formats text by replacing indices with arguments, then appends the result to the string buffer.
'
' @param Text The text that contains indices to be formatted.
' @param Args A list of arguments to replace the indices with.
'
Public Sub AppendFormat(ByVal Text As String, ParamArray Args() As Variant)
    Dim LocalArgs() As Variant
    LocalArgs = Args
    Call AppendFormatArray(Text, LocalArgs)
End Sub

''
' Formats text by replacing indices with arguments, then appends the result to the string buffer.
'
' @param Text The text that contains indices to be formatted.
' @param Args A list of arguments to replace the indices with.
'
Public Sub AppendFormatArray(ByVal Text As String, ByRef Args() As Variant)
    Dim Segments() As String
    Segments = Split(Text, "{{")
    
    Dim j As Long
    For j = 0 To UBound(Segments)
        Dim i As Long
        For i = 0 To UBound(Args)
            Segments(j) = Replace$(Segments(j), "{" & i & "}", Args(i))
        Next
    Next
    
    Text = Join(Segments, "{")
    Text = Replace$(Text, "}}", "}")
    
    Call Append(Text)
End Sub


''
' Appends text to the end of the current string buffer.
'
' @param Text The string to be appended to the buffer.
'
Public Sub Append(ByVal Text As String)
    Call EnsureCapacity(Length + Len(Text))
    Call AppendToBuffer(Text)
End Sub

Private Sub EnsureCapacity(ByVal RequiredCapacity As Long)
    mRequiredCapacity = RequiredCapacity
    
    If NeedsMoreCapacity Then
        Call CalculateNewCapacity
        Call ExpandCapacity
    End If
End Sub

Private Sub AppendToBuffer(ByVal Text As String)
    Call CopyMemory(ByVal StrPtr(mBuffer) + (Length * 2), ByVal StrPtr(Text), LenB(Text))
    Call SetLength(Text)
End Sub

Private Sub SetLength(ByVal Text As String)
    mLength = mLength + Len(Text)
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
    
    Call CopyMemory(ByVal StrPtr(NewBuffer), ByVal StrPtr(mBuffer), Length * 2)
    mBuffer = vbNullString
    
    StringPtr(mBuffer) = StringPtr(NewBuffer)
    StringPtr(NewBuffer) = vbNullPtr
End Sub


''
' Returns the current length of the string buffer.
'
' @return A Long indicating the length of the internal string buffer.
'
Public Property Get Length() As Long
    Length = mLength
End Property


''
' Sets the length for the current string buffer.
'
' @param RHS The length to set the string buffer to.
'
Public Property Let Length(ByVal RHS As Long)
    If RHS < MIN_LENGTH Then _
        Call Err.Raise(ErrorCode.ArgumentOutOfRange, Resource.GetString(Src_StringBuilder_Length), Resource.GetString(Arg_CannotBeNegative, "Length"))
    
    If RHS > mLength Then
        Dim FillSpace As String
        FillSpace = Space$(RHS - mLength)
        
        Call Append(FillSpace)
    Else
        mLength = RHS
    End If
End Property


''
' Returns the current string buffer.
'
' @return The current string buffer.
'
Public Function ToString() As String
    ToString = Left$(mBuffer, Length)
End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mBuffer = String$(DEF_CAPACITY, vbNullChar)
End Sub
