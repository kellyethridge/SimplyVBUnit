VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextMessageWriter"
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
' Module: MessageWriter
'
Option Explicit

Private Const FIRST_ARGUMENT    As Long = 0
Private Const SECOND_ARGUMENT   As Long = FIRST_ARGUMENT + 1
Private Const PFX_EXPECTED      As String = "Expected: "
Private Const PFX_ACTUAL        As String = "But was : "
Private Const FMT_EMPTYSTRING   As String = "<Empty String>"
Private Const FMT_STRING        As String = """{0}"""
Private Const FMT_NULL          As String = "Null"
Private Const FMT_EMPTY         As String = "Empty"
Private Const FMT_NOTHING       As String = "Nothing"
Private Const FMT_MISSING       As String = "Missing"



Private mMessage As New StringBuilder


Public Function ToString() As String
    ToString = mMessage.ToString
End Function

Public Sub DisplayConstraintDifference(ByVal Constraint As IConstraint)
    Call WriteExpectedLineConstraint(Constraint)
    Call WriteActualLineConstraint(Constraint)
End Sub

Private Sub WriteExpectedLineConstraint(ByVal Constraint As IConstraint)
    Call WriteText(PFX_EXPECTED)
    Call Constraint.WriteDescriptionTo(Me)
    Call WriteLine
End Sub

Private Sub WriteActualLineConstraint(ByVal Constraint As IConstraint)
    Call WriteText(PFX_ACTUAL)
    Call Constraint.WriteActualValueTo(Me)
    Call WriteLine
End Sub

Public Sub DisplayDifference(ByRef Expected As Variant, ByRef Actual As Variant)
    Call WriteExpectedLine(Expected)
    Call WriteActualLine(Actual)
End Sub

Private Sub WriteExpectedLine(ByRef Expected As Variant)
    Call WriteText(PFX_EXPECTED)
    Call WriteExpectedValue(Expected)
    Call WriteLine
End Sub

Private Sub WriteActualLine(ByRef Actual As Variant)
    Call WriteText(PFX_ACTUAL)
    Call WriteActualValue(Actual)
    Call WriteLine
End Sub

Public Sub WriteActualValue(ByRef Value As Variant)
    Call WriteValue(Value)
End Sub

Public Sub WriteExpectedValue(ByRef Value As Variant)
    Call WriteValue(Value)
End Sub

Public Sub WriteValue(ByRef Value As Variant)
    If IsEnumerable(Value) Then
        Call WriteEnumerable(Value)
    ElseIf IsMissing(Value) Then
        Call WriteText(FMT_MISSING)
    Else
        Call WriteDefault(Value)
    End If
End Sub

Private Sub WriteEnumerable(ByRef Values As Variant)
    Call WriteText("< ")
    
    Dim FirstIsWritten  As Boolean
    Dim Value           As Variant
    For Each Value In Values
        If FirstIsWritten Then
            Call WriteText(", ")
        End If
        
        Call WriteValue(Value)
        FirstIsWritten = True
    Next
    
    Call WriteText(" >")
End Sub

Private Sub WriteDefault(ByRef Value As Variant)
    Select Case VarType(Value)
        Case vbString:      Call WriteString(Value)
        Case vbNull:        Call WriteText(FMT_NULL)
        Case vbEmpty:       Call WriteText(FMT_EMPTY)
        Case vbObject:      Call WriteText(FMT_NOTHING)
        Case vbDataObject:  Call WriteText(FMT_NOTHING)
        Case Else:          Call WriteText(CStr(Value))
    End Select
End Sub

Private Sub WriteString(ByVal Value As String)
    If Len(Value) = 0 Then
        Call WriteText(FMT_EMPTYSTRING)
    Else
        Call WriteText(FMT_STRING, Value)
    End If
End Sub

Public Sub WriteLine(ParamArray Args() As Variant)
    Dim UpperBound As Long
    UpperBound = UBound(Args)
    
    If HasOneArgument(UpperBound) Then
        Call WriteText(Args(FIRST_ARGUMENT))
    ElseIf HasMultipleArguments(UpperBound) Then
        Dim LocalArgs() As Variant
        LocalArgs = Args
        Call WriteFormattedText(LocalArgs(FIRST_ARGUMENT), LocalArgs(SECOND_ARGUMENT), UpperBound)
    End If
    
    Call WriteText(vbCrLf)
End Sub

Private Function HasOneArgument(ByVal UpperBound As Long) As Boolean
    HasOneArgument = (UpperBound = 0)
End Function

Private Function HasMultipleArguments(ByVal UpperBound As Long) As Boolean
    HasMultipleArguments = (UpperBound > 0)
End Function

Public Sub WriteText(ByVal Text As String, ParamArray Args() As Variant)
    Dim UpperBound As Long
    UpperBound = UBound(Args)
    
    If HasNoArguments(UpperBound) Then
        Call mMessage.Append(Text)
    Else
        Dim LocalArgs() As Variant
        LocalArgs = Args
        Call mMessage.AppendFormatArray(Text, LocalArgs)
    End If
End Sub

Private Function HasNoArguments(ByVal UpperBound As Long) As Boolean
    HasNoArguments = (UpperBound < 0)
End Function

Private Sub WriteFormattedText(ByVal Text As String, ByRef FirstArgument As Variant, ByVal ArgumentCount As Long)
    Dim Proxy As ArrayProxy
    Call InitArrayProxy(Proxy, FirstArgument, ArgumentCount)
    Call mMessage.AppendFormatArray(Text, Proxy.Data)
End Sub
