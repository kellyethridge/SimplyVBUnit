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

Private Const FIRST_ARGUMENT        As Long = 0
Private Const SECOND_ARGUMENT       As Long = FIRST_ARGUMENT + 1
Private Const PFX_EXPECTED          As String = "Expected: "
Private Const PFX_ACTUAL            As String = "But was : "
Private Const FMT_STRING            As String = """{0}"""
Private Const FMT_NULL              As String = "Null"
Private Const FMT_EMPTY             As String = "Empty"
Private Const FMT_NOTHING           As String = "Nothing"
Private Const FMT_MISSING           As String = "Missing"
Private Const FMT_CURRENCY          As String = "0.0000@"
Private Const FMT_EMPTYARRAY        As String = "<empty>"
Private Const FMT_PREDICATE         As String = "{0} "
Private Const FMT_MODIFIER          As String = ", {0}"
Private Const FMT_TOLERANCE         As String = "+/-"
Private Const FMT_ERRORNUMBER       As String = "error {0}"
Private Const FMT_ERRORDESCRIPTION  As String = ", {0}"
Private Const FMT_ERRORSOURCE       As String = " from [{0}]"
Private Const FMT_CONNECTOR         As String = " {0} "
Private Const FMT_STRINGSDIFFER_1   As String = "String lengths are both {0}. Strings differ at index {1}."
Private Const FMT_STRINGSDIFFER_2   As String = "Expected string length {0} but was {1}. Strings differ at index {2}."
Private Const FMT_DEFAULT           As String = "<{0}>"
Private Const MAX_DISPLAY_LENGTH    As Long = 64

Private Const SYM_SINGLE            As String = "!"
Private Const SYM_DOUBLE            As String = "#"
Private Const SYM_CURRENCY          As String = "@"
Private Const SYM_DECIMAL           As String = "D"



Private mMessage    As New StringBuilder
Private mUtils      As New MsgUtils


Public Function ToString() As String
    ToString = mMessage.ToString
End Function

Public Sub DisplayStringDifference(ByVal Expected As String, ByVal Actual As String, ByVal IgnoreCase As Boolean, ByVal Clipping As Boolean)
    Dim MismatchIndex As Long
    MismatchIndex = MsgUtils.FindMismatchIndex(Expected, Actual, IgnoreCase)
    
    If Len(Expected) = Len(Actual) Then
        WriteLine FMT_STRINGSDIFFER_1, Len(Expected), MismatchIndex
    Else
        WriteLine FMT_STRINGSDIFFER_2, Len(Expected), Len(Actual), MismatchIndex
    End If
    
    If Clipping Then
        mUtils.ClipExpectedAndActual Expected, Actual, MAX_DISPLAY_LENGTH, MismatchIndex
    End If
    
    MismatchIndex = mUtils.FindMismatchIndex(Expected, Actual, IgnoreCase)
    
    WriteExpectedLine Expected
    WriteActualLine Actual
    
    WriteCaretLine MismatchIndex
End Sub

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

Public Sub WriteActualValue(ByRef Value As Variant)
    Call WriteValue(Value)
End Sub

Public Sub WriteExpectedValue(ByRef Value As Variant)
    Call WriteValue(Value)
End Sub

Public Sub WriteValue(ByRef Value As Variant)
    If IsArray(Value) Then
        WriteArray Value
    Else
        Select Case VarType(Value)
            Case vbString:                  WriteString CStr(Value)
            Case vbLong, vbInteger, vbByte: WriteText CStr(Value)
            Case vbBoolean:                 WriteText CStr(Value)
            Case vbSingle:                  WriteWithDecimalPoint CStr(Value), SYM_SINGLE
            Case vbDouble:                  WriteWithDecimalPoint CStr(Value), SYM_DOUBLE
            Case vbCurrency:                WriteCurrency CCur(Value)
            Case vbDecimal:                 WriteWithDecimalPoint CStr(Value), SYM_DECIMAL
            Case vbDate:                    WriteText CStr(Value)
            Case vbObject, vbDataObject:    WriteObject Value
            Case vbNull:                    WriteText FMT_NULL
            Case vbEmpty:                   WriteText FMT_EMPTY
            Case vbError:                   WriteText FMT_MISSING
            Case Else:                      WriteDefault Value
        End Select
    End If
End Sub

Public Sub WriteLine(ParamArray Args() As Variant)
    Dim UpperBound As Long
    UpperBound = UBound(Args)
    
    If HasOneArgument(UpperBound) Then
        Call WriteText(CStr(Args(FIRST_ARGUMENT)))
    ElseIf HasMultipleArguments(UpperBound) Then
        Dim LocalArgs() As Variant
        LocalArgs = Args
        Call WriteFormattedText(CStr(LocalArgs(FIRST_ARGUMENT)), LocalArgs(SECOND_ARGUMENT), UpperBound)
    End If
    
    Call WriteText(vbCrLf)
End Sub

Public Sub WriteText(ByRef Text As String, ParamArray Args() As Variant)
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

Public Sub WritePredicate(ByRef Text As String)
    Call WriteText(FMT_PREDICATE, Text)
End Sub

Public Sub WriteConnector(ByRef Text As String)
    Call WriteText(FMT_CONNECTOR, Text)
End Sub

Public Sub WriteErrorInfo(ByVal Info As ErrorInfo)
    Call WriteText(FMT_ERRORNUMBER, Info.Number)
    Call WriteOptionalText(Info.Description, FMT_ERRORDESCRIPTION)
    Call WriteOptionalText(Info.Source, FMT_ERRORSOURCE)
End Sub

Public Sub WriteCollectionElements(ByRef Values As Variant, ByVal StartIndex As Long, ByVal Count As Long)
    Dim Index       As Long
    Dim PassedFirst As Boolean
    Dim Element     As Variant
    
    Index = GetLBound(Values)
    
    WriteText "< "
    
    For Each Element In Values
        If Count = 0 Then
            WriteText "..."
            Exit For
        End If
        
        If Index >= StartIndex Then
            If Index > StartIndex Then
                WriteText ", "
            End If
        
            WriteValue Element
            
            Count = Count - 1
        End If
        
        Index = Index + 1
    Next Element
    
    WriteText " >"
End Sub

Public Sub WriteTypePredicate(ByRef Value As Variant)
    Call WriteText("type of <{0}> ", TypeName(Value))
End Sub

Public Sub WriteError(ByVal Info As ErrorInfo)
    Call WriteText(FMT_ERRORNUMBER, Info.Number)
    
    Dim Connector As String
    If Len(Info.Description) > 0 Then
        Call WriteText(FMT_MODIFIER, Info.Description)
        Connector = FMT_CONNECTOR
    Else
        Connector = FMT_MODIFIER
    End If
    
    If Len(Info.Source) > 0 Then
        Call WriteText(Connector, "From")
        Call WriteText(FMT_ERRORSOURCE, Info.Source)
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
    Call InitArrayProxy(Proxy, FirstArgument, ArgumentCount)
    Call mMessage.AppendFormatArray(Text, Proxy.Data)
End Sub

Private Sub WriteOptionalText(ByRef Text As Variant, ByRef TextFormat As String)
    If Len(Text) > 0 Then
        Call WriteText(TextFormat, Text)
    End If
End Sub

Private Sub WriteString(ByRef Value As String)
    Call WriteText(FMT_STRING, Value)
End Sub

Private Sub WriteWithDecimalPoint(ByRef Value As String, ByRef Symbol As String)
    If InStr(Value, ".") > 0 Then
        Call WriteText(Value & Symbol)
    Else
        Call WriteText(Value & ".0" & Symbol)
    End If
End Sub

Private Sub WriteCurrency(ByVal Value As Currency)
    Call WriteText(Format$(Value, FMT_CURRENCY))
End Sub

Private Sub WriteObject(ByRef Value As Variant)
    If Value Is Nothing Then
        Call WriteText(FMT_NOTHING)
    ElseIf TypeOf Value Is Collection Then
        Call WriteCollectionElements(Value, 1, 10)
    ElseIf TypeOf Value Is ErrorInfo Then
        Call WriteError(Value)
    Else
        Call WriteDefault(TypeName(Value))
    End If
End Sub

Private Sub WriteDefault(ByRef Value As Variant)
    Call WriteText(FMT_DEFAULT, Value)
End Sub

Private Sub WriteArray(ByRef Arr As Variant)
    Dim Rank As Long
    Rank = SafeArrayGetDim(GetArrayPointer(Arr))
    
    Select Case Rank
        Case 0: WriteText FMT_NULL
        Case 1
            If UBound(Arr) >= LBound(Arr) Then
                WriteCollectionElements Arr, LBound(Arr), 10
            Else
                WriteDefault "empty"
            End If
            
        Case Else:
            If UBound(Arr) >= LBound(Arr) Then
                Dim Dimensions() As Long
                Dimensions = GetSizeOfEachDimension(Arr, Rank)
                
                Dim Count   As Long
                Dim Element As Variant
                For Each Element In Arr
                    If Count > 0 Then
                        WriteText ", "
                    End If
                                    
                    WriteCheveronForDimensions Dimensions, Count, "< "
                    
                    WriteValue Element
                    Count = Count + 1
                    
                    WriteCheveronForDimensions Dimensions, Count, " >"
                Next Element
            Else
                WriteDefault "empty"
            End If
    End Select
End Sub

Private Sub WriteCaretLine(ByVal MismatchIndex As Long)
    Call WriteLine("  {0}^", String$(Len(PFX_EXPECTED) + MismatchIndex - 2, "-"))
End Sub

Private Function GetSizeOfEachDimension(ByRef Arr As Variant, ByVal Rank As Long) As Long()
    Dim Result() As Long
    ReDim Result(0 To Rank - 1)
    
    Dim PrevSize As Long
    PrevSize = 1
    
    Dim i As Long
    For i = 1 To Rank
        Result(i - 1) = PrevSize * (UBound(Arr, i) - LBound(Arr, i) + 1)
        PrevSize = Result(i - 1)
    Next i

    GetSizeOfEachDimension = Result
End Function

Private Sub WriteCheveronForDimensions(ByRef Dimensions() As Long, ByVal Count As Long, ByRef Cheveron As String)
    Dim i As Long
    
    For i = 0 To UBound(Dimensions)
        If (Count Mod Dimensions(i)) = 0 Then
            WriteText Cheveron
        End If
    Next i
End Sub

