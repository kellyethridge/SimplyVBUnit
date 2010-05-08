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

Private Const DEF_MAXDISPLAYLENGTH  As Long = 64
Private Const FIRST_ARGUMENT        As Long = 0
Private Const SECOND_ARGUMENT       As Long = FIRST_ARGUMENT + 1

Private Const PFX_EXPECTED          As String = "  Expected: "
Private Const PFX_ACTUAL            As String = "  But was : "
Private Const FMT_STRING            As String = """{0}"""
Private Const FMT_NULL              As String = "Null"
Private Const FMT_EMPTY             As String = "Empty"
Private Const FMT_NOTHING           As String = "Nothing"
Private Const FMT_MISSING           As String = "Missing"
Private Const FMT_EMPTYARRAY        As String = "<empty>"
Private Const FMT_PREDICATE         As String = "{0} "
Private Const FMT_MODIFIER          As String = ", {0}"
Private Const FMT_TOLERANCE         As String = "+/-"
Private Const FMT_ERRORNUMBER       As String = "error {0}"
Private Const FMT_ERRORDESCRIPTION  As String = ", {0}"
Private Const FMT_ERRORSOURCE       As String = " from [{0}]"
Private Const FMT_CONNECTOR         As String = " {0} "
Private Const FMT_DEFAULT           As String = "<{0}>"
Private Const FMT_TYPEOF            As String = "type of <{0}> "
Private Const FMT_DATE              As String = "yyyy-MM-dd h:nn:ss AM/PM"
Private Const SYM_SINGLE            As String = "!"
Private Const SYM_DOUBLE            As String = "#"
Private Const SYM_CURRENCY          As String = "@"
Private Const SYM_DECIMAL           As String = "D"



Private mMaxDisplayLength   As Long
Private mMessage            As New StringBuilder
Private mUtils              As New MsgUtils


Public Property Get MaxDisplayLength() As Long
    MaxDisplayLength = mMaxDisplayLength
End Property

Public Property Let MaxDisplayLength(ByVal RHS As Long)
    mMaxDisplayLength = RHS
End Property

Public Function ToString() As String
    ToString = mMessage.ToString
End Function

Public Sub DisplayStringDifferences(ByVal Expected As String, ByVal Actual As String, ByVal MismatchIndex As Long, ByVal IgnoreCase As Boolean, ByVal Clipping As Boolean)
    If Clipping Then
        mUtils.ClipExpectedAndActual Expected, Actual, Me.MaxDisplayLength, MismatchIndex
    End If
    
    Expected = mUtils.EscapeControlChars(Expected)
    Actual = mUtils.EscapeControlChars(Actual)
    
    MismatchIndex = mUtils.FindMismatchIndex(Expected, Actual, IgnoreCase)
    
    WriteText PFX_EXPECTED
    WriteValue Expected
    
    If IgnoreCase Then
        WriteModifier "ignoring case"
    End If
    
    WriteLine
    WriteActualLine Actual
    WriteCaretLine MismatchIndex
End Sub

Public Sub DisplayDifferences(ByRef Expected As Variant, ByRef Actual As Variant, Optional ByVal Tolerance As Tolerance)
    If Tolerance Is Nothing Then
        Set Tolerance = modStatics.Tolerance.EmptyTolerance
    End If
    
    WriteText PFX_EXPECTED
    WriteValue Expected
    
    If Tolerance.Mode <> NoneMode Then
        WriteConnector FMT_TOLERANCE
        WriteText Tolerance.Amount
        
        Select Case Tolerance.Interval
            Case "h":       WriteText " hour(s)"
            Case "n":       WriteText " minute(s)"
            Case "s":       WriteText " second(s)"
            Case "d":       WriteText " day(s)"
            Case "yyyy":    WriteText " year(s)"
            Case "ww":      WriteText " weeks(s)"
            Case "m":       WriteText " month(s)"
            Case "q":       WriteText " quarter(s)"
            Case "td":      WriteText " total day(s)"
            Case "th":      WriteText " total hour(s)"
            Case "tn":      WriteText " total minute(s)"
        End Select
    End If
    
    WriteLine
    WriteActualLine Actual
End Sub

Public Sub DisplayConstraintDifferences(ByVal Constraint As IConstraint)
    WriteExpectedLineConstraint Constraint
    WriteActualLineConstraint Constraint
End Sub

Private Sub WriteExpectedLineConstraint(ByVal Constraint As IConstraint)
    WriteText PFX_EXPECTED
    Constraint.WriteDescriptionTo Me
    WriteLine
End Sub

Private Sub WriteActualLineConstraint(ByVal Constraint As IConstraint)
    WriteText PFX_ACTUAL
    Constraint.WriteActualValueTo Me
    WriteLine
End Sub

Public Sub WriteActualValue(ByRef Value As Variant)
    WriteValue Value
End Sub

Public Sub WriteExpectedValue(ByRef Value As Variant)
    WriteValue Value
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
            Case vbDecimal:                 WriteDecimal Value
            Case vbDate:                    WriteText Format$(Value, FMT_DATE)
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
        WriteText CStr(Args(FIRST_ARGUMENT))
    ElseIf HasMultipleArguments(UpperBound) Then
        Dim LocalArgs() As Variant
        LocalArgs = Args
        WriteFormattedText CStr(LocalArgs(FIRST_ARGUMENT)), LocalArgs(SECOND_ARGUMENT), UpperBound
    End If
    
    WriteText vbCrLf
End Sub

Public Sub WriteText(ByRef Text As String, ParamArray Args() As Variant)
    Dim UpperBound As Long
    UpperBound = UBound(Args)
    
    If HasNoArguments(UpperBound) Then
        mMessage.Append Text
    Else
        Dim LocalArgs() As Variant
        LocalArgs = Args
        mMessage.AppendFormatArray Text, LocalArgs
    End If
End Sub

Public Sub WritePredicate(ByRef Text As String)
    WriteText FMT_PREDICATE, Text
End Sub

Public Sub WriteModifier(ByRef Text As String)
    WriteText FMT_MODIFIER, Text
End Sub

Public Sub WriteConnector(ByRef Text As String)
    WriteText FMT_CONNECTOR, Text
End Sub

Public Sub WriteErrorInfo(ByVal Info As ErrorInfo)
    If Info Is Nothing Then _
        Error.CannotBeNothing Me, "WriteErrorInfo", "Info"
    
    WriteText FMT_ERRORNUMBER, Info.Number
    WriteOptionalText Info.Description, FMT_ERRORDESCRIPTION
    WriteOptionalText Info.Source, FMT_ERRORSOURCE
End Sub

Public Sub WriteCollectionElements(ByRef Values As Variant, ByVal StartIndex As Long, ByVal Count As Long)
    Dim Index       As Long
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

Public Sub WriteTypeOfPredicate(ByRef Value As Variant)
    WriteText FMT_TYPEOF, TypeName(Value)
End Sub

Public Sub WriteError(ByVal Info As ErrorInfo)
    WriteText FMT_ERRORNUMBER, Info.Number
    
    Dim Connector As String
    If Len(Info.Description) > 0 Then
        WriteModifier Info.Description
        Connector = FMT_CONNECTOR
    Else
        Connector = FMT_MODIFIER
    End If
    
    If Len(Info.Source) > 0 Then
        WriteText Connector, "From"
        WriteText FMT_ERRORSOURCE, Info.Source
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub WriteExpectedLine(ByRef Expected As Variant)
    WriteText PFX_EXPECTED
    WriteExpectedValue Expected
    WriteLine
End Sub

Private Sub WriteActualLine(ByRef Actual As Variant)
    WriteText PFX_ACTUAL
    WriteActualValue Actual
    WriteLine
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
    mMessage.AppendFormatArray Text, Proxy.Data
End Sub

Private Sub WriteOptionalText(ByRef Text As Variant, ByRef TextFormat As String)
    If Len(Text) > 0 Then
        WriteText TextFormat, Text
    End If
End Sub

Private Sub WriteString(ByRef Value As String)
    WriteText FMT_STRING, Value
End Sub

Private Sub WriteWithDecimalPoint(ByRef Value As String, ByRef Symbol As String)
    If InStr(Value, ".") > 0 Then
        WriteText Value & Symbol
    Else
        WriteText Value & ".0" & Symbol
    End If
End Sub

Private Sub WriteDecimal(ByRef Value As Variant)
    WriteText CStr(Value) & SYM_DECIMAL
End Sub

Private Sub WriteCurrency(ByVal Value As Currency)
    WriteText CStr(Value) & SYM_CURRENCY
End Sub

Private Sub WriteObject(ByRef Value As Variant)
    If Value Is Nothing Then
        WriteText FMT_NOTHING
    ElseIf TypeOf Value Is Collection Then
        WriteCollectionElements Value, 1, 10
    ElseIf TypeOf Value Is ErrorInfo Then
        WriteError Value
    Else
        WriteDefault TypeName(Value)
    End If
End Sub

Private Sub WriteDefault(ByRef Value As Variant)
    WriteText FMT_DEFAULT, Value
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
    WriteLine "  {0}^", String$(Len(PFX_EXPECTED) + MismatchIndex - 2, "-")
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


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Me.MaxDisplayLength = DEF_MAXDISPLAYLENGTH
End Sub
