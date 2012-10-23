VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestOutputWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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
' Module: SComponent.TestOutputWriter
'
Option Explicit
Implements ITextWriter

Private mListener   As IEventListener
Private mOutputType As TestOutputType
Private mBuffer     As New StringBuilder


Public Sub WriteText(ByRef Text As String, ParamArray Args() As Variant)
    WriteTextDirect(Text) = Args
End Sub

Public Sub WriteLine(ParamArray TextAndArgs() As Variant)
    WriteLineDirect = TextAndArgs
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Listener As IEventListener, ByVal OutputType As TestOutputType)
    If Listener Is Nothing Then _
        Err.Raise ErrorCode.ArgumentNull, "TestOutputWriter", "Listener cannot be Nothing."
    
    Set mListener = Listener
    mOutputType = OutputType
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Let WriteTextDirect(ByRef Text As String, ByVal Args As Variant)
    mBuffer.Length = 0
    Dim LocalArgs() As Variant
    LocalArgs = Args
    mBuffer.AppendFormatArray Text, LocalArgs
    OutputText
End Property

Private Property Let WriteLineDirect(ByVal TextAndArgs As Variant)
    mBuffer.Length = 0
    Dim LocalArgs() As Variant
    LocalArgs = TextAndArgs
    mBuffer.AppendFormatArrayEx LocalArgs
    mBuffer.Append vbCrLf
    OutputText
End Property

Private Sub OutputText()
    Dim Output As TestOutput
    Set Output = Sim.NewTestOutput(mBuffer.ToString, mOutputType)
    mListener.TestOutput Output
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ITextWriter Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ITextWriter_WriteLine(ParamArray TextAndArgs() As Variant)
    WriteLineDirect = TextAndArgs
End Sub

Private Sub ITextWriter_WriteText(ByVal Text As String, ParamArray Args() As Variant)
    WriteTextDirect(Text) = Args
End Sub
