VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextWriterDispatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'    CopyRight (c) 2008 Kelly Ethridge
'
'    This file is part of SimplyVBUnit.
'
'    SimplyVBUnit is free software; you can redistribute it and/or modify
'    it under the terms of the GNU Library General Public License as published by
'    the Free Software Foundation; either version 2.1 of the License, or
'    (at your option) any later version.
'
'    SimplyVBUnit is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Library General Public License for more details.
'
'    You should have received a copy of the GNU Library General Public License
'    along with Foobar; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'    Module: TextWriterDispatcher
'

Option Explicit
Implements ITextWriter

Private mOutputType As TestOutputType
Private mListener   As IEventListener


Public Sub WriteText(ByVal Text As String)
    Dim Output As TestOutput
    Set Output = Sim.NewTestOutput(Text, mOutputType)
    Call mListener.TestOutput(Output)
End Sub

Public Sub WriteLine(Optional ByVal Text As String)
    Call WriteText(Text & vbCrLf)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Listener As IEventListener, ByVal OutputType As TestOutputType)
    Set mListener = Listener
    mOutputType = OutputType
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ITextWriter Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ITextWriter_WriteLine(Optional ByVal Text As String)
    Call WriteLine(Text)
End Sub

Private Sub ITextWriter_WriteText(ByVal Text As String)
    Call WriteText(Text)
End Sub
