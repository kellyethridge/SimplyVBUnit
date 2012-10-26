VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
' FileWriter
'
Option Explicit

Private Const ERR_FILEALREADYOPEN As Long = 55

Private mFileNumber As Long



Public Sub OpenFile(ByVal Filename As String)
    If Me.IsOpen Then _
        Call Err.Raise(ERR_FILEALREADYOPEN, , "File <" & Filename & "> is already open.")
        
    mFileNumber = FreeFile
    Open Filename For Output As #mFileNumber
End Sub

Public Sub CloseFile()
    If Me.IsOpen Then
        Close #mFileNumber
        mFileNumber = 0
    End If
End Sub

Public Property Get IsOpen() As Boolean
    IsOpen = CBool(mFileNumber)
End Property

Public Sub WriteText(ByVal Text As String)
    Print #mFileNumber, Text;
End Sub

Private Sub Class_Terminate()
    Call CloseFile
End Sub
