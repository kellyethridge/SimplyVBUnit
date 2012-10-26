VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
' FileReader
'
Option Explicit

Private Const ERR_FILEALREADYOPEN As Long = 55

Private mFileNumber As Long


Public Sub OpenFile(ByVal Filename As String)
    If Me.IsOpen Then _
        Call Err.Raise(ERR_FILEALREADYOPEN, , "File <" & Filename & "> is already open.")
        
    mFileNumber = FreeFile
    Open Filename For Input As #mFileNumber
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

Public Function ReadLine() As String
    Line Input #mFileNumber, ReadLine
End Function

Public Property Get HasMore() As Boolean
    If Me.IsOpen Then
        HasMore = Not EOF(mFileNumber)
    End If
End Property

Private Sub Class_Terminate()
    Call CloseFile
End Sub
