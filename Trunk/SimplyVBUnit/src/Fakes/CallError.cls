VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CallError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mErrors As New Collection



Public Sub Add(ByVal MethodName As String, ByVal Number As Long, Optional ByVal Source As String, Optional ByVal Description As String)
    Dim ErrorInfo As String
    ErrorInfo = CStr(Number) & vbNullChar & Source & vbNullChar & Description
    
    Call mErrors.Add(ErrorInfo, MethodName)
End Sub

Public Sub Raise(ByVal MethodName As String)
    Dim Info() As String
    
    If TryGetErrorInfo(MethodName, Info) Then
        Call Err.Raise(CLng(Info(0)), Info(1), Info(2))
    End If
End Sub

Private Function TryGetErrorInfo(ByRef MethodName As String, ByRef ErrorInfo() As String) As Boolean
    On Error GoTo errTrap
    Dim Info As String
    Info = mErrors(MethodName)
    
    ErrorInfo = Split(Info, vbNullChar)
    TryGetErrorInfo = True
    Exit Function
    
errTrap:
End Function


