Attribute VB_Name = "modWinApi"
Option Explicit

Public Declare Sub PutSAPtr Lib "msvbvm60.dll" Alias "PutMem4" (ByRef Destination() As Any, ByVal Source As Long)

