Attribute VB_Name = "modWin32API"
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
' Module: SComponent.modWin32API
'
Option Explicit

Public Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Public Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Public Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Public Declare Function DestroyWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Public Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Public Declare Function CreateFontIndirect Lib "gdi32.dll" Alias "CreateFontIndirectA" (ByRef lpLogFont As LOGFONT) As Long
Public Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32.dll" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Declare Function GetSysColorBrush Lib "user32.dll" (ByVal nIndex As Long) As Long
Public Declare Function SetBkColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Public Declare Function SetTextColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Sub GetSAPtr Lib "msvbvm60.dll" Alias "GetMem4" (ByRef Arr() As Any, ByRef Result As Long)
Public Declare Sub SetSAPtr Lib "msvbvm60.dll" Alias "PutMem4" (ByRef Arr() As Any, ByVal Value As Long)

Public Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName As String * 32
End Type

Public Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

Public Type DRAWITEMSTRUCT
    CtlType     As Long
    CtlID       As Long
    itemID      As Long
    itemAction  As Long
    itemState   As Long
    hwndItem    As Long
    hdc         As Long
    rcItem      As RECT
    itemData    As Long
End Type

Public Type MEASUREITEMSTRUCT
    CtlType     As Long
    CtlID       As Long
    itemID      As Long
    itemWidth   As Long
    itemHeight  As Long
    itemData    As Long
End Type

Public Type SafeArray1d
    cDims       As Integer
    fFeatures   As Integer
    cbElements  As Long
    cLocks      As Long
    pvData      As Long
    cElements   As Long
    lLBound     As Long
End Type

Public Const vbNullPtr              As Long = 0
Public Const BOOL_TRUE              As Long = 1
Public Const GWL_WNDPROC            As Long = -4
Public Const LB_SETITEMHEIGHT       As Long = &H1A0
Public Const LB_INSERTSTRING        As Long = &H181
Public Const LBS_HASSTRINGS         As Long = &H40&
Public Const LBS_OWNERDRAWVARIABLE  As Long = &H20&
Public Const LBS_OWNERDRAWFIXED     As Long = &H10&
Public Const LB_GETTEXT             As Long = &H189
Public Const LB_GETTEXTLEN          As Long = &H18A
Public Const LB_ADDSTRING           As Long = &H180
Public Const LB_GETCOUNT            As Long = &H18B
Public Const LB_SETHORIZONTALEXTENT As Long = &H194
Public Const LB_RESETCONTENT        As Long = &H184
Public Const LOGPIXELSY             As Long = 90
Public Const LF_FACESIZE            As Long = 32

Public Const WM_DRAWITEM            As Long = &H2B
Public Const WM_MEASUREITEM         As Long = &H2C
Public Const WM_CREATE              As Long = &H1
Public Const WS_EX_CLIENTEDGE       As Long = &H200&
Public Const WS_VSCROLL As Long = &H200000
Public Const WS_HSCROLL As Long = &H100000
Public Const WM_PAINT As Long = &HF&

Public Const WS_EX_WINDOWEDGE As Long = &H100&
Public Const WS_EX_OVERLAPPEDWINDOW As Long = (WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE)
Public Const WS_CHILD As Long = &H40000000
Public Const WS_VISIBLE As Long = &H10000000
Public Const WS_BORDER As Long = &H800000
Public Const ODT_LISTBOX As Long = 2
Public Const ODS_SELECTED As Long = &H1
Public Const DT_LEFT As Long = &H0
Public Const SM_CYCAPTION      As Long = 4
Public Const SM_CYFRAME        As Long = 33
Public Const SM_CXFRAME        As Long = 32

':) Ulli's VB Code Formatter V2.24.17 (2008-Nov-22 13:22)  Decl: 136  Code: 0  Total: 136 Lines
':) CommentOnly: 17 (12.5%)  Commented: 0 (0%)  Filled: 119 (87.5%)  Empty: 17 (12.5%)  Max Logic Depth: 1
