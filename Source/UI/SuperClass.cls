VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SuperClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Name.......... cSuperClass
'File.......... cSuperClass.cls
'Dependencies.. Requires iSuperClass as the model implementation interface.
'Description... A novel window subclassing class that echews the use of a module by dynamically
'               assembling machine code.
'Author........ Paul_Caton@hotmail.com
'Date.......... June, 13th 2002
'Copyright..... None.
'
'v1.00 20020613 First cut......................................................................
'
'v1.01 20020621 Decided to split the single interface iSuperClass_Message into two,
'               iSuperClass_After and iSuperClass_Before. This is slightly more efficient
'               in that the more common *AFTER* the previous WndProc subclassing mode
'               was receiving a redundant parameter (lHandled) also, it reminds the
'               user in which of the two modes the message was added (AddMsg)..................
'
'               Optimized the assembler opcodes a bit.
'               Now using EIP relative calls.
'               WNDPROC_FILTERED is now 10 bytes shorter and slightly faster
'               WNDPROC_ALL is now 20 bytes shorter and slightly faster........................
'
'v1.02 20020627 Spotted that you could UnSubclass and still receive 1 more callback which
'               could stop an unload or worse. Scenario: you AddMsg WM_NCLBUTTONDOWN and
'               click on the close button, the message goes to default processing first which
'               tells the form to unload wherein you call UnSubclass; at this point default
'               processing ends and execution returns to our WndProc who now wants to call
'               iSuperClass_After. The solution is to patch the WndProc code in UnSubclass
'               so that a return is patched between def processing and the call to
'               iSubClass_After................................................................
'
'v1.03 20020627 Added the AllMsgs mode of operation
'               I'm now reasonably confident that cSuperClass is immune to the IDE End button,
'               I think this is because the WndProc remains executable after the End button....
'
'v1.04 20020701 Added a couple of assembler optimizations to WndProc.asm
'               Zeroed lReturn before calling iSuperClass_Before
'               Fixed a few comments...........................................................
'
'v1.05 20020702 Cleaned up patching in SubClass
'               Cleaned up patching in Unsubclass
'               Re-inserted the commented out code to crash the app............................
'

Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function htonl Lib "ws2_32.dll" (ByVal hostlong As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_WNDPROC       As Long = (-4)      'Get/Set the WndProc address with GetWindowLong/SetWindowLong
Private Const BUF_TOP           As Long = 511       'Max offset in opcode buffer. Requires 136 + (# Messages * 11)
Private Const OFFSET_BEFORE     As Long = 3         'Offset into the opcode bytes for the BEFORE default processing code
Private Const OFFSET_AFTER      As Long = 65        'Offset into the opcode bytes for the AFTER default processing code
Private Const CODE_RETURN       As Long = &H10C2C9  'Leave-return opcode sequence
Private Const OPCODE_CMP_EAX    As String = "3D"    'Opcode for cmp eax,????????
Private Const OPCODE_JE         As String = "0F84"  'Opcode for je with a 4 byte relative offset.
Private Const WNDPROC_ALL       As String = "558BEC83C4FCFF7514FF7510FF750CFF7508E8wnd_proc8945FCFF7514FF7510FF750CFF75088D45FC50B8ptrOwner8BC88B0950FF511C8B45FCC9C21000"
Private Const WNDPROC_FILTERED  As String = "558BEC83C4F8EB6AC745FC000000008D45FC50C745F8000000008D45F850B8ptrOwner8BC88B0950FF5120837DF800753AFF7514FF7510FF750CFF7508E8wnd_procC9C21000E8wnd_proc8945FCFF7514FF7510FF750CFF75088D45FC50B8ptrOwner8BC88B0950FF511C8B45FCC9C210008B450CFF7514FF751050FF7508"
Private Const MSG_UNHANDLED     As String = "E8wnd_procC9C21000"

Private Type tCode
  Buf(0 To BUF_TOP) As Byte       'Opcode buffer
End Type

Private Type tCodeBuf
  Code              As tCode      'WndProc opcodes
  nBuf              As Long       'Opcode buffer index
End Type

Private All         As Boolean    'All messages?
Private Running     As Boolean    'Is the subclasser running?
Private hWnd        As Long       'Window being subclassed
Private WndProcPrev As Long       'The address of the existing WndProc
Private pCode       As Long       'Pointer to the WndProc opcode buffer
Private CodeBuf     As tCodeBuf   'Opcode buffer

'Add a message to those that will call back either before or after the existing WndProc.
Public Sub AddMsg(MsgNum As Long, Optional Before As Boolean = False)
  Debug.Assert (Running = False)                        'You don't add messages whilst the subclasser is running
  
  With CodeBuf
    If .nBuf = 0 Then                                   'If the buffer is empty (first msg to be added)
    
      Call AddCode(WNDPROC_FILTERED)                    'Add the filtered mode WndProc opcodes
    End If
    
    Call AddCode(OPCODE_CMP_EAX & Hex8(htonl(MsgNum)))  'Add the opcodes to compare the MsgNum
    
    'Add the opcodes to jump if matched
    Call AddCode(OPCODE_JE & Hex8(htonl(Not (.nBuf - IIf(Before, OFFSET_BEFORE, OFFSET_AFTER)))))
  End With
End Sub

'Subclass the passed window handle.
Public Sub Subclass(hWndSub As Long, Owner As ISuperClass, Optional AllMsgs As Boolean = False)
  Dim pOwner  As Long                                   'Object address of the owner
  Dim nPos    As Long                                   'Buf pos temporary
  
  All = AllMsgs
  
  With CodeBuf
    Debug.Assert (Running = False)                      'Subclasser already running
    Debug.Assert (IsWindow(hWndSub))                    'Invalid hWnd
    Debug.Assert (Not All And .nBuf > 0) Or _
                 (All And .nBuf = 0)                    'Either filtered mode but no messages added OR All message mode but messages added.
    hWnd = hWndSub                                      'Save the window handle
    WndProcPrev = GetWindowLong(hWnd, GWL_WNDPROC)      'Save the address of the current WndProc
    pOwner = ObjPtr(Owner)                              'Get the address of the owner
    pCode = VarPtr(.Code.Buf(0))                        'Get the address of our WndProc code
    
    If AllMsgs Then
    
      Call AddCode(WNDPROC_ALL)                         'Add the All messages WndProc opcodes
      Call PatchOffset(19)                              'Patch the WndProcPrev call
      Call PatchValue(43, pOwner)                       'Patch the owner
    Else
    
      Call PatchValue(31, pOwner)                       'Patch the owner
      Call PatchOffset(62)                              'Patch the BEFORE WndProcPrev call
      Call PatchOffset(71)                              'Patch the AFTER WndProcPrev call
      Call PatchValue(95, pOwner)                       'Patch the owner
      
      nPos = .nBuf + 1                                  'Save the buf pos
      Call AddCode(MSG_UNHANDLED)                       'Add the trailing unhandled WndProcPrev call
      Call PatchOffset(nPos)                            'Patch the WndProcPrev call
    End If
  End With
  
  'Debug support: uncomment the line below to crash the application which will (assuming VS is setup correctly)
  'allow you into the VS debugger where you can examine the generated opcodes and trace execution.
  'Don't call the Crash routine inside the IDE :)
  '
  'Call Crash
  
  Call SetWindowLong(hWnd, GWL_WNDPROC, pCode)          'Set our WndProc in place of the original
  Running = True
End Sub

'Unsubclass the window
Public Sub UnSubclass()
  If Running Then
    If All Then
      
      Call PatchValue(23, CODE_RETURN)                  'Patch a Leave-Return after default processing and before iSuperClass_After
    Else
    
      CodeBuf.Code.Buf(7) = &H29                        'Patch the WndProc entrance to jump to default processing JIC
      Call PatchValue(75, CODE_RETURN)                  'Patch a Leave-Return after default processing and before iSuperClass_After
    End If
    
    Call SetWindowLong(hWnd, GWL_WNDPROC, WndProcPrev)  'Restore the previous WndProc
    CodeBuf.nBuf = 0                                    'Reset the opcode buffer
    Running = False                                     'Not running
  End If
End Sub

Private Sub Class_Terminate()
  If Running Then UnSubclass                            'Unsubclass if the Subclasser is running
End Sub

'Translate the passed hex string character pairs to bytes and stuff into the opcode buffer.
Private Sub AddCode(sOps As String)
  Dim i As Long
  Dim j As Long
  
  With CodeBuf
    j = Len(sOps)                                       'Get length of opcode string
    Debug.Assert (.nBuf + (j \ 2) <= BUF_TOP)           'Opcode buffer overflow, increase value of BUF_TOP
    
    For i = 1 To j Step 2                               'For each pair of hex chars
      
      .Code.Buf(.nBuf) = Val("&H" & Mid$(sOps, i, 2))   'Convert from hex to byte, add to buffer at index
      .nBuf = .nBuf + 1                                 'Bump the opcode buffer index
    Next i
  End With
End Sub

'Return an 8 character hex representation of the passed 32 bit value
Private Function Hex8(lValue As Long) As String
  Dim s As String
  
  s = Hex$(lValue)
  Hex8 = String$(8 - Len(s), "0") & s
End Function

'Patch the passed code buffer offset with the passed value
Private Sub PatchValue(nOffset As Long, nValue As Long)
  Call CopyMemory(ByVal (pCode + nOffset), nValue, 4)
End Sub

'Patch the passed code buffer offset with the relative offset to the previous WndProc
Private Sub PatchOffset(nOffset As Long)
  Call CopyMemory(ByVal (pCode + nOffset), WndProcPrev - pCode - nOffset - 4, 4)
End Sub

'Debug Support:
'
'Crash the app allowing us into the debugger to examine opcodes
'Private Sub Crash()
'  Dim bCrash As Boolean
'
'  bCrash = True
'  If bCrash Then Call CopyMemory(ByVal 0, &HFFFFFFFF, 1)
'End Sub
