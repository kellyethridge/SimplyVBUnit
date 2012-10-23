VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Anchor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
' Module: SComponent.Anchor
'
Option Explicit

Private Const DEF_CAPACITY As Long = 16

' Error codes
Public Enum AnchorExceptionCodesEnum
   CapacityTooSmall = vbObjectError + 1000
   InvalidControlReference
   IndexOutOfRange
End Enum


' Sides of a control to be anchored. Both  and
' Perctage anchors can be mixed.
Public Enum AnchorsEnum
   ' These are used to anchor the control an exact
   ' amount from each edge of the container.
   ToLeft = &H1
   ToRight = &H2
   ToTop = &H4
   ToBottom = &H8
   ToAllSides = ToLeft Or ToRight Or ToTop Or ToBottom
End Enum



Private Enum ResizeMethodEnum
   Untested
   Unknown
   MoveMethod
   X1Y1X2Y2Method
   IndividualAttributesMethod
End Enum



' We use a mock control to increase performance with calculations
' by not using late-binding to retrieve the values when needed.
Private Type ControlType
   Left As Single
   Top As Single
   Width As Single
   Height As Single
   Control As Object
End Type


' Types to help organize each control's specifications
Private Type ContainerSizeType
   Width As Single
   Height As Single
End Type
Private Type AnchorOffsetsType
   Left As Long
   Right As Long
   Top As Long
   Bottom As Long
End Type
Private Type AnchorType
   Control As ControlType
   Container As Object
   Anchors As AnchorsEnum
   Offsets As AnchorOffsetsType
   ResizeMethod As ResizeMethodEnum
End Type



' Fields
Private mAnchors() As AnchorType
Private mCount As Long



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Business Properties and Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Add(ByVal Control As Control, _
                    ByVal Anchors As AnchorsEnum) As Long
   Dim Count As Long
   
   If Control Is Nothing Then ThrowException InvalidControlReference, "Control cannot be Nothing."
   
   Count = IncrementCount
   EnsureCapacity Count
   mAnchors(Count - 1) = NewAnchor(Control, Anchors)
   Add = Count - 1
End Function



Public Property Get Capacity() As Long
   Capacity = UBound(mAnchors) - 1
End Property
Public Property Let Capacity(ByVal Value As Long)
   If Value < DEF_CAPACITY Then Value = DEF_CAPACITY
   If Value < Count Then ThrowException CapacityTooSmall, "Cannot set capacity less than the number of anchors."
   ReDim Preserve mAnchors(0 To Value - 1)
End Property



Public Property Get Count() As Long
   Count = mCount
End Property



Public Sub ReAnchor()
   Dim i As Long
   
   On Error Resume Next
   For i = 0 To Count - 1
      ReAnchorControl mAnchors(i)
   Next i
End Sub



Public Sub Clear()
   Erase mAnchors
   mCount = 0
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnsureCapacity(ByVal RequiredCapacity As Long)
   Dim NewCapacity As Long
   
   If NeedsMoreCapacity(RequiredCapacity) Then
      NewCapacity = Capacity * 2
      If RequiredCapacity > NewCapacity Then NewCapacity = RequiredCapacity
      Capacity = NewCapacity
   End If
End Sub



Private Function NeedsMoreCapacity(ByVal RequiredCapacity As Long) As Boolean
   NeedsMoreCapacity = (RequiredCapacity > Capacity)
End Function



Private Function NewAnchorOffsets(ByRef Control As ControlType) As AnchorOffsetsType
   With NewAnchorOffsets
      .Left = Control.Left
      .Top = Control.Top
      .Right = Control.Control.Container.Width - (.Left + Control.Width)
      .Bottom = Control.Control.Container.Height - (.Top + Control.Height)
   End With
End Function



Private Function NewControl(ByVal Control As Object) As ControlType
   With NewControl
      On Error GoTo tryX1Y1X2Y2
      .Left = Control.Left
      .Top = Control.Top
      .Width = Control.Width
      .Height = Control.Height
      GoTo cleanExit
      
tryX1Y1X2Y2:
      .Left = Control.X1
      .Top = Control.Y1
      .Width = Control.X2 - .Left
      .Height = Control.Y2 - .Top
   
cleanExit:
      Set .Control = Control
      
   End With
End Function



Private Function NewAnchor(ByVal Control As Object, _
                           ByVal AnchorSides As AnchorsEnum) As AnchorType
   With NewAnchor
      .Anchors = AnchorSides
      .Control = NewControl(Control)
      .Offsets = NewAnchorOffsets(.Control)
      .ResizeMethod = AttemptResize(.Control)
      Set .Container = Control.Container
   End With
End Function



Private Sub ThrowException(ByVal Code As AnchorExceptionCodesEnum, _
                  Optional ByVal Message As String, _
                  Optional ByVal Source As String = "Anchor")
   Err.Raise Code, Source, Message
End Sub



Private Function IncrementCount() As Long
   mCount = mCount + 1
   IncrementCount = mCount
End Function



Private Sub ReAnchorControl(ByRef Anchor As AnchorType)
   With Anchor
      If .ResizeMethod = Unknown Then Exit Sub
      
      If .Anchors And ToBottom Then
         If .Anchors And ToTop Then
            .Control.Height = .Container.Height - (.Offsets.Top + .Offsets.Bottom)
         Else
            .Control.Top = .Container.Height - (.Offsets.Bottom + .Control.Height)
         End If
      End If
   
      If .Anchors And ToRight Then
         If .Anchors And ToLeft Then
            .Control.Width = .Container.Width - (.Offsets.Right + .Offsets.Left)
         Else
            .Control.Left = .Container.Width - (.Offsets.Right + .Control.Width)
         End If
      End If
      
      With .Control
         Select Case Anchor.ResizeMethod
            Case ResizeMethodEnum.MoveMethod
               .Control.Move .Left, .Top, .Width, .Height
            
            Case ResizeMethodEnum.X1Y1X2Y2Method
               .Control.X1 = .Left
               .Control.Y1 = .Top
               .Control.X2 = .Left + .Width
               .Control.Y2 = .Top + .Height
            
            Case ResizeMethodEnum.IndividualAttributesMethod
               .Control.Left = .Left
               .Control.Top = .Top
               .Control.Width = .Width
               .Control.Height = .Height
         End Select
      End With
   End With
End Sub



Private Function AttemptResize(ByRef Control As ControlType) As ResizeMethodEnum
   On Error Resume Next
   With Control
      .Control.Move .Left, .Top, .Width, .Height
      If Err.Number = 0 Then
         AttemptResize = MoveMethod
         Exit Function
      End If
      
      Err.Clear
      .Control.X1 = .Left
      .Control.Y1 = .Top
      .Control.X2 = .Left + .Width
      .Control.Y2 = .Top + .Height
      If Err.Number = 0 Then
         AttemptResize = X1Y1X2Y2Method
         Exit Function
      End If
      
      Err.Clear
      .Control.Left = .Left
      .Control.Top = .Top
      .Control.Width = .Width
      If Err.Number = 0 Then
         AttemptResize = IndividualAttributesMethod
         .Control.Height = .Height
         Exit Function
      End If
      
      AttemptResize = Unknown
   End With
End Function



Private Sub InitCapacity()
   Capacity = DEF_CAPACITY
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
   InitCapacity
End Sub
