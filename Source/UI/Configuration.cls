VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UIConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'    CopyRight (c) 2008 Kelly Ethridge
'
'    This file is part of SimplyVBUnitUI.
'
'    SimplyVBUnitUI is free software; you can redistribute it and/or modify
'    it under the terms of the GNU Library General Public License as published by
'    the Free Software Foundation; either version 2.1 of the License, or
'    (at your option) any later version.
'
'    SimplyVBUnitUI is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Library General Public License for more details.
'
'    You should have received a copy of the GNU Library General Public License
'    along with Foobar; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'    Module: Configuration
'

Option Explicit

Private mPath   As String
Private mDOM    As New DOMDocument60


Public Sub Load(ByVal Path As String)
    mPath = Path
    mDOM.async = False
    
    Call mDOM.Load(mPath)
    
    If mDOM.documentElement Is Nothing Then
        Call mDOM.appendChild(mDOM.createElement("Configuration"))
        Call mDOM.documentElement.appendChild(mDOM.createElement("General"))
    End If
End Sub

Public Sub Save()
    Dim writer As New MXXMLWriter60
    writer.byteOrderMark = True
    writer.omitXMLDeclaration = True
    writer.indent = True
    
    Dim reader As New SAXXMLReader60
    Set reader.contentHandler = writer
    Set reader.dtdHandler = writer
    Set reader.errorHandler = writer
    
    Call reader.putProperty("http://xml.org/sax/properties/lexical-handler", writer)
    Call reader.putProperty("http://xml.org/sax/properties/declaration-handler", writer)
    Call reader.parse(mDOM)
    
    Dim ff As Long
    ff = FreeFile
    Open mPath For Output As #ff
    Print #ff, writer.Output
    Close #ff
End Sub

Public Function GetSettings(ByVal SectionName As String) As Collection
    Dim Result As New Collection
    Dim Element As IXMLDOMElement
    Set Element = mDOM.documentElement.selectSingleNode(SectionName)
    
    If Not Element Is Nothing Then
        Dim i As Long
        For i = 0 To Element.childNodes.Length - 1
            Dim Setting As UISetting
            Set Setting = FromElement(Element.childNodes(i))
            
            Call Result.Add(Setting, Setting.Name)
        Next i
    End If
    
    Set GetSettings = Result
End Function

Public Sub SetSettings(ByVal SectionName As String, ByVal Settings As Collection)
    Dim Element As IXMLDOMElement
    Set Element = mDOM.createElement(SectionName)
    
    Dim Setting As UISetting
    For Each Setting In Settings
        Dim Child As IXMLDOMElement
        Set Child = Element.appendChild(mDOM.createElement("Setting"))
        Call Child.setAttribute("name", Setting.Name)
        Call Child.setAttribute("value", CStr(Setting.Value))
        Call Child.setAttribute("type", TypeName(Setting.Value))
    Next Setting
    
    Dim Original As IXMLDOMElement
    Set Original = mDOM.documentElement.selectSingleNode(SectionName)
    
    If Original Is Nothing Then
        Call mDOM.documentElement.appendChild(Element)
    Else
        Call mDOM.documentElement.replaceChild(Element, Original)
    End If
End Sub

Public Sub SetValue(ByVal Name As String, ByVal Value As Variant)
    Dim General As IXMLDOMElement
    Set General = mDOM.documentElement.selectSingleNode("//General")
    
    Dim Element As IXMLDOMElement
    Set Element = General.selectSingleNode("//Setting[@name='" & Name & "']")
    
    If Element Is Nothing Then
        Set Element = General.appendChild(mDOM.createElement("Setting"))
        Call Element.setAttribute("name", Name)
    End If
    
    Call Element.setAttribute("value", CStr(Value))
    Call Element.setAttribute("type", TypeName(Value))
End Sub

Public Function GetValue(ByVal Name As String, ByVal Default As Variant) As Variant
    Dim Element As IXMLDOMElement
    Set Element = mDOM.documentElement.selectSingleNode("//General/Setting[@name='" & Name & "']")
    
    If Element Is Nothing Then
        GetValue = Default
    Else
        GetValue = FromElement(Element).Value
    End If
End Function

Public Property Get AutoRun() As Boolean
    AutoRun = GetValue("AutoRun", False)
End Property

Public Property Let AutoRun(ByVal RHS As Boolean)
    Call SetValue("AutoRun", RHS)
End Property

Public Property Get TreeViewStartUpState() As String
    TreeViewStartUpState = GetValue("TreeViewStartUpState", "Auto")
End Property

Public Property Let TreeViewStartUpState(ByVal RHS As String)
    Call SetValue("TreeViewStartUpState", RHS)
End Property

Public Property Get OutputToTextConsole() As Boolean
    OutputToTextConsole = GetValue("OutputToTextConsole", True)
End Property

Public Property Let OutputToTextConsole(ByVal RHS As Boolean)
    Call SetValue("OutputToTextConsole", RHS)
End Property

Public Property Get OutputToLogConsole() As Boolean
    OutputToLogConsole = GetValue("OutputToLogConsole", True)
End Property

Public Property Let OutputToLogConsole(ByVal RHS As Boolean)
    Call SetValue("OutputToLogConsole", RHS)
End Property

Public Property Get OutputToErrorConsole() As Boolean
    OutputToErrorConsole = GetValue("OutputToErrorConsole", True)
End Property

Public Property Let OutputToErrorConsole(ByVal RHS As Boolean)
    Call SetValue("OutputToErrorConsole", RHS)
End Property

Public Property Get AllowStop() As Boolean
    AllowStop = GetValue("AllowStop", True)
End Property

Public Property Let AllowStop(ByVal RHS As Boolean)
    Call SetValue("AllowStop", RHS)
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FromElement(ByVal Element As IXMLDOMElement) As UISetting
    Dim Value As Variant
    Value = Element.getAttribute("value")
    
    Select Case Element.getAttribute("type")
        Case "String":
        Case "Long":    Value = CLng(Value)
        Case "Integer": Value = CInt(Value)
        Case "Byte":    Value = CByte(Value)
        Case "Boolean": Value = CBool(Value)
        Case "Single":  Value = CSng(Value)
        Case "Double":  Value = CDbl(Value)
        Case "Date":    Value = CDate(Value)
        Case Else
            Call Err.Raise(5, "UIConfiguration.GetValue", "Unsupported datatype <" & Element.getAttribute("type") & ">.")
    End Select
    
    Set FromElement = UI.NewUISetting(Element.getAttribute("name"), Value)
End Function

Private Sub InitArray(ByRef Arr As Variant)
    ' this will call a UDT array to be initalized to a 0-length array.
End Sub


