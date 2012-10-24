VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Config As UIConfiguration

Public Sub Setup()
    On Error Resume Next
    Kill App.Path & "\TestConfig.xml"
    Err.Clear
    
    Set Config = New UIConfiguration
    Call Config.Load(App.Path & "\TestConfig.xml")
End Sub

Public Sub SaveEmptyConfig()
    Config.save
    
    Dim Dom As New DOMDocument
    Dom.Load App.Path & "\TestConfig.xml"
    
    Assert.That Dom.selectSingleNode("//Configuration"), Iz.Not.Nothing
    Assert.That Dom.selectSingleNode("//Configuration/General"), Iz.Not.Nothing
End Sub

Public Sub SaveValues()
    Config.setValue "AString", "Hello"
    Config.setValue "ALong", 1&
    Config.setValue "ABool", True
    Config.save

    Dim cfg As New UIConfiguration
    cfg.Load App.Path & "\TestConfig.xml"

    Assert.That cfg.getValue("AString", ""), Iz.EqualTo("Hello")
    Assert.That cfg.getValue("ALong", 0), Iz.EqualTo(1)
    Assert.That cfg.getValue("ABool", False), Iz.True
End Sub

Public Sub GetSectionValues()
    Config.setValue "One", 1
    Config.setValue "Two", "two"
    Config.setValue "Three", "3"
    Config.save
    
    Dim cfg As New UIConfiguration
    cfg.Load App.Path & "\TestConfig.xml"
    
    Dim Settings As Collection
    Set Settings = cfg.GetSettings("General")
    
    Assert.That Settings.Count, Iz.EqualTo(3)
    Assert.That Array(Settings(1).Name, Settings(2).Name, Settings(3).Name), Iz.EquivalentTo(Array("One", "Two", "Three"))
End Sub

Public Sub SetSectionValues()
    Dim Settings As New Collection
    Settings.Add UI.NewUISetting("Setting One", 1)
    Settings.Add UI.NewUISetting("Setting Two", "Two")
    Settings.Add UI.NewUISetting("Three", "a 3")
    
    Config.SetSettings "Custom", Settings
    Config.save
    
    Dim Dom As New DOMDocument
    Dom.Load App.Path & "\TestConfig.xml"
    
    Dim Custom As IXMLDOMElement
    Set Custom = Dom.selectSingleNode("//Custom")
    
    Assert.That Custom.childNodes.length, Iz.EqualTo(3)
End Sub

Public Sub GetEmptySection()
    Dim Settings As Collection
    Set Settings = Config.GetSettings("General")
    
    Assert.That Settings, Iz.Empty
End Sub
