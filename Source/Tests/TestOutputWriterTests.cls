VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestOutputWriterTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestOutputWriterTests
'
Option Explicit

Public Sub WriteText_WhenCalled_OutputsTextToListener()
    Dim Listener As New MockEventListener
    Dim Writer As TestOutputWriter
    Set Writer = Sim.NewTestOutputWriter(Listener, ErrorOutput)
    
    Writer.WriteText "Hello {0}", "Kelly"
    
    Assert.That Listener.Output.Text, Iz.EqualTo("Hello Kelly")
    Assert.That Listener.Output.OutputType, Iz.EqualTo(TestOutputType.ErrorOutput)
End Sub

Public Sub WriteText_WithNoArgs_OutputsJustText()
    Dim Listener As New MockEventListener
    Dim Writer As TestOutputWriter
    Set Writer = Sim.NewTestOutputWriter(Listener, ErrorOutput)
    
    Writer.WriteText "Hello Kelly"
    
    Assert.That Listener.Output.Text, Iz.EqualTo("Hello Kelly")
End Sub

Public Sub WriteLine_WithNoText_OutputsNewLine()
    Dim Listener As New MockEventListener
    Dim Writer As TestOutputWriter
    Set Writer = Sim.NewTestOutputWriter(Listener, StandardOutput)
    
    Writer.WriteLine
    
    Assert.That Listener.Output.Text, Iz.EqualTo(vbCrLf)
End Sub

Public Sub WriteLine_WithTextOnly_OutputsTextWithNewLine()
    Dim Listener As New MockEventListener
    Dim Writer As TestOutputWriter
    Set Writer = Sim.NewTestOutputWriter(Listener, StandardOutput)
    
    Writer.WriteLine "hello"
    
    Assert.That Listener.Output.Text, Iz.EqualTo("hello" & vbCrLf)
End Sub

Public Sub WriteLine_WithTextAndArgs_OutputsFormattedLineWithNewLine()
    Dim Listener As New MockEventListener
    Dim Writer As TestOutputWriter
    Set Writer = Sim.NewTestOutputWriter(Listener, StandardOutput)
    
    Writer.WriteLine "Hello {0}", "Kelly"
    
    Assert.That Listener.Output.Text, Iz.EqualTo("Hello Kelly" & vbCrLf)
End Sub
