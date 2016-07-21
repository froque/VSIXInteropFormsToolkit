VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdLaunchNetForm 
      Caption         =   "Launch .NET Form"
      Height          =   1335
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Declare the .NET HelloWorldForm.
Dim WithEvents hello As HelloWorldForm
Attribute hello.VB_VarHelpID = -1

Private Sub cmdLaunchNetForm_Click()
    
    ' Instantiate the form.
    Set hello = New HelloWorldForm
    
    ' Call one of the Initialize methods which
    ' is calling one of the .NET
    ' parameterized constructors.
    'hello.Initialize "Hello from VB6"
    
    ' Show the form.
    hello.Show vbModeless
    
    ' Other standard Form methods and properties are available...
    'hello.Move 0, 0, 800, 800
    'hello.Caption = "Other Caption"
    
    ' Other custom methods and properties are available, as well...
    'hello.HelloText = "Other Hello Text"
    'hello.ReverseBackgroundColors

End Sub

' Handle the SampleEvent event.
Private Sub hello_SampleEvent(ByVal sampleEventText As String)
    MsgBox sampleEventText
End Sub


