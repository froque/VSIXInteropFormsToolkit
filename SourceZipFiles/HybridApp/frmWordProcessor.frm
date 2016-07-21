VERSION 5.00
Object = "{DA3F9C1B-6C01-4365-B278-772BDA4DE98C}#1.0#0"; "mscoree.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmWordProcessor 
   Caption         =   "Word Processor Hybrid Application - (Untitled Document)"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   9780
   StartUpPosition =   3  'Windows Default
   Begin HybridAppControlsCtl.VB6Toolbar VB6Toolbar1 
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9735
      CopyEnabled     =   "False"
      Size            =   "649, 49"
      CutEnabled      =   "False"
      Enabled         =   "True"
      Object.Visible         =   "True"
      BackColor       =   "Control"
      BackgroundColor =   "-2147483633"
      AutoSize        =   "True"
      ForegroundColor =   "-2147483630"
      Location        =   "0, 0"
      PasteEnabled    =   "True"
      Object.TabIndex        =   "0"
      ForeColor       =   "ControlText"
      Name            =   "VB6Toolbar"
   End
   Begin HybridAppControlsCtl.StatusBar StatusBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   7080
      Width           =   9735
      Size            =   "649, 17"
      Enabled         =   "True"
      Object.Visible         =   "True"
      BackColor       =   "Control"
      BackgroundColor =   "-2147483633"
      AutoSize        =   "True"
      ForegroundColor =   "-2147483630"
      Location        =   "0, 472"
      Object.TabIndex        =   "0"
      ForeColor       =   "ControlText"
      Name            =   "StatusBar"
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1800
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Text documents (*.*)|*.txt"
   End
   Begin VB.TextBox txtDocument 
      Height          =   6375
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      ToolTipText     =   "Enter text here"
      Top             =   720
      Width           =   9735
   End
End
Attribute VB_Name = "frmWordProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'=======================================================================
'Important Note:   
'		In order to use this sample, you must first open the .NET solution and build it.
'		You will then be able to load and build this project.
'=======================================================================

'For an example of how to deploy the controls in this application using RegFree COM,
'please see the WordProcessor.exe.manifest file in this directory.

'Global variables
Public g_InteropToolbox As InteropToolbox
Public UserName As String
Public CurrentFileName As String

'This object will handle the events that fire on the UserControl.
'Notice that the type is HybridAppControls and not HybridAppControlsCtl.
Dim WithEvents VB6Toolbar1Events As HybridAppControls.VB6Toolbar
Attribute VB6Toolbar1Events.VB_VarHelpID = -1

Dim WithEvents eventMessenger As InteropEventMessenger
Attribute eventMessenger.VB_VarHelpID = -1

Private Sub Form_Initialize()
    'Instantiate the Toolbox
    Set g_InteropToolbox = New InteropToolbox
    
    'Call Initialize method only when first creating the toolbox
    'This aids in the debugging experience
    g_InteropToolbox.Initialize
    
    'Hook up the EventMessenger
    Set eventMessenger = g_InteropToolbox.eventMessenger
    
    'Store a global variable so the .NET controls can see it
    UserName = InputBox("Please enter your name:")
    g_InteropToolbox.Globals.Add "UserName", UserName
    
End Sub

'Handle the custom application event.
Private Sub eventMessenger_ApplicationEventRaised(ByVal eventName As String, ByVal eventArgs As Variant)
    Select Case eventName
        Case "CRITICAL ERROR"
            'Show the error which is passed in the eventArgs for this type of event
            MsgBox eventArgs & ".  This application will now shut down.", vbExclamation, "Critical Error"
            
            'Clean up any VB6 resources here
            
            'Shut down the application.
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    'This line hooks up the events.  Events should be handled by VB6Toolbar1Events
    Set VB6Toolbar1Events = VB6Toolbar1
End Sub

'Ensure that the controls resize appropriately
Private Sub Form_Resize()
    VB6Toolbar1.Width = Me.Width
    txtDocument.Width = Me.Width - 120
    txtDocument.Height = Me.Height - Me.VB6Toolbar1.Height - StatusBar1.Height - 480
    StatusBar1.Width = Me.Width - 120
    StatusBar1.Top = txtDocument.Top + txtDocument.Height
End Sub

Private Sub VB6Toolbar1Events_About()
    Dim frmAbout As New HybridAppForm.AboutBox
    frmAbout.Show vbModal
End Sub

Private Sub VB6Toolbar1Events_ExitApplication(ByVal saveBeforeExit As Boolean)
    If saveBeforeExit Then
        VB6Toolbar1Events_Save
    End If
    Unload Me
End Sub

Private Sub VB6Toolbar1Events_Help()
    MsgBox "This sample application uses the Microsoft Interop Forms Toolkit 2.0 to display a " & _
           ".NET control on a VB6 form.  If you have any questions on how to use " & _
           "the Toolkit, please post a question at the Visual Basic Interop & " & _
           "Upgrade forum at http://go.microsoft.com/fwlink/?LinkID=69260.", 64
End Sub

Private Sub VB6Toolbar1Events_NewDocument()
    txtDocument.Text = ""
    CurrentFileName = ""
    Me.Caption = "Word Processor Hybrid Application - (Untitled Document)"
End Sub

Private Sub VB6Toolbar1Events_Open()
    Dim strBuffer As String
    Dim FNum As Integer
    FNum = FreeFile
    
    CommonDialog1.ShowOpen
    
    'If the user didn't click Cancel then prompt for filename
    If CommonDialog1.FileName <> "" Then
        
        CurrentFileName = CommonDialog1.FileName
        txtDocument.Text = ""
      
        'Load the file into the TextBox
        Open CurrentFileName For Binary As #FNum
        strBuffer = Space(LOF(FNum))
        Get #FNum, , strBuffer
        txtDocument.Text = strBuffer
        
        Close #FNum
        Me.Caption = "Word Processor Hybrid Application - " & CurrentFileName
    End If
   
End Sub

Private Sub VB6Toolbar1Events_Print()
    Printer.Print txtDocument.Text
    Printer.EndDoc
End Sub

Private Sub VB6Toolbar1Events_Save()
    If CurrentFileName = "" Then
        VB6Toolbar1Events_SaveAs
    Else
        SaveFile
    End If
End Sub

Private Sub VB6Toolbar1Events_SaveAs()
    CommonDialog1.ShowSave
    If CommonDialog1.FileName <> "" Then
        CurrentFileName = CommonDialog1.FileName
    
        SaveFile
    End If
End Sub

'Writes the file to <CurrentFileName>
Private Sub SaveFile()
    Dim FNum As Integer
    FNum = FreeFile
    Open CurrentFileName For Output As #FNum
    Print #FNum, txtDocument.Text
                
    Close #FNum
    Me.Caption = "Word Processor Hybrid Application - " & CurrentFileName
End Sub


'
'All remaining methods deal with Cut, Copy, and Paste
'
Private Sub UpdateClipboardStatus()
    VB6Toolbar1.PasteEnabled = Clipboard.GetText() <> ""
    VB6Toolbar1.CutEnabled = txtDocument.SelText <> ""
    VB6Toolbar1.CopyEnabled = txtDocument.SelText <> ""
End Sub

Private Sub txtDocument_KeyUp(KeyCode As Integer, Shift As Integer)
    UpdateClipboardStatus
End Sub

Private Sub txtDocument_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UpdateClipboardStatus
End Sub

Private Sub VB6Toolbar1Events_Copy()
    SetClipboardText
End Sub

Private Sub VB6Toolbar1Events_Cut()
    SetClipboardText
    txtDocument.SelText = ""
    UpdateClipboardStatus
End Sub

Private Sub VB6Toolbar1Events_Paste()
    txtDocument.SelText = Clipboard.GetText()
    Clipboard.Clear
    UpdateClipboardStatus
End Sub

Private Sub SetClipboardText()
    Clipboard.Clear
    Clipboard.SetText txtDocument.SelText
    VB6Toolbar1.PasteEnabled = True
End Sub

Private Sub VB6Toolbar1Events_SelectAll()
    txtDocument.SelStart = 0
    txtDocument.SelLength = Len(txtDocument.Text)
    VB6Toolbar1.CutEnabled = True
    VB6Toolbar1.CopyEnabled = True
End Sub
