VERSION 5.00
Begin VB.Form frmLaunchPad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Launch Pad (VB6)"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   3135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCustomers 
      Caption         =   "Customers (.NET)"
      Height          =   1095
      Left            =   360
      TabIndex        =   2
      Top             =   2760
      Width           =   2370
   End
   Begin VB.CommandButton cmdOrders 
      Caption         =   "Orders (Hybrid)"
      Height          =   1095
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   2370
   End
   Begin VB.CommandButton cmdProducts 
      Caption         =   "Products (VB6)"
      Height          =   1095
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   2370
   End
End
Attribute VB_Name = "frmLaunchPad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Declare member level product list form.
Dim orders As frmOrderList
Attribute orders.VB_VarHelpID = -1

' Declare member level product list form.
Dim products As frmProductList

' Declare member level product list form.
Dim customers As CustomerList

Dim WithEvents eventMessenger As InteropEventMessenger
Attribute eventMessenger.VB_VarHelpID = -1

' Shows Customer List form.
Private Sub cmdCustomers_Click()
    If customers Is Nothing Then
        ' The CustomerList has not yet been created.
        Set customers = New CustomerList
        customers.Show vbModeless
    ElseIf customers.IsFormDisposed Then
        ' The CustomerList was created but closed by the user.
        ' Note that this is different than VB6.
        Set customers = New CustomerList
        customers.Show vbModeless
    Else
        ' Instance of CustomerList already exists and is still open.
        customers.ZOrder ZOrderConstants.vbBringToFront
    End If
End Sub

' Shows Order List form.
Private Sub cmdOrders_Click()
    If orders Is Nothing Then
        Set orders = New frmOrderList
        orders.Show vbModeless
    Else
        ' If instance of OrderList already exists:
        orders.Visible = True
        orders.ZOrder ZOrderConstants.vbBringToFront
    End If
End Sub

' Shows Product List form.
Private Sub cmdProducts_Click()
    If products Is Nothing Then
        Set products = New frmProductList
        products.Show vbModeless
    Else
        ' If instance of ProductList already exists:
        products.Visible = True
        products.ZOrder ZOrderConstants.vbBringToFront
    End If
End Sub







' Handle the custom application event.
Private Sub eventMessenger_ApplicationEventRaised(ByVal eventName As String, ByVal eventArgs As Variant)
    Select Case eventName
        Case "CRITICAL_ERROR"
            ' Show the error code which is passed in the eventArgs for this type of event.
            MsgBox "Error Code: " & eventArgs, vbExclamation, "Critical Error"
            ' Shut down the application.
            Unload Me
    End Select
End Sub


Private Sub Form_Load()

    ' Instantiate the Toolbox.
    Set g_InteropToolbox = New InteropToolbox
    
    ' Call Initialize method only when first creating the toolbox
    ' This aids in the debugging experience
    g_InteropToolbox.Initialize
    
    ' Set the eventMessenger member variable to the
    ' EventMessenger property from the toolbox
    ' to be able to handle events
    Set eventMessenger = g_InteropToolbox.eventMessenger
    
    ' Signal Application Startup.
    g_InteropToolbox.eventMessenger.RaiseApplicationStartedupEvent
    
    ' Get the user's credentials.
    Dim login As New frmLogin
    login.Show vbModal
    
    ' Shutdown if login failed.
    If Not login.LoginSucceeded Then
        MsgBox "Login Failed.  Application will exit."
        Unload Me
    Else
        ' Get Authentication Token.
        Dim authToken As String
        authToken = GetAuthenticationToken(login.UserID, login.Password)
        ' Add the token to the Interop Globals collection so that
        ' .NET components can use it.
        g_InteropToolbox.Globals.Add AUTHENTICATION_TOKEN, authToken
    End If
End Sub

' This method simulates getting a token used for authentication
' needed for retrieving data.
Private Function GetAuthenticationToken(UserID As String, Password As String) As String
    GetAuthenticationToken = UserID & "|" & Password
End Function


Private Sub Form_Unload(Cancel As Integer)
    
    ' Unload all VB6 Forms.
    Dim f As Form
    For Each f In Forms
        Unload f
    Next


    ' Signal Application Shutdown.
    g_InteropToolbox.eventMessenger.RaiseApplicationShutdownEvent
    
    Set eventMessenger = Nothing
    
End Sub




