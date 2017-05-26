VERSION 5.00
Begin VB.Form frmProductDetail 
   Caption         =   "Product #2323 Detail (VB6)"
   ClientHeight    =   3675
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4005
   Icon            =   "frmProductDetail.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3675
   ScaleWidth      =   4005
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Product Detail"
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.TextBox txtPrice 
         Height          =   285
         Left            =   1875
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox txtVintage 
         Height          =   285
         Left            =   1875
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txtState 
         Height          =   285
         Left            =   1875
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtVineyard 
         Height          =   285
         Left            =   1875
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtVariety 
         Height          =   285
         Left            =   1875
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Retail Price:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   675
         TabIndex        =   6
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Vintage:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1035
         TabIndex        =   5
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "State:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1275
         TabIndex        =   4
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Vineyard:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1035
         TabIndex        =   3
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Variety:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1155
         TabIndex        =   2
         Top             =   480
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmProductDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Click event handler for "OK" button.
Private Sub btnOK_Click()
    Me.Hide
End Sub

' Populates controls with data for requested product.
Public Sub Initialize(productID As String)
    ' Gets requested product data.
    Dim productManager As New ProductDataManager
    Dim product As Recordset
    Set product = productManager.GetProductByID(productID)
    
    If product.RecordCount > 0 Then
        product.MoveFirst
    
        ' Binds controls to product data.
        txtVariety.Text = product(1).Value
        txtVineyard.Text = product(2).Value
        txtState.Text = product(3).Value
        txtVintage.Text = product(4).Value
        txtPrice.Text = product(5).Value
    
    End If
    
    Me.Caption = "Product #" & productID & " Detail (VB6)"

End Sub
