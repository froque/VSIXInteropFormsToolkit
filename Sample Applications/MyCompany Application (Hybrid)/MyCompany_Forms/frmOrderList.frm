VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmOrderList 
   Caption         =   "Order List (VB6)"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10485
   Icon            =   "frmOrderList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   10485
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4575
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   8070
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1.5
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton btnDetails 
      Caption         =   "Details"
      Default         =   -1  'True
      Height          =   375
      Left            =   9120
      TabIndex        =   3
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton btnDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton btnEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   1740
      TabIndex        =   1
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton btnAdd 
      Caption         =   "Add..."
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   4800
      Width           =   975
   End
End
Attribute VB_Name = "frmOrderList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Declare member level product list form - needed to handle events.
Dim WithEvents orderDetails As orderDetails
Attribute orderDetails.VB_VarHelpID = -1

' Click event handler for "Add" button.
Private Sub btnAdd_Click()
    MsgBox "Not Implemented in Sample Application.", , "Not Implemented"
End Sub

' Click event handler for "Delete" button.
Private Sub btnDelete_Click()
    MsgBox "Not Implemented in Sample Application.", , "Not Implemented"
End Sub

' Click event handler for "Details" button.
Private Sub btnDetails_Click()
    ShowOrderDetail
End Sub

' Click event handler for "Edit" button.
Private Sub btnEdit_Click()
    MsgBox "Not Implemented in Sample Application.", , "Not Implemented"
End Sub

' Double-Click event handler for DataGrid1.
Private Sub DataGrid1_DblClick()
    ShowOrderDetail
End Sub

Private Sub ShowOrderDetail()
    ' Retrieves orderID from selected row.
    Dim orderID As String
    Dim rs As Recordset
    Set rs = Me.DataGrid1.DataSource
    orderID = rs(0).Value
    
    ' Show OrderDetails form.
    If orderDetails Is Nothing Then
        ' The OrderDetails has not yet been created.
        Set orderDetails = New orderDetails
        orderDetails.Initialize orderID
        orderDetails.Show vbModeless
    ElseIf orderDetails.IsFormDisposed Then
        ' The OrderDetails was created but closed by the user.
        ' Note that this is different than VB6.
        Set orderDetails = New orderDetails
        orderDetails.Initialize orderID
        orderDetails.Show vbModeless
    Else
        ' Instance of OrderDetails already exists and is still open.
        orderDetails.SetOrder orderID
        orderDetails.Visible = True
        orderDetails.ZOrder ZOrderConstants.vbBringToFront
    End If
End Sub
' Sets properties of DataGrid.
Private Sub Form_Load()
    ' Binds DataGrid1 to OrderDataManager.GetAllOrders.
    Dim odm As New OrderDataManager
    Set DataGrid1.DataSource = odm.GetAllOrders
    
    ' Sets column properties (header, width, etc.)
    DataGrid1.Columns(0).Caption = "Order ID"
    DataGrid1.Columns(1).Caption = "Product #"
    DataGrid1.Columns(2).Caption = "Order Date"
    DataGrid1.Columns(3).Caption = "Quantity"
    DataGrid1.Columns(4).Caption = "Customer ID"
    DataGrid1.Columns(5).Caption = "Paid"
    DataGrid1.RowHeight = 350
    DataGrid1.Columns(0).Width = 1600
    DataGrid1.Columns(1).Width = 1600
    DataGrid1.Columns(2).Width = 1600
    DataGrid1.Columns(3).Width = 1600
    DataGrid1.Columns(4).Width = 1600
    DataGrid1.Columns(5).Width = 1600
End Sub

' Shows new instance of Customer Details Form.
Private Sub orderDetails_CustomerDetailsRequested(ByVal customerID As String)
    Dim custDetails As New CustomerDetails
    custDetails.Initialize customerID
    custDetails.Show vbModal
End Sub

' Shows new instance of Product Details Form.
Private Sub orderDetails_ProductDetailsRequested(ByVal productID As String)
    Dim prodDetails As New frmProductDetail
    prodDetails.Initialize productID
    prodDetails.Show vbModal
End Sub
