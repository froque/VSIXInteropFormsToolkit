Imports Microsoft.InteropFormTools
Imports MyCompany.Components.Orders

<InteropForm()> _
Public Class OrderDetails

#Region "Events"

    ' Event declarations requesting Customer and Product Details.

    <InteropFormEvent()> _
    Public Event CustomerDetailsRequested As CustomerDetailRequestedEventHandler
    Public Delegate Sub CustomerDetailRequestedEventHandler(ByVal customerID As String)

    <InteropFormEvent()> _
    Public Event ProductDetailsRequested As ProductDetailRequestedEventHandler
    Public Delegate Sub ProductDetailRequestedEventHandler(ByVal productID As String)

#End Region

#Region "Constructors"

    ''' <summary>
    ''' Default constructor.  
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        InitializeComponent()
    End Sub

    ''' <summary>
    ''' Creates a new OrderDetails form displaying the order 
    ''' specified by the given order ID.
    ''' </summary>
    ''' <param name="orderID">ID of order to display.</param>
    ''' <remarks></remarks>
    <InteropFormInitializer()> _
    Public Sub New(ByVal orderID As String)
        InitializeComponent()

        SetOrder(orderID)

    End Sub

#End Region

#Region "Public Methods"

    <InteropFormMethod()> _
    Public Sub SetOrder(ByVal orderID As String)

        Dim authToken As String = My.InteropToolbox.Globals("AUTHENTICATION_TOKEN")

        ' Creates new instance of Order Data Manager.
        Dim orderManager As New OrderDataManager(authToken)

        ' Creates new instance of tblOrdersRow to act as Order Object.
        Dim [order] As dsOrders.tblOrdersRow

        ' Gets order info from Order Data Manager.
        [order] = orderManager.GetOrderByID(orderID)

        ' Sets control values from Order info.
        txtProduct.Text = [order]("Product")
        txtDate.Text = [order]("Date")
        txtQuantity.Text = [order]("Quantity")
        txtCustomer.Text = [order]("CustomerID")
        txtPaid.Text = [order]("Paid")

        'sets Form.Text = "Order #" & orderID & " Detail (.NET)"
        Me.Text = String.Format(My.Resources.OrderDetailsFormText, orderID)

    End Sub

#End Region

#Region "Event Handlers"

    ''' <summary>
    ''' Event Handler for "OK" button on Form.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Me.Close()
    End Sub

    ''' <summary>
    ''' Event Handler for "Customer Detail" LinkLabel on Form.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub lblCustomerDetail_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lblCustomerDetail.LinkClicked
        ' Raises event declared above.
        RaiseEvent CustomerDetailsRequested(txtCustomer.Text)
    End Sub

    ''' <summary>
    ''' Event Handler for "Product Detail" LinkLabel on Form.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub lblProductDetail_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lblProductDetail.LinkClicked
        ' Raises event declared above.
        RaiseEvent ProductDetailsRequested(txtProduct.Text)
    End Sub

#End Region

End Class