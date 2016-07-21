Imports Microsoft.InteropFormTools
Imports MyCompany.Components.Customers

<InteropForm()> _
Public Class CustomerDetails

#Region "Constructors"

    Public Sub New()
        InitializeComponent()
    End Sub

    ''' <summary>
    ''' Creates a new CustomerDetails form displaying the customer 
    ''' specified by the given customer ID.
    ''' </summary>
    ''' <param name="customerID">ID of customer to display.</param>
    ''' <remarks></remarks>
    <InteropFormInitializer()> _
    Public Sub New(ByVal customerID As String)
        InitializeComponent()

        Dim authToken As String = My.InteropToolbox.Globals("AUTHENTICATION_TOKEN")

        ' Creates new instance of CustomerDataManager.
        Dim customerManager As New CustomerDataManager(authToken)

        ' Creates new instance of tblCustomersRow to act as Customer Object.
        Dim customer As dsCustomers.tblCustomersRow

        ' Gets customer info from CustomerDataManager.
        customer = customerManager.GetCustomerByID(customerID)

        ' Sets control values from customer info.
        txtFirstName.Text = customer.FirstName
        txtLastName.Text = customer.LastName
        txtSale.Text = customer.LastSale
        txtPhone.Text = customer.Phone
        txtZIP.Text = customer.ZIP

        'Sets Form text = "Customer #" & customerID & " Detail (.NET)"
        Me.Text = String.Format(My.Resources.CustomerDetailsFormText, customerID)
    End Sub

#End Region

#Region "Event Handlers"

    ''' <summary>
    ''' Handles "OK" button on Form.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Me.Close()
    End Sub

#End Region

End Class
