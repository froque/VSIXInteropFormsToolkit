Public Class CustomerDataManager

#Region "Private Variables"

    'empty dataset from schema in dsCustomers.xsd
    Private _customersData As dsCustomers = Nothing

#End Region

#Region "Public Methods"

    ''' <summary>
    ''' Returns dataset containing all customers.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetAllCustomers() As dsCustomers
        Return GetCustomerData()
    End Function

    ''' <summary>
    ''' Returns record for requested customer.
    ''' </summary>
    ''' <param name="customerID">CustomerID of requested customer.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetCustomerByID(ByVal customerID As String) As dsCustomers.tblCustomersRow
        Return GetCustomerData().tblCustomers.FindByCustomerID(customerID)
    End Function

    ''' <summary>
    ''' Populates and returns dataset containing all customers.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetCustomerData() As dsCustomers
        'creates new instance of customers dataset
        _customersData = New dsCustomers

        'populates _customersdata with (imaginary) data
        _customersData.tblCustomers.AddtblCustomersRow("John", "Doe", "11/12/05", "555-555-5555", "44444", "10384")
        _customersData.tblCustomers.AddtblCustomersRow("Jane", "Doe", "02/27/06", "444-444-4444", "12345", "10465")
        _customersData.tblCustomers.AddtblCustomersRow("Jim", "Public", "05/01/06", "123-455-7890", "67890", "11648")
        _customersData.tblCustomers.AddtblCustomersRow("Sally", "Consumer", "06/02/03", "098-765-4321", "83502", "19985")
        _customersData.tblCustomers.AddtblCustomersRow("Tom", "Dove", "11/11/05", "123-545-2983", "23543", "12345")
        _customersData.tblCustomers.AddtblCustomersRow("Dorothy", "Lark", "03/24/06", "677-326-3333", "16436", "28743")
        _customersData.tblCustomers.AddtblCustomersRow("Sam", "Sparrow", "05/01/06", "987-455-3567", "67890", "83745")
        _customersData.tblCustomers.AddtblCustomersRow("Barb", "Crow", "06/01/03", "245-245-2466", "83502", "99987")
        _customersData.tblCustomers.AddtblCustomersRow("Tom", "Filmore", "02/27/06", "235-444-8787", "12345", "73645")
        _customersData.tblCustomers.AddtblCustomersRow("Julie", "Jones", "05/01/06", "254-455-7890", "67890", "92864")
        _customersData.tblCustomers.AddtblCustomersRow("Katie", "Shawl", "06/02/03", "856-765-4321", "83502", "63523")
        _customersData.tblCustomers.AddtblCustomersRow("Brenna", "Mannion", "11/11/05", "123-545-2983", "23543", "01928")
        _customersData.tblCustomers.AddtblCustomersRow("Rick", "Rock", "03/24/06", "677-326-3333", "16436", "37462")
        _customersData.tblCustomers.AddtblCustomersRow("Brian", "Bear", "05/01/06", "987-455-3567", "67890", "83746")
        _customersData.tblCustomers.AddtblCustomersRow("Kevin", "Wolf", "06/01/03", "245-245-2466", "83502", "10098")
        _customersData.tblCustomers.AddtblCustomersRow("John", "Smith", "11/12/05", "555-555-5555", "44444", "16323")
        Return _customersData
    End Function

#End Region

#Region "Constructors"

    ''' <summary>
    ''' Instantiates a new CustomerDataManager which is responsible for 
    ''' retrieving and updating customer related data
    ''' </summary>
    ''' <param name="authenticationToken">Token used for use to authenticate when getting data</param>
    ''' <remarks>
    ''' The authenticationToken is used to show a use case where data may need to be 
    ''' stored in VB6 and retrieved in .NET
    ''' </remarks>
    Public Sub New(ByVal authenticationToken As String)
        ' Store fictitious authentication token
    End Sub

#End Region

End Class
