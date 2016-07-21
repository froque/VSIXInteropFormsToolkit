Public Class OrderDataManager

#Region "Private Variables"

    ' Empty dataset from schema in dsOrders.xsd
    Private _ordersData As dsOrders = Nothing

#End Region

#Region "Public Methods"

    ''' <summary>
    ''' Returns record for requested order.
    ''' </summary>
    ''' <param name="orderID">OrderID of requested order.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetOrderByID(ByVal orderID As String) As dsOrders.tblOrdersRow
        Return GetOrderData().tblOrders.FindByOrderID(orderID)
    End Function

    ''' <summary>
    ''' Returns dataset containing all orders.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetOrderData() As dsOrders
        ' Creates new instance of Orders Dataset.
        _ordersData = New dsOrders

        ' Populates _ordersdata with (imaginary) data.
        _ordersData.tblOrders.AddtblOrdersRow("a0987", "7329", "11/12/05", "1", "10465", "Y")
        _ordersData.tblOrders.AddtblOrdersRow("b8273", "6354", "11/13/05", "3", "11648", "Y")
        _ordersData.tblOrders.AddtblOrdersRow("b2394", "1234", "09/04/05", "5", "11648", "Y")
        _ordersData.tblOrders.AddtblOrdersRow("b2322", "1234", "01/02/06", "19", "19985", "Y")
        _ordersData.tblOrders.AddtblOrdersRow("b1234", "0798", "05/12/06", "37", "12345", "N")
        _ordersData.tblOrders.AddtblOrdersRow("a8473", "2534", "03/22/03", "16", "28743", "Y")
        _ordersData.tblOrders.AddtblOrdersRow("a9827", "6453", "05/29/02", "22", "83745", "Y")
        _ordersData.tblOrders.AddtblOrdersRow("c2384", "6453", "06/14/05", "1", "99987", "Y")
        _ordersData.tblOrders.AddtblOrdersRow("c1029", "2760", "07/04/05", "64", "16323", "Y")
        _ordersData.tblOrders.AddtblOrdersRow("b2323", "0897", "12/08/05", "34", "73645", "Y")
        _ordersData.tblOrders.AddtblOrdersRow("c4353", "2760", "10/12/04", "27", "92864", "Y")
        _ordersData.tblOrders.AddtblOrdersRow("c2367", "0897", "12/30/05", "50", "63523", "Y")
        Return _ordersData
    End Function

#End Region

#Region "Constructors"

    ''' <summary>
    ''' Instantiates a new OrderDataManager which is responsible for 
    ''' retrieving and updating order related data.
    ''' </summary>
    ''' <param name="authenticationToken">Token used for use to authenticate when getting data.</param>
    ''' <remarks>
    ''' The authenticationToken is used to show a use case where data may need to be 
    ''' stored in VB6 and retrieved in .NET.
    ''' </remarks>
    Public Sub New(ByVal authenticationToken As String)
        ' Store fictitious authentication token.
    End Sub

#End Region

End Class