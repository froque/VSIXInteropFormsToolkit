VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderDataManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Empty recordset for use in DataManager.
Private ordersData As ADODB.Recordset
 
' Returns recordset containing all orders.
Public Function GetAllOrders() As ADODB.Recordset
   Set GetAllOrders = GetOrderData()
End Function

' Creates and returns recordset containing all orders.
Public Function GetOrderData() As ADODB.Recordset
    ' Instantiates ordersData recordset.
    Set ordersData = New ADODB.Recordset
    
    ' Defines columns for recordset.
    ordersData.Fields.Append "OrderID", adVarChar, 10
    ordersData.Fields.Append "Item", adVarChar, 10
    ordersData.Fields.Append "Date", adVarChar, 10
    ordersData.Fields.Append "Quantity", adVarChar, 10
    ordersData.Fields.Append "CustomerID", adVarChar, 10
    ordersData.Fields.Append "Paid", adVarChar, 10
    ordersData.Open
    
    ' Populates recordset with (imaginary) data.
    ordersData.AddNew Array("OrderID", "Item", "Date", "Quantity", "CustomerID", "Paid"), _
    Array("a0987", "7329", "11/12/05", "1", "10465", "Y")
    ordersData.AddNew Array("OrderID", "Item", "Date", "Quantity", "CustomerID", "Paid"), _
    Array("b8273", "6354", "11/13/05", "3", "11648", "Y")
    ordersData.AddNew Array("OrderID", "Item", "Date", "Quantity", "CustomerID", "Paid"), _
    Array("b2394", "1234", "09/04/05", "5", "11648", "Y")
    ordersData.AddNew Array("OrderID", "Item", "Date", "Quantity", "CustomerID", "Paid"), _
    Array("b2322", "1234", "01/02/06", "19", "19985", "Y")
    ordersData.AddNew Array("OrderID", "Item", "Date", "Quantity", "CustomerID", "Paid"), _
    Array("b1234", "2534", "05/12/06", "37", "12345", "N")
    ordersData.AddNew Array("OrderID", "Item", "Date", "Quantity", "CustomerID", "Paid"), _
    Array("a8473", "0798", "03/22/03", "16", "28743", "Y")
    ordersData.AddNew Array("OrderID", "Item", "Date", "Quantity", "CustomerID", "Paid"), _
    Array("a9827", "6453", "05/29/02", "22", "83745", "Y")
    ordersData.AddNew Array("OrderID", "Item", "Date", "Quantity", "CustomerID", "Paid"), _
    Array("c2384", "6453", "06/14/05", "1", "99987", "Y")
    ordersData.AddNew Array("OrderID", "Item", "Date", "Quantity", "CustomerID", "Paid"), _
    Array("c1029", "2760", "07/04/05", "64", "16323", "Y")
    ordersData.AddNew Array("OrderID", "Item", "Date", "Quantity", "CustomerID", "Paid"), _
    Array("b2323", "0897", "12/08/05", "34", "73645", "Y")
    ordersData.AddNew Array("OrderID", "Item", "Date", "Quantity", "CustomerID", "Paid"), _
    Array("c4353", "2760", "10/12/04", "27", "92864", "Y")
    ordersData.AddNew Array("OrderID", "Item", "Date", "Quantity", "CustomerID", "Paid"), _
    Array("c2367", "0897", "12/30/05", "50", "63523", "Y")
    Set GetOrderData = ordersData
End Function
