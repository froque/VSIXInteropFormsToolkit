VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProductDataManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Empty recordset for use in DataManager.
Private productsData As ADODB.Recordset
 
' Returns recordset containing all products.
Public Function GetAllProducts() As ADODB.Recordset
   Set GetAllProducts = GetProductData()
End Function

' Returns record for requested product.
Public Function GetProductByID(productID As String) As ADODB.Recordset
    Dim rsID As New ADODB.Recordset
    
    ' Fills recordset with all records for products.
    Set rsID = GetProductData()
    
    ' Filters recordset to reflect only requested productID.
    rsID.Filter = "ProductID = " & Join(Split(productID, ","), " OR MyField = ")
    Set GetProductByID = rsID
End Function

' Creates and returns recordset containing all products.
Public Function GetProductData() As ADODB.Recordset
    ' Instantiates productsData recordset.
    Set productsData = New ADODB.Recordset
    
    ' Defines columns for recordset.
    productsData.Fields.Append "ProductID", adVarChar, 20
    productsData.Fields.Append "Variety", adVarChar, 20
    productsData.Fields.Append "Vineyard", adVarChar, 20
    productsData.Fields.Append "State", adVarChar, 20
    productsData.Fields.Append "Vintage", adVarChar, 20
    productsData.Fields.Append "Retail", adVarChar, 20
    productsData.Open
    
    ' Populates recordset with (imaginary) data.
    productsData.AddNew Array("ProductID", "Variety", "Vineyard", "State", "Vintage", "Retail"), _
    Array("7329", "Merlot", "Marshall Valley", "CA", "1999", "$23")
    productsData.AddNew Array("ProductID", "Variety", "Vineyard", "State", "Vintage", "Retail"), _
    Array("6354", "Syrah", "Doughtery Hills", "CA", "2000", "$13")
    productsData.AddNew Array("ProductID", "Variety", "Vineyard", "State", "Vintage", "Retail"), _
    Array("1234", "Zinfandel", "Holstad Bros.", "WA", "1997", "$42")
    productsData.AddNew Array("ProductID", "Variety", "Vineyard", "State", "Vintage", "Retail"), _
    Array("8265", "Pinot Noir", "Powers", "CA", "1995", "$25")
    productsData.AddNew Array("ProductID", "Variety", "Vineyard", "State", "Vintage", "Retail"), _
    Array("2534", "Cabernet", "Beverley", "FL", "1999", "$19")
    productsData.AddNew Array("ProductID", "Variety", "Vineyard", "State", "Vintage", "Retail"), _
    Array("0798", "Merlot", "Hughes Simon", "CA", "2000", "$14")
    productsData.AddNew Array("ProductID", "Variety", "Vineyard", "State", "Vintage", "Retail"), _
    Array("6453", "Pinot Grigio", "Ravani", "WA", "2002", "$13")
    productsData.AddNew Array("ProductID", "Variety", "Vineyard", "State", "Vintage", "Retail"), _
    Array("1834", "Chardonnay", "Dougherty Hills", "OR", "2004", "$27")
    productsData.AddNew Array("ProductID", "Variety", "Vineyard", "State", "Vintage", "Retail"), _
    Array("5026", "Sauvignon Blanc", "Smith Goren", "CA", "2001", "$52")
    productsData.AddNew Array("ProductID", "Variety", "Vineyard", "State", "Vintage", "Retail"), _
    Array("2760", "Reisling", "Rauschen", "WA", "1990", "$62")
    productsData.AddNew Array("ProductID", "Variety", "Vineyard", "State", "Vintage", "Retail"), _
    Array("0897", "Zinfandel", "Gorny Creek", "WA", "2005", "$25")
    productsData.AddNew Array("ProductID", "Variety", "Vineyard", "State", "Vintage", "Retail"), _
    Array("6241", "Syrah", "Marshall Valley", "CA", "2004", "$23")
    Set GetProductData = productsData
End Function


