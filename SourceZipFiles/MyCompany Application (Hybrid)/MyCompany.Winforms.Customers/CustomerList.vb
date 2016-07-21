Imports Microsoft.InteropFormTools
Imports MyCompany.Components.Customers

<InteropForm()> _
Public Class CustomerList

#Region "Private Variables"

    ' Creates new CustomerDataManager.
    Dim _customerManager As CustomerDataManager

#End Region

#Region "Event Handlers"

    ''' <summary>
    ''' Event Handler for "Add..." button on Form.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        'displays "Not Implemented in Sample Application."
        MsgBox(My.Resources.NYIMessageText)
    End Sub

    ''' <summary>
    ''' Event Handler for "Edit" button on Form.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        'displays "Not Implemented in Sample Application."
        MsgBox(My.Resources.NYIMessageText)
    End Sub

    ''' <summary>
    ''' Event Handler for "Delete" button on Form.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        'displays "Not Implemented in Sample Application."
        MsgBox(My.Resources.NYIMessageText)
    End Sub

    ''' <summary>
    ''' Event Handler for "Customer Details" button on Form.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnDetails_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDetails.Click
        ' If no rows are selected, set the first row to selected.
        If (DataGridView1.SelectedRows.Count = 0) Then
            DataGridView1.Rows(0).Selected = True
        End If

        ' Creates new instance of Customer Details Form from selected CustomerID.
        Dim custinfo As New CustomerDetails(DataGridView1.SelectedRows(0).Cells(5).Value)
        custinfo.ShowDialog()
    End Sub

    ''' <summary>
    ''' Event Handler for Form's Load event.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub customerForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim authToken As String = My.InteropToolbox.Globals("AUTHENTICATION_TOKEN")

        ' Instantiates CustomerDataManager.
        _customerManager = New CustomerDataManager(authToken)

        ' Populates local dataset DsCustomers with rows from CustomerManager.
        Me.DsCustomers = _customerManager.GetAllCustomers()

        ' DataBinds DataGridView1 to local Dataset DsCustomers.
        DataGridView1.DataSource = Me.DsCustomers
        DataGridView1.Sort(DataGridView1.Columns(5), ComponentModel.ListSortDirection.Ascending)
    End Sub

    ''' <summary>
    ''' Event Handler for double-click event of DataGridView1.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        ' Creates new instance of Customer Details Form from selected CustomerID.
        Dim custinfo As New CustomerDetails(DataGridView1.SelectedRows(0).Cells(5).Value)
        custinfo.ShowDialog()
    End Sub

    Private Sub lnkError_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkError.LinkClicked
        ' Tell user that we're simulating a critical error.
        ' displays:          MsgBox("This will simulate a critical error occurring." & vbNewLine & "An event will be raised to notify VB6 that the application should shut down.", MsgBoxStyle.OkCancel, "Critical Error Event")
        MsgBox(String.Format(My.Resources.CriticalErrMsg, vbNewLine), MsgBoxStyle.OkCancel, My.Resources.CriticalErrMsgTitle)
        ' Raise error event passing a sample error code.
        My.InteropToolbox.EventMessenger.RaiseApplicationEvent("CRITICAL_ERROR", My.Resources.SampleErrorCode)
    End Sub

#End Region

End Class

