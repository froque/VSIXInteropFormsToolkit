<ComClass(VB6Toolbar.ClassId, VB6Toolbar.InterfaceId, VB6Toolbar.EventsId)> _
Public Class VB6Toolbar

#Region "VB6 Interop Code"

#If COM_INTEROP_ENABLED Then

#Region "COM Registration"

    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.

    Public Const ClassId As String = "0f19177f-929a-4147-b9e4-58910c127c50"
    Public Const InterfaceId As String = "320a922c-d577-460f-bb6e-4a85d62059ae"
    Public Const EventsId As String = "68367253-4add-4385-96a9-8fbbdc6cb7a0"

    'These routines perform the additional COM registration needed by ActiveX controls
    <EditorBrowsable(EditorBrowsableState.Never)> _
    <ComRegisterFunction()> _
    Private Shared Sub Register(ByVal t As Type)
        ComRegistration.RegisterControl(t)
    End Sub

    <EditorBrowsable(EditorBrowsableState.Never)> _
    <ComUnregisterFunction()> _
    Private Shared Sub Unregister(ByVal t As Type)
        ComRegistration.UnregisterControl(t)
    End Sub

#End Region

#Region "VB6 Events"

    'This section shows some examples of exposing a UserControl's events to VB6.  Typically, you just
    '1) Declare the event as you want it to be shown in VB6
    '2) Raise the event in the appropriate UserControl event.

    Public Shadows Event Click() 'Event must be marked as Shadows since .NET UserControls have the same name.
    Public Event DblClick()

    Private Sub InteropUserControl_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Click
        RaiseEvent Click()
    End Sub

    Private Sub InteropUserControl_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.DoubleClick
        RaiseEvent DblClick()
    End Sub

#End Region

#Region "VB6 Properties"

    'The following are examples of how to expose typical form properties to VB6.  
    'You can also use these as examples on how to add additional properties.

    'Must Shadow this property as it exists in Windows.Forms and is not overridable
    Public Shadows Property Visible() As Boolean
        Get
            Return MyBase.Visible
        End Get
        Set(ByVal value As Boolean)
            MyBase.Visible = value
        End Set
    End Property

    Public Shadows Property Enabled() As Boolean
        Get
            Return MyBase.Enabled
        End Get
        Set(ByVal value As Boolean)
            MyBase.Enabled = value
        End Set
    End Property

    Public Shadows Property ForegroundColor() As Integer
        Get
            Return ActiveXControlHelpers.GetOleColorFromColor(MyBase.ForeColor)
        End Get
        Set(ByVal value As Integer)
            MyBase.ForeColor = ActiveXControlHelpers.GetColorFromOleColor(value)
        End Set
    End Property

    Public Shadows Property BackgroundColor() As Integer
        Get
            Return ActiveXControlHelpers.GetOleColorFromColor(MyBase.BackColor)
        End Get
        Set(ByVal value As Integer)
            MyBase.BackColor = ActiveXControlHelpers.GetColorFromOleColor(value)
        End Set
    End Property

    Public Overrides Property BackgroundImage() As System.Drawing.Image
        Get
            Return Nothing
        End Get
        Set(ByVal value As System.Drawing.Image)
            If value IsNot Nothing Then
                MsgBox("Setting the background image of an Interop UserControl is not supported, please use a PictureBox instead.", MsgBoxStyle.Information)
            End If
            MyBase.BackgroundImage = Nothing
        End Set
    End Property

#End Region

#Region "VB6 Methods"

    Public Overrides Sub Refresh()
        MyBase.Refresh()
    End Sub

    'Ensures that tabbing across VB6 and .NET controls works as expected
    Private Sub UserControl_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.LostFocus
        ActiveXControlHelpers.HandleFocus(Me)
    End Sub

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        'Raise Load event
        Me.OnCreateControl()
    End Sub

    <SecurityPermission(SecurityAction.LinkDemand, Flags:=SecurityPermissionFlag.UnmanagedCode)> _
    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)

        Const WM_SETFOCUS As Integer = &H7
        Const WM_PARENTNOTIFY As Integer = &H210
        Const WM_DESTROY As Integer = &H2
        Const WM_LBUTTONDOWN As Integer = &H201
        Const WM_RBUTTONDOWN As Integer = &H204

        If m.Msg = WM_SETFOCUS Then
            'Raise Enter event
            Me.OnEnter(New System.EventArgs) 

        ElseIf m.Msg = WM_PARENTNOTIFY AndAlso _
            (m.WParam.ToInt32 = WM_LBUTTONDOWN OrElse _
             m.WParam.ToInt32 = WM_RBUTTONDOWN) Then

            If Not Me.ContainsFocus Then
                'Raise Enter event
                Me.OnEnter(New System.EventArgs) 
            End If

        ElseIf m.Msg = WM_DESTROY AndAlso Not Me.IsDisposed AndAlso Not Me.Disposing Then
            'Used to ensure that VB6 will cleanup control properly
            Me.Dispose()
        End If

        MyBase.WndProc(m)
    End Sub

    'This event will hook up the necessary handlers
    Private Sub InteropUserControl_ControlAdded(ByVal sender As Object, ByVal e As ControlEventArgs) Handles Me.ControlAdded
        ActiveXControlHelpers.WireUpHandlers(e.Control, AddressOf ValidationHandler)
    End Sub

    'Ensures that the Validating and Validated events fire appropriately
    Friend Sub ValidationHandler(ByVal sender As Object, ByVal e As EventArgs)

        If Me.ContainsFocus Then Return

        'Raise Leave event
        Me.OnLeave(e) 

        If Me.CausesValidation Then
            Dim validationArgs As New CancelEventArgs
            Me.OnValidating(validationArgs)

            If validationArgs.Cancel AndAlso Me.ActiveControl IsNot Nothing Then
                Me.ActiveControl.Focus()
            Else
                'Raise Validated event
                Me.OnValidated(e) 
            End If
        End If

    End Sub

#End Region

#End If

#End Region

    'Please enter any new code here, below the Interop code

    'Events that can be raised to VB6
    Public Event NewDocument()
    Public Event Open()
    Public Event Save()
    Public Event SaveAs()
    Public Event Print()
    Public Event Cut()
    Public Event Copy()
    Public Event Paste()
    Public Event Help()
    Public Event SelectAll()
    Public Event ExitApplication(ByVal saveBeforeExit As Boolean)
    Public Event About()


    'Interop Properties - note that we don't use the <InteropProperty> attribute 
    'or click Tools->Generate, since this is a UserControl, not an InteropForm

    Public Property CutEnabled() As Boolean
        Get
            Return Me.CutToolStripMenuItem.Enabled
        End Get
        Set(ByVal value As Boolean)
            Me.CutToolStripButton.Enabled = value
            Me.CutToolStripMenuItem.Enabled = value
        End Set
    End Property

    Public Property CopyEnabled() As Boolean
        Get
            Return Me.CopyToolStripMenuItem.Enabled
        End Get
        Set(ByVal value As Boolean)
            Me.CopyToolStripButton.Enabled = value
            Me.CopyToolStripMenuItem.Enabled = value
        End Set
    End Property

    Public Property PasteEnabled() As Boolean
        Get
            Return Me.PasteToolStripMenuItem.Enabled
        End Get
        Set(ByVal value As Boolean)
            Me.PasteToolStripButton.Enabled = value
            Me.PasteToolStripMenuItem.Enabled = value
        End Set
    End Property


#Region "Toolbar/Menu Event Handlers"

    Private Sub NewToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewToolStripButton.Click, NewToolStripMenuItem.Click
        RaiseEvent NewDocument()
    End Sub

    Private Sub OpenToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenToolStripButton.Click, OpenToolStripMenuItem.Click
        RaiseEvent Open()
    End Sub

    Private Sub SaveToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripButton.Click, SaveToolStripMenuItem.Click
        RaiseEvent Save()
    End Sub

    Private Sub PrintToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintToolStripButton.Click, PrintToolStripMenuItem.Click
        RaiseEvent Print()
    End Sub

    Private Sub CutToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CutToolStripButton.Click, CutToolStripMenuItem.Click
        RaiseEvent Cut()
    End Sub

    Private Sub CopyToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopyToolStripButton.Click, CopyToolStripMenuItem.Click
        RaiseEvent Copy()
    End Sub

    Private Sub PasteToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PasteToolStripButton.Click, PasteToolStripMenuItem.Click
        RaiseEvent Paste()
    End Sub

    Private Sub HelpToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HelpToolStripButton.Click
        RaiseEvent Help()
    End Sub

    Private Sub SaveAsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveAsToolStripMenuItem.Click
        RaiseEvent SaveAs()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        'Prompt the user, and pass the decision back to VB6
        If MsgBox("Would you like to save your file first?", CType(36, MsgBoxStyle)) = MsgBoxResult.Yes Then
            RaiseEvent ExitApplication(True)
        Else
            RaiseEvent ExitApplication(False)
        End If

    End Sub

    Private Sub SelectAllToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SelectAllToolStripMenuItem.Click
        RaiseEvent SelectAll()
    End Sub

    Private Sub ContentsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ContentsToolStripMenuItem.Click
        RaiseEvent Help()
    End Sub

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        RaiseEvent About()
    End Sub

    Private Sub btnSimulateError_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSimulateError.Click
        'This method illustrates the proper pattern for handling exceptions and 
        'passing them back to VB6.  All code is wrapped in a Try/Catch/Finally block,
        'exceptions are caught, and the VB6 app is notified through the EventMessenger

        Try
            OperationThatMightThrow()
        Catch ex As ApplicationException
            My.InteropToolbox.EventMessenger.RaiseApplicationEvent("CRITICAL ERROR", ex.Message)
        Finally
            'Any cleanup code goes here
        End Try

    End Sub

    Private Sub OperationThatMightThrow()
        Throw New ApplicationException("Simulated Error")
    End Sub

#End Region

End Class
