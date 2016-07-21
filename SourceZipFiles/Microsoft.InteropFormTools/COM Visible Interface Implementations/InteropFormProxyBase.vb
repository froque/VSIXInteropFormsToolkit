Imports System.Drawing

Imports System.Runtime.InteropServices

''' <summary>
''' Base class for the .NET wrapper class that is generated 
''' by the Visual Studio Addin
''' </summary>
''' <remarks></remarks>
<ClassInterface(ClassInterfaceType.None)> _
Public Class InteropFormProxyBase
    Implements IInteropForm, IDisposable

#Region "VB6 Constants"

    ' todo: take right from vb6? or make COM visible enum?

    ''' <summary>
    ''' Shows form forced to top of all other forms.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Const VB_SHOW_vbModal As Int32 = 1
    ''' <summary>
    ''' Shows form normally.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Const VB_SHOW_vbModeless As Int32 = 0

    ' BorderStyle constants

    ''' <summary>
    ''' None (no border or border-related elements).
    ''' </summary>
    ''' <remarks></remarks>
    Protected Const VB_BS_vbBSNone As Int32 = 0
    ''' <summary>
    ''' Fixed Single. Can include Control-menu box, title bar, Maximize button, and Minimize button. Resizable only using Maximize and Minimize buttons. 
    ''' </summary>
    ''' <remarks></remarks>
    Protected Const VB_BS_vbFixedSingle As Int32 = 1
    ''' <summary>
    ''' (Default) Sizable. Resizable using any of the optional border elements listed for setting 1. 
    ''' </summary>
    ''' <remarks></remarks>
    Protected Const VB_BS_vbSizable As Int32 = 2
    ''' <summary>
    ''' Fixed Dialog. Can include Control-menu box and title bar; can't include Maximize or Minimize buttons. Not resizable. 
    ''' </summary>
    ''' <remarks></remarks>
    Protected Const VB_BS_vbFixedDouble As Int32 = 3
    ''' <summary>
    ''' Fixed ToolWindow. Displays a non-sizable window with a Close button and title bar text in a reduced font size. The form does not appear in the Windows taskbar. 
    ''' </summary>
    ''' <remarks></remarks>
    Protected Const VB_BS_vbFixedToolWindow As Int32 = 4
    ''' <summary>
    ''' Sizable ToolWindow. Displays a sizable window with a Close button and title bar text in a reduced font size. The form does not appear in the Windows taskbar. 
    ''' </summary>
    ''' <remarks></remarks>
    Protected Const VB_BS_vbSizableToolWindow As Int32 = 5


    ' Startup constants

    ''' <summary>
    ''' No initial setting specified.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Const VB_SP_vbStartUpManual As Int32 = 0
    ''' <summary>
    ''' Center on the item to which the UserForm belongs. 
    ''' </summary>
    ''' <remarks></remarks>
    Protected Const VB_SP_vbStartUpOwner As Int32 = 1
    ''' <summary>
    ''' Center on the whole screen.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Const VB_SP_vbStartUpScreen As Int32 = 2
    ''' <summary>
    ''' Position in upper-left corner of screen. 
    ''' </summary>
    ''' <remarks></remarks>
    Protected Const VB_SP_vbStartUpWindowsDefault As Int32 = 3

    ' WindowState constants

    ''' <summary>
    ''' (Default) Normal. 
    ''' </summary>
    ''' <remarks></remarks>
    Protected Const VB_WS_vbNormal As Int32 = 0
    ''' <summary>
    ''' Minimized (minimized to an icon)
    ''' </summary>
    ''' <remarks></remarks>
    Protected Const VB_WS_vbMinimized As Int32 = 1
    ''' <summary>
    ''' Maximized (enlarged to maximum size)
    ''' </summary>
    ''' <remarks></remarks>
    Protected Const VB_WS_vbMaximized As Int32 = 2

#End Region

#Region "Private Variables"

    ''' <summary>
    ''' An instance of a form.
    ''' </summary>
    ''' <remarks></remarks>
    Private WithEvents _formInstance As Windows.Forms.Form

#End Region

#Region "Friend Methods"

    ''' <summary>
    ''' Closes the form.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Dispose() Implements System.IDisposable.Dispose
        If _formInstance IsNot Nothing Then
            If Not _formInstance.IsDisposed Then
                ' Close form so that any custom close logic fires
                _formInstance.Close()
            Else
                _formInstance = Nothing
            End If
        End If
    End Sub



#End Region

#Region "Protected Properties"

    ''' <summary>
    ''' An instance of a form.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Property FormInstance() As Windows.Forms.Form
        Get
            Return _formInstance
        End Get
        Set(ByVal value As Windows.Forms.Form)
            _formInstance = value
            OnFormInstanceChanged()
        End Set
    End Property

#End Region

#Region "Public Methods"

    ''' <summary>
    ''' Places a specified Form at the front or back of the z-order within its graphical level. Doesn't support named arguments.
    ''' </summary>
    ''' <param name="position"></param>
    ''' <remarks></remarks>
    Public Sub ZOrder(Optional ByVal position As Int32 = 0) Implements IInteropForm.ZOrder
        If position = 0 Then
            _formInstance.BringToFront()
        Else
            _formInstance.SendToBack()
        End If
    End Sub

    ''' <summary>
    ''' Closes a form.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Close() Implements IInteropForm.Close
        _formInstance.Close()
    End Sub

    ''' <summary>
    ''' Hides a form, but does not dispose it.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Hide() Implements IInteropForm.Hide
        _formInstance.Hide()
    End Sub

    ''' <summary>
    ''' Shows form.
    ''' </summary>
    ''' <param name="style"></param>
    ''' <remarks></remarks>
    Public Sub Show(Optional ByVal style As Integer = 0) Implements IInteropForm.Show
        If style = VB_SHOW_vbModal Then
            _formInstance.ShowDialog()
        Else
            _formInstance.Show()
        End If
    End Sub


    Public Sub New()
        Me._formInstance = Nothing
        Me._modelessTabbing = True
    End Sub


    ''' <summary>
    ''' Moves a Form. Doesn't support named arguments.
    ''' </summary>
    ''' <param name="left"></param>
    ''' <param name="top"></param>
    ''' <param name="width"></param>
    ''' <param name="height"></param>
    ''' <remarks></remarks>
    Public Sub Move(ByVal left As Integer, ByVal top As Integer, ByVal width As Integer, ByVal height As Integer) Implements IInteropForm.Move
        _formInstance.Location = New Point(left, top)
        _formInstance.Width = width
        _formInstance.Height = height
    End Sub

    ''' <summary>
    ''' Returns or sets the border style for a form. Read-only at run time.
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public WriteOnly Property BorderStyle() As Integer Implements IInteropForm.BorderStyle
        Set(ByVal value As Integer)
            Select Case value
                Case VB_BS_vbBSNone
                    _formInstance.FormBorderStyle = Windows.Forms.FormBorderStyle.None
                Case VB_BS_vbFixedDouble
                    _formInstance.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedDialog
                Case VB_BS_vbFixedSingle
                    _formInstance.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedSingle
                Case VB_BS_vbFixedToolWindow
                    _formInstance.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedToolWindow
                Case VB_BS_vbSizable
                    _formInstance.FormBorderStyle = Windows.Forms.FormBorderStyle.Sizable
                Case VB_BS_vbSizableToolWindow
                    _formInstance.FormBorderStyle = Windows.Forms.FormBorderStyle.SizableToolWindow
                Case Else
                    _formInstance.FormBorderStyle = Windows.Forms.FormBorderStyle.Sizable
            End Select
        End Set
    End Property

    ''' <summary>
    ''' Returns the name used in code to identify a form. Read-only at run time.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Name() As String Implements IInteropForm.Name
        Get
            Return _formInstance.Name
        End Get
    End Property

    ''' <summary>
    ''' Forces a complete repaint of a form.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Refresh() Implements IInteropForm.Refresh
        _formInstance.Refresh()
    End Sub

    ''' <summary>
    ''' Returns or sets a value that determines whether a Form object appears in the Windows taskbar. Read-only at run time.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ShowInTaskbar() As Boolean Implements IInteropForm.ShowInTaskbar
        Get
            Return _formInstance.ShowInTaskbar
        End Get
        Set(ByVal value As Boolean)
            _formInstance.ShowInTaskbar = value
        End Set
    End Property

    ''' <summary>
    ''' Returns or sets a value indicating whether an object is visible or hidden.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Visible() As Boolean Implements IInteropForm.Visible
        Get
            Return _formInstance.Visible
        End Get
        Set(ByVal value As Boolean)
            _formInstance.Visible = value
        End Set
    End Property

    ''' <summary>
    ''' Returns or sets a value that determines whether a form can respond to user-generated events.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Enabled() As Boolean Implements IInteropForm.Enabled
        Get
            Return _formInstance.Enabled
        End Get
        Set(ByVal value As Boolean)
            _formInstance.Enabled = value
        End Set
    End Property

    ''' <summary>
    ''' Returns or sets a value specifying the position of an object when it first appears. Not available at run time.
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public WriteOnly Property StartUpPosition() As Integer Implements IInteropForm.StartUpPosition
        Set(ByVal value As Integer)
            Select Case value
                Case VB_SP_vbStartUpManual
                    _formInstance.StartPosition = Windows.Forms.FormStartPosition.Manual
                Case VB_SP_vbStartUpOwner
                    _formInstance.StartPosition = Windows.Forms.FormStartPosition.CenterParent
                Case VB_SP_vbStartUpScreen
                    _formInstance.StartPosition = Windows.Forms.FormStartPosition.CenterScreen
                Case VB_SP_vbStartUpWindowsDefault
                    _formInstance.StartPosition = Windows.Forms.FormStartPosition.WindowsDefaultLocation
                Case Else
                    _formInstance.StartPosition = Windows.Forms.FormStartPosition.WindowsDefaultLocation
            End Select

        End Set
    End Property

    ''' <summary>
    ''' Returns or sets a value indicating the visual state of a form window at run time.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property WindowState() As Integer Implements IInteropForm.WindowState
        Get
            Select Case _formInstance.WindowState
                Case Windows.Forms.FormWindowState.Maximized
                    Return VB_WS_vbMaximized
                Case Windows.Forms.FormWindowState.Minimized
                    Return VB_WS_vbMinimized
                Case Windows.Forms.FormWindowState.Normal
                    Return VB_WS_vbNormal
                Case Else
                    Return VB_WS_vbNormal
            End Select
        End Get
        Set(ByVal value As Integer)
            Select Case value
                Case VB_WS_vbMaximized
                    _formInstance.WindowState = Windows.Forms.FormWindowState.Maximized
                Case VB_WS_vbMinimized
                    _formInstance.WindowState = Windows.Forms.FormWindowState.Minimized
                Case VB_WS_vbNormal
                    _formInstance.WindowState = Windows.Forms.FormWindowState.Normal
                Case Else
                    _formInstance.WindowState = Windows.Forms.FormWindowState.Normal
            End Select
        End Set
    End Property

    ''' <summary>
    ''' Return or set the height of a form.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Height() As Integer Implements IInteropForm.Height
        Get
            Return _formInstance.Height
        End Get
        Set(ByVal value As Integer)
            _formInstance.Height = value
        End Set
    End Property

    ''' <summary>
    ''' Gets or Sets the Width of the Form.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Width() As Integer Implements IInteropForm.Width
        Get
            Return _formInstance.Width
        End Get
        Set(ByVal value As Integer)
            _formInstance.Width = value
        End Set
    End Property

    ''' <summary>
    ''' Determines the text displayed in the Form's title bar. When the form is minimized, this text is displayed below the form's icon.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Caption() As String Implements IInteropForm.Caption
        Get
            Return _formInstance.Text
        End Get
        Set(ByVal value As String)
            _formInstance.Text = value
        End Set
    End Property

    ''' <summary>
    ''' Indicates whether the Form has been disposed.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>
    ''' When a .NET Form is closed it is disposed, meaning it
    ''' can no longer be used. This is different from the 
    ''' closing behavior of VB6 Forms which are merely hidden.
    ''' </remarks>
    Public ReadOnly Property IsFormDisposed() As Boolean Implements IInteropForm.IsFormDisposed
        Get
            If _formInstance Is Nothing Then
                Return True
            Else
                Return _formInstance.IsDisposed
            End If
        End Get
    End Property

#End Region

#Region "Protected Methods"

    ''' <summary>
    ''' Adds a form to the InteropFormManager's collection.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Sub RegisterFormInstance()
        InteropToolbox.FormsManagerInstance.RegisterForm(Me)
        AddHandler _formInstance.Disposed, AddressOf formInstance_Disposed
        HookCustomEvents()

    End Sub

    ''' <summary>
    ''' Removes a form to the InteropFormManager's collection.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Sub UnregisterFormInstance()
        InteropToolbox.FormsManagerInstance.UnregisterForm(Me)
        RemoveHandler _formInstance.Disposed, AddressOf formInstance_Disposed
        _formInstance.Dispose()
        _formInstance = Nothing
        UnhookCustomEvents()
    End Sub

    ''' <summary>
    ''' Method to be overridden by derived classes to hook up 
    ''' custom events for reraising.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overridable Sub HookCustomEvents()

    End Sub

    ''' <summary>
    ''' Method to be overridden by derived classes to unhook  
    ''' custom events.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overridable Sub UnhookCustomEvents()

    End Sub

#End Region

#Region "Private Methods"
    Private Sub formInstance_Disposed(ByVal sender As Object, ByVal e As System.EventArgs)
        UnregisterFormInstance()
    End Sub

#End Region



#Region "    Modeless Form Tabbing Support"

    'This code works around a known issue with tabbing in modeless forms.
    'This behavior can be enabled/disabled by setting the EnableModelessTabbing 
    'property;  the default is True.  
    'Note also that Me.KeyPreview must = True.  
    Private _modelessTabbing As Boolean

    ''' <summary>
    ''' Gets or sets value determining if tabbing should be enabled 
    ''' for modeless Interop Form.  
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Calls overridable OnModelessTabbingChange method when value 
    ''' is set.  
    ''' </remarks>
    Public Property ModelessTabbing() As Boolean
        Get
            Return _modelessTabbing
        End Get

        Set(ByVal value As Boolean)
            _modelessTabbing = value
            Me.OnModelessTabbingChange(value)
        End Set
    End Property

    ''' <summary>
    ''' Custom logic that runs when ModeLessTabbing property changes.  
    ''' </summary>
    ''' <param name="newValue">New value set on ModelessTabbing property.</param>
    ''' <remarks>
    ''' Sets KeyPreview to True on form instance when newValue is True.  
    ''' </remarks>
    Protected Overridable Sub OnModelessTabbingChange(ByVal newValue As Boolean)
        If newValue = True AndAlso _formInstance IsNot Nothing Then
            _formInstance.KeyPreview = True
        End If
    End Sub

    ''' <summary>
    ''' Custom logic that runs when _formInstance changes.  
    ''' </summary>
    ''' <remarks>
    ''' Sets KeyPreview to True on form instance when newValue is True.  
    ''' </remarks>
    Protected Overridable Sub OnFormInstanceChanged()
        If Me._formInstance IsNot Nothing AndAlso Me.ModelessTabbing = True Then
            _formInstance.KeyPreview = True
        End If
    End Sub

    Protected Overridable Sub OnFormKeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles _formInstance.KeyDown

        If Me._modelessTabbing = True Then

            If e.KeyCode = Windows.Forms.Keys.Tab Then

                Dim ctl As Windows.Forms.Control = _formInstance.ActiveControl
                Dim firstCtrl As Windows.Forms.Control = _formInstance.ActiveControl

                Do
                    ctl = _formInstance.GetNextControl(ctl, Not e.Shift)
                Loop Until (ctl IsNot Nothing) AndAlso (ctl.CanSelect) OrElse (ctl Is firstCtrl)

                If (ctl IsNot Nothing) AndAlso (ctl.CanSelect) Then
                    ctl.Focus()
                End If
            End If
        End If

    End Sub

#End Region


End Class
