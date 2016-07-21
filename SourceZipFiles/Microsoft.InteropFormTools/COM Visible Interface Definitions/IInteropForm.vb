''' <summary>
''' Defines the methods and properties that will be 
''' exposed by any InteropForm wrapper class exposed to VB6
''' </summary>
''' <remarks></remarks>
Public Interface IInteropForm

#Region "Methods"

    ''' <summary>
    ''' Closes an object and terminates the connection to the application that provided the object.
    ''' </summary>
    ''' <remarks></remarks>
    Sub Close()
    ''' <summary>
    ''' Hides an MDIForm or Form object but doesn't unload it.
    ''' </summary>
    ''' <remarks></remarks>
    Sub Hide()
    ''' <summary>
    ''' Moves a Form. Doesn't support named arguments.
    ''' </summary>
    ''' <param name="left"></param>
    ''' <param name="top"></param>
    ''' <param name="width"></param>
    ''' <param name="height"></param>
    ''' <remarks></remarks>
    Sub Move(ByVal left As Int32, ByVal top As Int32, ByVal width As Int32, ByVal height As Int32)
    ''' <summary>
    ''' Forces a complete repaint of a form or control.
    ''' </summary>
    ''' <remarks></remarks>
    Sub Refresh()
    ''' <summary>
    ''' Displays an MDIForm or Form object. Doesn't support named arguments.
    ''' </summary>
    ''' <param name="style"></param>
    ''' <remarks></remarks>
    Sub Show(Optional ByVal style As Int32 = 0)
    ''' <summary>
    ''' Places a specified Form at the front or back of the z-order within its graphical level. Doesn't support named arguments.
    ''' </summary>
    ''' <param name="position"></param>
    ''' <remarks></remarks>
    Sub ZOrder(Optional ByVal position As Int32 = 0)

#End Region

#Region "Properties"

    ''' <summary>
    ''' Returns or sets the border style for a form. Read-only at run time.
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    WriteOnly Property BorderStyle() As Int32
    ''' <summary>
    ''' Determines the text displayed in the Form's title bar. When the form is minimized, this text is displayed below the form's icon.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property Caption() As String
    ''' <summary>
    ''' Returns or sets a value that determines whether a form can respond to user-generated events.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property Enabled() As Boolean
    ''' <summary>
    ''' Return or set the height of a form.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property Height() As Int32
    ''' <summary>
    ''' Indicates whether the underlying .NET Form has been disposed.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>
    ''' When a .NET Form is closed it is disposed, meaning it
    ''' can no longer be used. This is different from the 
    ''' closing behavior of VB6 Forms which are merely hidden.
    ''' </remarks>
    ReadOnly Property IsFormDisposed() As Boolean
    ''' <summary>
    ''' Returns the name used in code to identify a form. Read-only at run time.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ReadOnly Property Name() As String
    ''' <summary>
    ''' Returns or sets a value that determines whether a Form object appears in the Windows taskbar. Read-only at run time.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property ShowInTaskbar() As Boolean
    ''' <summary>
    ''' Returns or sets a value specifying the position of an object when it first appears. Not available at run time.
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    WriteOnly Property StartUpPosition() As Int32
    ''' <summary>
    ''' Returns or sets a value indicating whether an object is visible or hidden.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property Visible() As Boolean
    ''' <summary>
    ''' Returns or sets a Single containing the width of the window in twips. Read/write.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property Width() As Int32
    ''' <summary>
    ''' Returns or sets a value indicating the visual state of a form window at run time.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property WindowState() As Int32

#End Region

End Interface
