Imports System.Runtime.InteropServices

''' <summary>
''' Class that serves as centralized collection of all interop communication pieces
''' </summary>
''' <remarks>
''' All functionality exposed to COM via explicit interface - No class interface.
''' Class is only created within this assembly and is only created by the 
''' InteropToolbox which then acts as a facade encapsulating all Interop
''' functionality in one place.
''' </remarks>
<ClassInterface(ClassInterfaceType.None)> _
Public Class InteropToolbox
    Implements IInteropToolbox

#Region "Private Variables"

    Private Shared _formsManager As InteropFormsManager
    Private Shared _globals As InteropGlobals
    Private Shared _eventMessenger As InteropEventMessenger

    Private Shared _isInitialized As Boolean = False

#End Region

#Region "Constructors"


    ''' <summary>
    ''' Default Constructor
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()

        MyBase.New()
        InitializeToolboxComponents()

    End Sub

#End Region

#Region "Private Methods"

    ''' <summary>
    ''' Creates the interop helper objects behind the 
    ''' properties exposed and wires up events
    ''' </summary>
    ''' <remarks></remarks>
    Private Shared Sub InitializeToolboxComponents()
        ' check outside of lock for speed
        If Not _isInitialized Then
            SyncLock GetType(InteropToolbox)
                Try
                    ' double check in lock
                    If Not _isInitialized Then

                        _formsManager = New InteropFormsManager()
                        _globals = New InteropGlobals()
                        _eventMessenger = New InteropEventMessenger()

                        AddHandler _eventMessenger.ApplicationStartedup, AddressOf _eventMessenger_ApplicationStartedup
                        AddHandler _eventMessenger.ApplicationShutdown, AddressOf _eventMessenger_OnApplicationShutdown

                        _isInitialized = True
                    End If
                Catch ex As Exception
                    _isInitialized = False
                    ' clear out shared items since null check is used for 
                    _eventMessenger = Nothing
                    _formsManager = Nothing
                    _globals = Nothing

                    Throw ex

                End Try
            End SyncLock
        End If
    End Sub


    ''' <summary>
    ''' Handles application start up
    ''' </summary>
    ''' <remarks></remarks>
    Private Shared Sub _eventMessenger_ApplicationStartedup()

        ' pass on event to forms manager
        If _formsManager IsNot Nothing Then
            _formsManager.OnApplicationStartup()
        End If
        ' pass on event to globals
        If _globals IsNot Nothing Then
            _globals.OnApplicationStartup()
        End If

    End Sub

    ''' <summary>
    ''' Handles application shutdown up
    ''' </summary>
    ''' <remarks></remarks>
    Private Shared Sub _eventMessenger_OnApplicationShutdown()

        ' pass on event to forms manager
        If _formsManager IsNot Nothing Then
            _formsManager.OnApplicationShutdown()
        End If
        ' pass on event to globals
        If _globals IsNot Nothing Then
            _globals.OnApplicationShutdown()
        End If

        If _eventMessenger IsNot Nothing Then
            RemoveHandler _eventMessenger.ApplicationStartedup, AddressOf _eventMessenger_ApplicationStartedup
            RemoveHandler _eventMessenger.ApplicationShutdown, AddressOf _eventMessenger_OnApplicationShutdown
        End If

    End Sub


#End Region

#Region "Friend Properties"

    ''' <summary>
    ''' Provides the FormsManager for the InteropFormW
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend ReadOnly Property FormsManager() As InteropFormsManager
        Get
            Return _formsManager
        End Get
    End Property

    Friend Shared ReadOnly Property FormsManagerInstance() As InteropFormsManager
        Get
            ' make sure it's initialized
            InitializeToolboxComponents()
            Return _formsManager
        End Get
    End Property


#End Region

#Region "Public Properties"


    ''' <summary>
    ''' Returns a class for storing and accessing application state data from VB6 or .NET
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Globals() As InteropGlobals Implements IInteropToolbox.Globals
        Get
            Return _globals
        End Get
    End Property


    ''' <summary>
    ''' Controls event-based communication between VB6 and .NET
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property EventMessenger() As InteropEventMessenger Implements IInteropToolbox.EventMessenger
        Get
            Return _eventMessenger
        End Get
    End Property

#End Region

#Region "Public Methods"

    ''' <summary>
    ''' This method should be called ONLY when the InteropToolbox
    ''' is created for the first time during the startup of 
    ''' the VB6 Application.
    ''' </summary>
    ''' <remarks>
    ''' This method aids in the Debugging experience in VB6
    ''' by clearing out some elements in Shared (static)
    ''' memory that may not have been cleared by simply
    ''' stopping (clicking the Stop button) in the VB6 IDE.
    ''' </remarks>
    Public Sub Initialize() Implements IInteropToolbox.Initialize

        ' Call shutdown logic
        _eventMessenger_OnApplicationShutdown()

        ' Reinitialize
        _isInitialized = False
        InitializeToolboxComponents()

    End Sub

#End Region

End Class


