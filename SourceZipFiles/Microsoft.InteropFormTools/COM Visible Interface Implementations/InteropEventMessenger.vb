Imports System.Runtime.InteropServices

''' <summary>
''' Class controlling event-based communication between VB6 and .NET
''' </summary>
''' <remarks></remarks>
<ClassInterface(ClassInterfaceType.None), _
System.Runtime.InteropServices.ComSourceInterfaces(GetType(IInteropEventMessengerEventSink))> _
Public Class InteropEventMessenger
    Implements IInteropEventMessenger

    ''' <summary>
    ''' Indicatest application has been shut down.
    ''' </summary>
    ''' <remarks></remarks>
    Public Event ApplicationShutdown() Implements IInteropEventMessenger.ApplicationShutdown

    ''' <summary>
    ''' Indicates application has started up.
    ''' </summary>
    ''' <remarks></remarks>
    Public Event ApplicationStartedup() Implements IInteropEventMessenger.ApplicationStartedup

    ''' <summary>
    ''' Indicates an application event has occurred.
    ''' </summary>
    ''' <param name="eventName">The name of the event.</param>
    ''' <param name="eventArgs">Additional event information</param>
    ''' <remarks></remarks>
    Public Event ApplicationEventRaised(ByVal eventName As String, ByVal eventArgs As Object) Implements IInteropEventMessenger.ApplicationEventRaised

    ''' <summary>
    ''' Raises event which indicates the application has shuts down.
    ''' </summary>
    ''' <remarks>
    ''' This is intended to be called by the VB6 shutdown code.
    ''' </remarks>
    Public Sub RaiseApplicationShutdownEvent() Implements IInteropEventMessenger.RaiseApplicationShutdownEvent
        RaiseEvent ApplicationShutdown()
    End Sub

    ''' <summary>
    ''' Raises event which indicates that the applicaiton has started up.
    ''' </summary>
    ''' <remarks>
    ''' This is intended to be called by the VB6 startup code.
    ''' </remarks>
    Public Sub RaiseApplicationStartedupEvent() Implements IInteropEventMessenger.RaiseApplicationStartedupEvent
        RaiseEvent ApplicationStartedup()
    End Sub

    ''' <summary>
    ''' Raises an indication that a custom application event has occurred.
    ''' </summary>
    ''' <param name="eventName"></param>
    ''' <remarks></remarks>
    Public Sub RaiseApplicationEvent(ByVal eventName As String, ByVal eventArgs As Object) Implements IInteropEventMessenger.RaiseApplicationEvent
        RaiseEvent ApplicationEventRaised(eventName, eventArgs)
    End Sub

    ''' <summary>
    ''' Default constructor.
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub New()

    End Sub
End Class
