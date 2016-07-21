''' <summary>
''' Defines the event signature exposed by the IInteropEventMessenger 
''' interface so that VB6 can correctly sink the events.
''' </summary>
''' <remarks></remarks>
<System.Runtime.InteropServices.InterfaceTypeAttribute(System.Runtime.InteropServices.ComInterfaceType.InterfaceIsIDispatch), _
 System.Runtime.InteropServices.ComVisible(True)> _
Public Interface IInteropEventMessengerEventSink

#Region "Public Methods"

    ''' <summary>
    ''' Raised when the application starts up (i.e. when VB6 code calls RaiseApplicationStartedupEvent)
    ''' </summary>
    ''' <remarks></remarks>
    <System.Runtime.InteropServices.DispId(1)> Sub ApplicationStartedup()

    ''' <summary>
    ''' Raised when application is about to shut down (i.e. when VB6 code calls OnApplicationShutdown)
    ''' </summary>
    ''' <remarks></remarks>
    <System.Runtime.InteropServices.DispId(2)> Sub ApplicationShutdown()

    ''' <summary>
    ''' Raised when a RaiseApplicationEvent is called. Includes event name and args as parameters.
    ''' </summary>
    ''' <remarks></remarks>
    <System.Runtime.InteropServices.DispId(3)> Sub ApplicationEventRaised(ByVal eventName As String, ByVal eventArgs As Object)

#End Region

End Interface
