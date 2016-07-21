''' <summary>
''' Interface defining event-based communication between VB6 and .NET.
''' </summary>
''' <remarks></remarks>
Public Interface IInteropEventMessenger

#Region "Public Methods"

    ''' <summary>
    ''' Called by VB6 when application starts up.  Raises ApplicationStartedUp event.
    ''' </summary>
    ''' <remarks></remarks>
    Sub RaiseApplicationStartedupEvent()

    ''' <summary>
    ''' Called by VB6 when application has shut down. Raises ApplicationShutdown event.
    ''' </summary>
    ''' <remarks></remarks>
    Sub RaiseApplicationShutdownEvent()

    ''' <summary>
    ''' Called to signal an application event.  Takes event name as and event args as paramters. 
    ''' </summary>
    ''' <param name="eventName">Name of event</param>
    ''' <remarks></remarks>
    Sub RaiseApplicationEvent(ByVal eventName As String, ByVal eventArgs As Object)

#End Region

#Region "Public Events"

    ''' <summary>
    ''' Raised when the application starts up (i.e. when VB6 code calls RaiseApplicationStartedupEvent).
    ''' </summary>
    ''' <remarks></remarks>
    Event ApplicationStartedup()

    ''' <summary>
    ''' Raised when application is about to shut down (i.e. when VB6 code calls OnApplicationShutdown).
    ''' </summary>
    ''' <remarks></remarks>
    Event ApplicationShutdown()

    ''' <summary>
    ''' Raised when a RaiseApplicationEvent is called. Includes event name and args as parameters.
    ''' </summary>
    ''' <remarks></remarks>
    Event ApplicationEventRaised(ByVal eventName As String, ByVal eventArgs As Object)

#End Region

End Interface
