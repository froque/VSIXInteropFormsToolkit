''' <summary>
''' Defines collection of key interop components as Properties
''' </summary>
''' <remarks></remarks>
Public Interface IInteropToolbox

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
    Sub Initialize()

    ''' <summary>
    ''' Returns a class for storing and accessing application state data from VB6 or .NET
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ReadOnly Property Globals() As InteropGlobals

    ''' <summary>
    ''' Controls event-based communication between VB6 and .NET
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ReadOnly Property EventMessenger() As InteropEventMessenger

End Interface
