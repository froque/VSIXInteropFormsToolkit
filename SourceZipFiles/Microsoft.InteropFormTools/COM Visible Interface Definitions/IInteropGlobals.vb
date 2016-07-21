''' <summary>
''' Defines methods for storing and accessing application state 
''' data for access from VB6 or .NET
''' </summary>
''' <remarks></remarks>
Public Interface IInteropGlobals

    ''' <summary>
    ''' Indexer that returns object for the given key
    ''' </summary>
    ''' <param name="key"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Default Property Item(ByVal key As Object) As Object

    ''' <summary>
    ''' Stores an Object value and associates it with the given key 
    ''' </summary>
    ''' <param name="key"></param>
    ''' <param name="value"></param>
    ''' <remarks></remarks>
    Sub Add(ByVal key As Object, ByVal value As Object)

    ''' <summary>
    ''' Removes an Object associated with the given key
    ''' </summary>
    ''' <param name="key"></param>
    ''' <remarks></remarks>
    Sub Remove(ByVal key As Object)

    ''' <summary>
    ''' Returns value indicating whether object exists with the given key
    ''' </summary>
    ''' <param name="key"></param>
    ''' <remarks></remarks>
    Function ContainsKey(ByVal key As Object) As Boolean

End Interface
