Imports System.Runtime.InteropServices

''' <summary>
''' Class for storing and accessing application state data from VB6 or .NET
''' </summary>
''' <remarks>
''' All functionality exposed to COM via explicit interface - No class interface.
''' Class is only created within this assembly and is only created by the 
''' InteropToolbox which then acts as a facade encapsulating all Interop
''' functionality in one place.
''' </remarks>
<ClassInterface(ClassInterfaceType.None)> _
Public Class InteropGlobals
    Implements IInteropGlobals

#Region "Private Variables"

    Private _globalItems As System.Collections.Hashtable

#End Region

#Region "Constructors"

    ''' <summary>
    ''' Friend constructor because this class will only be available
    ''' to outside callers through the InteropToolbox class
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub New()
        MyBase.New()

        ' instantiate the internal collection
        InitializeCollection()

    End Sub

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' Gets or sets value stored for the given key.  If key exists previous value
    ''' is overwritten.  If key does not exist key is added (i.e. Add is called).
    ''' </summary>
    ''' <param name="key"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Default Public Property Item(ByVal key As Object) As Object Implements IInteropGlobals.Item
        Get
            key = key.ToString.ToLower()
            'Check to see if the requested item exists in the cache 
            If _globalItems.ContainsKey(key) Then
                Return _globalItems(key)
            Else
                Return Nothing
            End If
        End Get
        Set(ByVal value As Object)
            key = key.ToString.ToLower()
            If _globalItems.ContainsKey(key) Then
                _globalItems(key) = value
            Else
                Add(key, value)
            End If
        End Set
    End Property

#End Region

#Region "Public Methods"

    ''' <summary>
    ''' Stores an Object value and associates it with the given key 
    ''' </summary>
    ''' <param name="key"></param>
    ''' <param name="value"></param>
    ''' <remarks></remarks>
    Public Sub Add(ByVal key As Object, ByVal value As Object) Implements IInteropGlobals.Add
        key = key.ToString.ToLower()
        _globalItems.Add(key, value)
    End Sub

    ''' <summary>
    ''' Removes an Object associated with the given key
    ''' </summary>
    ''' <param name="key"></param>
    ''' <remarks></remarks>
    Public Sub Remove(ByVal key As Object) Implements IInteropGlobals.Remove
        key = key.ToString.ToLower()
        If _globalItems.ContainsKey(key) Then
            _globalItems.Remove(key)
        End If
    End Sub

    ''' <summary>
    ''' Returns value indicating whether object exists with the given key
    ''' </summary>
    ''' <param name="key"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ContainsKey(ByVal key As Object) As Boolean Implements IInteropGlobals.ContainsKey
        key = key.ToString.ToLower()
        Return _globalItems.ContainsKey(key)
    End Function
#End Region

#Region "Private Methods"

    ''' <summary>
    ''' Clears out any old items and recreates the internal collection
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub InitializeCollection()
        CleanupCollection()
        _globalItems = New Hashtable()
    End Sub

    ''' <summary>
    ''' Clears out any old items and recreates the internal collection
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub CleanupCollection()
        If _globalItems IsNot Nothing Then
            Try
                For Each globalItem As Object In _globalItems
                    If globalItem Is GetType(IDisposable) Then
                        CType(globalItem, IDisposable).Dispose()
                    End If
                    globalItem = Nothing
                Next
            Catch
                ' eat any exceptions
            Finally
                _globalItems.Clear()
                _globalItems = Nothing
            End Try
        End If
    End Sub


#End Region

#Region "Friend Methods"

    ''' <summary>
    ''' Handles application start up to reinitialize the collection.
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub OnApplicationStartup()
        InitializeCollection()
    End Sub

    ''' <summary>
    ''' Handles application shutdown up to clean up the collection.
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub OnApplicationShutdown()
        CleanupCollection()
    End Sub

#End Region

End Class


