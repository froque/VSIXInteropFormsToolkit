Imports System.Runtime.InteropServices
Imports System.Windows.Forms

''' <summary>
''' Class that manages instances of .NET InteropForms.
''' </summary>
''' <remarks></remarks>
<ComVisible(False)> _
Public Class InteropFormsManager

#Region "Private Variables"

    ' List of InteropForm proxies.
    Private _proxies As List(Of InteropFormProxyBase)

#End Region

#Region "Friend Methods"

    ''' <summary>
    ''' Registers an interop form in the collection.
    ''' </summary>
    ''' <param name="proxy"></param>
    ''' <remarks></remarks>
    Friend Sub RegisterForm(ByVal proxy As InteropFormProxyBase)
        _proxies.Add(proxy)
    End Sub

    ''' <summary>
    ''' Removes an interop form wrapper from the collection.
    ''' </summary>
    ''' <param name="proxy"></param>
    ''' <remarks></remarks>
    Friend Sub UnregisterForm(ByVal proxy As InteropFormProxyBase)
        If _proxies.Contains(proxy) Then
            _proxies.Remove(proxy)
        End If
    End Sub


    ''' <summary>
    ''' Friend constructor because this class will only be available
    ''' to outside callers through the InteropToolbox class.
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub New()
        MyBase.New()
        InitializeCollection()
    End Sub

    ''' <summary>
    ''' Tears down all forms.
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub OnApplicationShutdown()
        Try
            CleanupCollection()
        Catch ex As Exception
            Throw New Exception(My.Resources.OnShutDownErrMsg)
        End Try
    End Sub

    Friend Sub OnApplicationStartup()
        ' To be safe, clear out anything.
        CleanupCollection()
        InitializeCollection()
    End Sub

#End Region

#Region "Private Methods"

    ''' <summary>
    ''' Cleans up and recreates the collection.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub InitializeCollection()
        ' Cleanup for good measure.
        CleanupCollection()
        ' Create the collection.
        _proxies = New List(Of InteropFormProxyBase)
    End Sub


    ''' <summary>
    ''' Disposes of any proxies still in the collection.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub CleanupCollection()
        If _proxies IsNot Nothing Then
            While _proxies.Count > 0
                Dim proxy As InteropFormProxyBase = _proxies(0)
                _proxies.RemoveAt(0)
                If proxy IsNot Nothing Then
                    proxy.Dispose()
                End If
            End While

            _proxies.Clear()
            _proxies = Nothing
        End If
    End Sub

#End Region

End Class
