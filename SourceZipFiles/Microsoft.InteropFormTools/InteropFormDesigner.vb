Imports System
Imports System.Collections
Imports System.ComponentModel
Imports System.ComponentModel.Design
Imports System.Drawing
Imports System.Windows.Forms

'Imports EnvDTE
'Imports EnvDTE80
'Imports VSLangProj
'Imports VSLangProj2
'Imports VSLangProj80

Public Class InteropFormDesigner
    Implements System.ComponentModel.Design.IDesigner




    ' Local reference to the designer's component.
    Private component_ As IComponent

    Public ReadOnly Property Component() As System.ComponentModel.IComponent Implements System.ComponentModel.Design.IDesigner.Component
        Get
            Return component_
        End Get
    End Property

    Public Sub DoDefaultAction() Implements System.ComponentModel.Design.IDesigner.DoDefaultAction


        'Dim proj As EnvDTE.Project
        'Dim vsproject As VSLangProj80.VSProject2
        'Dim d As New DTE()
        'proj = d.Solution.Projects.Item(1)
        'vsproject = CType(proj.Object, VSLangProj80.VSProject2)

        'MessageBox.Show(vsproject.Project.FullName)
        ' Shows a message box indicating that the default action for the designer was invoked.
        MessageBox.Show("The DoDefaultAction method of an IDesigner implementation was invoked.", "Information")


    End Sub

    Public Sub Initialize(ByVal component As System.ComponentModel.IComponent) Implements System.ComponentModel.Design.IDesigner.Initialize
        ' This method is called after a designer for a component is created,
        ' and stores a reference to the designer's component.

        Me.component_ = component
        'MessageBox.Show(GetType(Component).ToString())

    End Sub

    Public ReadOnly Property Verbs() As System.ComponentModel.Design.DesignerVerbCollection Implements System.ComponentModel.Design.IDesigner.Verbs
        Get
            Dim verbs_ As New DesignerVerbCollection()
            Dim dv1 As New DesignerVerb("Display Component Name", New EventHandler(AddressOf Me.ShowComponentName))
            verbs_.Add(dv1)
            Return verbs_

        End Get
    End Property

    ' Event handler for displaying a message box showing the designer's component's name.
    Private Sub ShowComponentName(ByVal sender As Object, ByVal e As EventArgs)
        If Not (Me.Component Is Nothing) Then
            MessageBox.Show(Me.Component.Site.Name, "Designer Component's Name")
        End If
    End Sub


    Private disposedValue As Boolean = False        ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: free unmanaged resources when explicitly called
            End If

            ' TODO: free shared unmanaged resources
        End If
        Me.disposedValue = True
    End Sub

#Region " IDisposable Support "
    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
