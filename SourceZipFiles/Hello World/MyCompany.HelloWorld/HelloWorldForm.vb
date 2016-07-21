Imports Microsoft.InteropFormTools

<InteropForm()> _
Public Class HelloWorldForm

#Region "Public Constructors"
    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    <InteropFormInitializer()> _
    Public Sub New(ByVal helloText As String)
        Me.New()

        lblHelloText.Text = helloText

    End Sub

    <InteropFormInitializer()> _
    Public Sub New(ByVal helloText As String, ByVal repeatTimes As Int32)
        Me.New()

        lblHelloText.Text = String.Empty

        For i As Int32 = 0 To repeatTimes - 1
            lblHelloText.Text += helloText + vbNewLine
        Next
    End Sub

#End Region

#Region "Public Methods"

    ''' <summary>
    ''' Swaps the background colors of the Label and Form
    ''' </summary>
    ''' <remarks>This is a trivial method showing the use of the InteropFormMethod attribute</remarks>
    <InteropFormMethod()> _
    Public Sub ReverseBackgroundColors()

        Dim lblOldBackColor As Color = Me.lblHelloText.BackColor
        lblHelloText.BackColor = Me.BackColor
        Me.BackColor = lblOldBackColor

    End Sub

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' Gets/Sets the hello text in the form
    ''' </summary>
    ''' <remarks>This is a trivial method showing the use of the InteropFormProperty attribute</remarks>
    <InteropFormProperty()> _
    Public Property HelloText() As String
        Get
            Return lblHelloText.Text
        End Get
        Set(ByVal value As String)
            lblHelloText.Text = value
        End Set
    End Property

#End Region

#Region "Public Events"

    <InteropFormEvent()> _
    Public Event SampleEvent As SampleEventHandler

    Public Delegate Sub SampleEventHandler(ByVal sampleEventText As String)

#End Region

#Region "Private Methods"

    ''' <summary>
    ''' Event handler for link label click event
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub lnkRaiseEvent_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkRaiseEvent.LinkClicked
        'displays "Current Form Text is """ + lblHelloText.Text & """"
        RaiseEvent SampleEvent(String.Format(My.Resources.SampleEventText, lblHelloText.Text))
    End Sub

#End Region

End Class
