<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class HelloWorldForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(HelloWorldForm))
        Me.lblHelloText = New System.Windows.Forms.Label
        Me.lnkRaiseEvent = New System.Windows.Forms.LinkLabel
        Me.SuspendLayout()
        '
        'lblHelloText
        '
        resources.ApplyResources(Me.lblHelloText, "lblHelloText")
        Me.lblHelloText.BackColor = System.Drawing.Color.CornflowerBlue
        Me.lblHelloText.ForeColor = System.Drawing.Color.DarkBlue
        Me.lblHelloText.Name = "lblHelloText"
        '
        'lnkRaiseEvent
        '
        Me.lnkRaiseEvent.ActiveLinkColor = System.Drawing.Color.DarkBlue
        resources.ApplyResources(Me.lnkRaiseEvent, "lnkRaiseEvent")
        Me.lnkRaiseEvent.BackColor = System.Drawing.Color.Transparent
        Me.lnkRaiseEvent.DisabledLinkColor = System.Drawing.Color.DarkBlue
        Me.lnkRaiseEvent.ForeColor = System.Drawing.Color.DarkBlue
        Me.lnkRaiseEvent.Name = "lnkRaiseEvent"
        Me.lnkRaiseEvent.TabStop = True
        '
        'HelloWorldForm
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.DarkOrange
        Me.Controls.Add(Me.lnkRaiseEvent)
        Me.Controls.Add(Me.lblHelloText)
        Me.Name = "HelloWorldForm"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents lblHelloText As System.Windows.Forms.Label
    Friend WithEvents lnkRaiseEvent As System.Windows.Forms.LinkLabel

End Class
