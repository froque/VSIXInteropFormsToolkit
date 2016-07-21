Option Strict On
Imports System.Windows.Forms
Imports System.ComponentModel

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class StatusBar
    Inherits System.Windows.Forms.UserControl

    'StatusBar overrides dispose to clean up the component list.
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
        Me.components = New System.ComponentModel.Container
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.lblUser = New System.Windows.Forms.ToolStripStatusLabel
        Me.StatusLabel = New System.Windows.Forms.ToolStripStatusLabel
        Me.lblCaps = New System.Windows.Forms.ToolStripStatusLabel
        Me.lblNum = New System.Windows.Forms.ToolStripStatusLabel
        Me.lblScroll = New System.Windows.Forms.ToolStripStatusLabel
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.lblUser, Me.ToolStripStatusLabel1, Me.StatusLabel, Me.lblCaps, Me.lblNum, Me.lblScroll})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(489, 22)
        Me.StatusStrip1.TabIndex = 0
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'lblUser
        '
        Me.lblUser.Name = "lblUser"
        Me.lblUser.Size = New System.Drawing.Size(36, 17)
        Me.lblUser.Text = "User: "
        '
        'StatusLabel
        '
        Me.StatusLabel.IsLink = True
        Me.StatusLabel.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline
        Me.StatusLabel.Name = "StatusLabel"
        Me.StatusLabel.Size = New System.Drawing.Size(193, 17)
        Me.StatusLabel.Text = "Click to visit Interop Toolkit Help Forum"
        '
        'lblCaps
        '
        Me.lblCaps.AutoSize = False
        Me.lblCaps.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCaps.ForeColor = System.Drawing.Color.Red
        Me.lblCaps.Name = "lblCaps"
        Me.lblCaps.Size = New System.Drawing.Size(36, 17)
        Me.lblCaps.Text = "CAPS"
        '
        'lblNum
        '
        Me.lblNum.AutoSize = False
        Me.lblNum.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNum.ForeColor = System.Drawing.Color.Red
        Me.lblNum.Name = "lblNum"
        Me.lblNum.Size = New System.Drawing.Size(32, 17)
        Me.lblNum.Text = "NUM"
        '
        'lblScroll
        '
        Me.lblScroll.AutoSize = False
        Me.lblScroll.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblScroll.ForeColor = System.Drawing.Color.Red
        Me.lblScroll.Name = "lblScroll"
        Me.lblScroll.Size = New System.Drawing.Size(49, 17)
        Me.lblScroll.Text = "SCROLL"
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 1500
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(97, 17)
        Me.ToolStripStatusLabel1.Spring = True
        '
        'StatusBar
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.Controls.Add(Me.StatusStrip1)
        Me.Name = "StatusBar"
        Me.Size = New System.Drawing.Size(489, 22)
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusLabel As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents lblNum As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents lblCaps As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents lblUser As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents lblScroll As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel

End Class
