<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmOptions
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.cboEventSeqType = New System.Windows.Forms.ComboBox()
        Me.cboWait = New System.Windows.Forms.ComboBox()
        Me.cboStopbits = New System.Windows.Forms.ComboBox()
        Me.cboDatabits = New System.Windows.Forms.ComboBox()
        Me.cboParity = New System.Windows.Forms.ComboBox()
        Me.cboSpeed = New System.Windows.Forms.ComboBox()
        Me.lblEventSeqType = New System.Windows.Forms.Label()
        Me.lblWait = New System.Windows.Forms.Label()
        Me.lblStopbits = New System.Windows.Forms.Label()
        Me.lblDatabits = New System.Windows.Forms.Label()
        Me.lblParity = New System.Windows.Forms.Label()
        Me.lblSpeed = New System.Windows.Forms.Label()
        Me.lblPort = New System.Windows.Forms.Label()
        Me.cboPort = New System.Windows.Forms.ComboBox()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.cboEventSeqType)
        Me.GroupBox1.Controls.Add(Me.cboWait)
        Me.GroupBox1.Controls.Add(Me.cboStopbits)
        Me.GroupBox1.Controls.Add(Me.cboDatabits)
        Me.GroupBox1.Controls.Add(Me.cboParity)
        Me.GroupBox1.Controls.Add(Me.cboSpeed)
        Me.GroupBox1.Controls.Add(Me.lblEventSeqType)
        Me.GroupBox1.Controls.Add(Me.lblWait)
        Me.GroupBox1.Controls.Add(Me.lblStopbits)
        Me.GroupBox1.Controls.Add(Me.lblDatabits)
        Me.GroupBox1.Controls.Add(Me.lblParity)
        Me.GroupBox1.Controls.Add(Me.lblSpeed)
        Me.GroupBox1.Controls.Add(Me.lblPort)
        Me.GroupBox1.Controls.Add(Me.cboPort)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(298, 224)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Communications"
        '
        'cboEventSeqType
        '
        Me.cboEventSeqType.FormattingEnabled = True
        Me.cboEventSeqType.Items.AddRange(New Object() {"1", "2", "3", "4", "5", "6", "7", "8"})
        Me.cboEventSeqType.Location = New System.Drawing.Point(151, 181)
        Me.cboEventSeqType.Name = "cboEventSeqType"
        Me.cboEventSeqType.Size = New System.Drawing.Size(121, 21)
        Me.cboEventSeqType.TabIndex = 13
        Me.cboEventSeqType.Text = "8"
        '
        'cboWait
        '
        Me.cboWait.FormattingEnabled = True
        Me.cboWait.Items.AddRange(New Object() {"200", "300", "400", "500", "600", "700", "800"})
        Me.cboWait.Location = New System.Drawing.Point(151, 154)
        Me.cboWait.Name = "cboWait"
        Me.cboWait.Size = New System.Drawing.Size(121, 21)
        Me.cboWait.TabIndex = 12
        Me.cboWait.Text = "300"
        '
        'cboStopbits
        '
        Me.cboStopbits.FormattingEnabled = True
        Me.cboStopbits.Items.AddRange(New Object() {"1", "1.5", "2"})
        Me.cboStopbits.Location = New System.Drawing.Point(151, 127)
        Me.cboStopbits.Name = "cboStopbits"
        Me.cboStopbits.Size = New System.Drawing.Size(121, 21)
        Me.cboStopbits.TabIndex = 11
        Me.cboStopbits.Text = "1"
        '
        'cboDatabits
        '
        Me.cboDatabits.FormattingEnabled = True
        Me.cboDatabits.Items.AddRange(New Object() {"4", "5", "6", "7", "8"})
        Me.cboDatabits.Location = New System.Drawing.Point(151, 100)
        Me.cboDatabits.Name = "cboDatabits"
        Me.cboDatabits.Size = New System.Drawing.Size(121, 21)
        Me.cboDatabits.TabIndex = 10
        Me.cboDatabits.Text = "8"
        '
        'cboParity
        '
        Me.cboParity.FormattingEnabled = True
        Me.cboParity.Items.AddRange(New Object() {"Even", "Mark", "None", "Odd", "Space"})
        Me.cboParity.Location = New System.Drawing.Point(151, 73)
        Me.cboParity.Name = "cboParity"
        Me.cboParity.Size = New System.Drawing.Size(121, 21)
        Me.cboParity.TabIndex = 9
        Me.cboParity.Text = "Odd"
        '
        'cboSpeed
        '
        Me.cboSpeed.FormattingEnabled = True
        Me.cboSpeed.Items.AddRange(New Object() {"2400", "9600"})
        Me.cboSpeed.Location = New System.Drawing.Point(151, 46)
        Me.cboSpeed.Name = "cboSpeed"
        Me.cboSpeed.Size = New System.Drawing.Size(121, 21)
        Me.cboSpeed.TabIndex = 8
        Me.cboSpeed.Text = "9600"
        '
        'lblEventSeqType
        '
        Me.lblEventSeqType.AutoSize = True
        Me.lblEventSeqType.Location = New System.Drawing.Point(28, 184)
        Me.lblEventSeqType.Name = "lblEventSeqType"
        Me.lblEventSeqType.Size = New System.Drawing.Size(117, 13)
        Me.lblEventSeqType.TabIndex = 7
        Me.lblEventSeqType.Text = "Event Sequence Type:"
        '
        'lblWait
        '
        Me.lblWait.AutoSize = True
        Me.lblWait.Location = New System.Drawing.Point(28, 157)
        Me.lblWait.Name = "lblWait"
        Me.lblWait.Size = New System.Drawing.Size(66, 13)
        Me.lblWait.TabIndex = 6
        Me.lblWait.Text = "Wait (msec):"
        '
        'lblStopbits
        '
        Me.lblStopbits.AutoSize = True
        Me.lblStopbits.Location = New System.Drawing.Point(28, 130)
        Me.lblStopbits.Name = "lblStopbits"
        Me.lblStopbits.Size = New System.Drawing.Size(48, 13)
        Me.lblStopbits.TabIndex = 5
        Me.lblStopbits.Text = "Stopbits:"
        '
        'lblDatabits
        '
        Me.lblDatabits.AutoSize = True
        Me.lblDatabits.Location = New System.Drawing.Point(28, 103)
        Me.lblDatabits.Name = "lblDatabits"
        Me.lblDatabits.Size = New System.Drawing.Size(49, 13)
        Me.lblDatabits.TabIndex = 4
        Me.lblDatabits.Text = "Databits:"
        '
        'lblParity
        '
        Me.lblParity.AutoSize = True
        Me.lblParity.Location = New System.Drawing.Point(28, 76)
        Me.lblParity.Name = "lblParity"
        Me.lblParity.Size = New System.Drawing.Size(36, 13)
        Me.lblParity.TabIndex = 3
        Me.lblParity.Text = "Parity:"
        '
        'lblSpeed
        '
        Me.lblSpeed.AutoSize = True
        Me.lblSpeed.Location = New System.Drawing.Point(28, 49)
        Me.lblSpeed.Name = "lblSpeed"
        Me.lblSpeed.Size = New System.Drawing.Size(41, 13)
        Me.lblSpeed.TabIndex = 2
        Me.lblSpeed.Text = "Speed:"
        '
        'lblPort
        '
        Me.lblPort.AutoSize = True
        Me.lblPort.Location = New System.Drawing.Point(28, 22)
        Me.lblPort.Name = "lblPort"
        Me.lblPort.Size = New System.Drawing.Size(29, 13)
        Me.lblPort.TabIndex = 1
        Me.lblPort.Text = "Port:"
        '
        'cboPort
        '
        Me.cboPort.FormattingEnabled = True
        Me.cboPort.Items.AddRange(New Object() {"1", "2", "3", "4", "5", "6", "7", "8"})
        Me.cboPort.Location = New System.Drawing.Point(151, 19)
        Me.cboPort.Name = "cboPort"
        Me.cboPort.Size = New System.Drawing.Size(121, 21)
        Me.cboPort.TabIndex = 0
        Me.cboPort.Text = "1"
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(233, 242)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 1
        Me.btnClose.Text = "&Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'frmOptions
        '
        Me.AcceptButton = Me.btnClose
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(320, 277)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmOptions"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Options"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents lblPort As System.Windows.Forms.Label
    Friend WithEvents cboPort As System.Windows.Forms.ComboBox
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents cboEventSeqType As System.Windows.Forms.ComboBox
    Friend WithEvents cboWait As System.Windows.Forms.ComboBox
    Friend WithEvents cboStopbits As System.Windows.Forms.ComboBox
    Friend WithEvents cboDatabits As System.Windows.Forms.ComboBox
    Friend WithEvents cboParity As System.Windows.Forms.ComboBox
    Friend WithEvents cboSpeed As System.Windows.Forms.ComboBox
    Friend WithEvents lblEventSeqType As System.Windows.Forms.Label
    Friend WithEvents lblWait As System.Windows.Forms.Label
    Friend WithEvents lblStopbits As System.Windows.Forms.Label
    Friend WithEvents lblDatabits As System.Windows.Forms.Label
    Friend WithEvents lblParity As System.Windows.Forms.Label
    Friend WithEvents lblSpeed As System.Windows.Forms.Label
End Class
