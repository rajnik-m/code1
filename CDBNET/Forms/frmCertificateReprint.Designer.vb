<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCertificateReprint
  Inherits ThemedForm

  'Form overrides dispose to clean up the component list.
  <System.Diagnostics.DebuggerNonUserCode()> _
  Protected Overrides Sub Dispose(ByVal disposing As Boolean)
    Try
      If disposing AndAlso components IsNot Nothing Then
        components.Dispose()
      End If
    Finally
      MyBase.Dispose(disposing)
    End Try
  End Sub

  'Required by the Windows Form Designer
  Private components As System.ComponentModel.IContainer

  'NOTE: The following procedure is required by the Windows Form Designer
  'It can be modified using the Windows Form Designer.  
  'Do not modify it using the code editor.
  <System.Diagnostics.DebuggerStepThrough()> _
  Private Sub InitializeComponent()
    Me.epl = New CDBNETCL.EditPanel()
    Me.cmdOk = New System.Windows.Forms.Button()
    Me.cmdCancel = New System.Windows.Forms.Button()
    Me.bpl = New CDBNETCL.ButtonPanel()
    Me.bpl.SuspendLayout()
    Me.SuspendLayout()
    '
    'epl
    '
    Me.epl.AddressChanged = False
    Me.epl.BackColor = System.Drawing.Color.Transparent
    Me.epl.DataChanged = False
    Me.epl.Dock = System.Windows.Forms.DockStyle.Top
    Me.epl.Location = New System.Drawing.Point(0, 0)
    Me.epl.Name = "epl"
    Me.epl.Recipients = Nothing
    Me.epl.Size = New System.Drawing.Size(378, 125)
    Me.epl.SuppressDrawing = False
    Me.epl.TabIndex = 7
    Me.epl.TabSelectedIndex = 0
    '
    'cmdOk
    '
    Me.cmdOk.BackColor = System.Drawing.Color.Transparent
    Me.cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK
    Me.cmdOk.Enabled = False
    Me.cmdOk.Location = New System.Drawing.Point(196, 6)
    Me.cmdOk.Name = "cmdOk"
    Me.cmdOk.Size = New System.Drawing.Size(96, 27)
    Me.cmdOk.TabIndex = 8
    Me.cmdOk.Text = "OK"
    Me.cmdOk.UseVisualStyleBackColor = False
    '
    'cmdCancel
    '
    Me.cmdCancel.BackColor = System.Drawing.Color.Transparent
    Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdCancel.Location = New System.Drawing.Point(85, 6)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(96, 27)
    Me.cmdCancel.TabIndex = 9
    Me.cmdCancel.Text = "Cancel"
    Me.cmdCancel.UseVisualStyleBackColor = False
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdOk)
    Me.bpl.Controls.Add(Me.cmdCancel)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bpl.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bpl.Location = New System.Drawing.Point(0, 125)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(378, 39)
    Me.bpl.TabIndex = 10
    '
    'frmCertificateReprint
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(378, 164)
    Me.Controls.Add(Me.bpl)
    Me.Controls.Add(Me.epl)
    Me.MaximizeBox = False
    Me.MinimizeBox = False
    Me.Name = "frmCertificateReprint"
    Me.ShowInTaskbar = False
    Me.Text = "Reprint Exam Certificate"
    Me.bpl.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents epl As CDBNETCL.EditPanel
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents cmdOk As System.Windows.Forms.Button
  Friend WithEvents bpl As CDBNETCL.ButtonPanel
End Class
