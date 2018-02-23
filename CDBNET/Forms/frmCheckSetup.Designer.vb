<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCheckSetup
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCheckSetup))
    Me.epl = New CDBNETCL.EditPanel()
    Me.ButtonPanel1 = New CDBNETCL.ButtonPanel()
    Me.cmdOk = New System.Windows.Forms.Button()
    Me.cmdCancel = New System.Windows.Forms.Button()
    Me.ButtonPanel1.SuspendLayout()
    Me.SuspendLayout()
    '
    'epl
    '
    Me.epl.AddressChanged = False
    Me.epl.BackColor = System.Drawing.Color.Transparent
    Me.epl.DataChanged = False
    Me.epl.Dock = System.Windows.Forms.DockStyle.Fill
    Me.epl.Location = New System.Drawing.Point(0, 0)
    Me.epl.Name = "epl"
    Me.epl.Recipients = Nothing
    Me.epl.Size = New System.Drawing.Size(436, 410)
    Me.epl.SuppressDrawing = False
    Me.epl.TabIndex = 0
    Me.epl.TabSelectedIndex = 0
    '
    'ButtonPanel1
    '
    Me.ButtonPanel1.Controls.Add(Me.cmdOk)
    Me.ButtonPanel1.Controls.Add(Me.cmdCancel)
    Me.ButtonPanel1.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.ButtonPanel1.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.ButtonPanel1.Location = New System.Drawing.Point(0, 371)
    Me.ButtonPanel1.Name = "ButtonPanel1"
    Me.ButtonPanel1.Size = New System.Drawing.Size(436, 39)
    Me.ButtonPanel1.TabIndex = 1
    '
    'cmdOk
    '
    Me.cmdOk.Location = New System.Drawing.Point(114, 6)
    Me.cmdOk.Name = "cmdOk"
    Me.cmdOk.Size = New System.Drawing.Size(96, 27)
    Me.cmdOk.TabIndex = 0
    Me.cmdOk.Text = "OK"
    Me.cmdOk.UseVisualStyleBackColor = True
    '
    'cmdCancel
    '
    Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdCancel.Location = New System.Drawing.Point(225, 6)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(96, 27)
    Me.cmdCancel.TabIndex = 1
    Me.cmdCancel.Text = "Cancel"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'frmCheckSetup
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.cmdCancel
    Me.ClientSize = New System.Drawing.Size(436, 410)
    Me.Controls.Add(Me.ButtonPanel1)
    Me.Controls.Add(Me.epl)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmCheckSetup"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    Me.Text = "Check Setup"
    Me.ButtonPanel1.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents epl As CDBNETCL.EditPanel
  Friend WithEvents ButtonPanel1 As CDBNETCL.ButtonPanel
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents cmdOk As System.Windows.Forms.Button
End Class
