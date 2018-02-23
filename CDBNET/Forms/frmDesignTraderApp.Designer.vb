<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDesignTraderApp
  Inherits ThemedForm

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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDesignTraderApp))
    Me.epl = New CDBNETCL.EditPanel()
    Me.ButtonPanel1 = New CDBNETCL.ButtonPanel()
    Me.cmdClose = New System.Windows.Forms.Button()
    Me.cmdRevert = New System.Windows.Forms.Button()
    Me.cmdPrevious = New System.Windows.Forms.Button()
    Me.cmdNext = New System.Windows.Forms.Button()
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
    Me.epl.Size = New System.Drawing.Size(726, 500)
    Me.epl.SuppressDrawing = False
    Me.epl.TabIndex = 0
    Me.epl.TabSelectedIndex = 0
    '
    'ButtonPanel1
    '
    Me.ButtonPanel1.Controls.Add(Me.cmdClose)
    Me.ButtonPanel1.Controls.Add(Me.cmdRevert)
    Me.ButtonPanel1.Controls.Add(Me.cmdPrevious)
    Me.ButtonPanel1.Controls.Add(Me.cmdNext)
    Me.ButtonPanel1.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.ButtonPanel1.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.ButtonPanel1.Location = New System.Drawing.Point(0, 461)
    Me.ButtonPanel1.Name = "ButtonPanel1"
    Me.ButtonPanel1.Size = New System.Drawing.Size(726, 39)
    Me.ButtonPanel1.TabIndex = 1
    '
    'cmdClose
    '
    Me.cmdClose.Location = New System.Drawing.Point(148, 6)
    Me.cmdClose.Name = "cmdClose"
    Me.cmdClose.Size = New System.Drawing.Size(96, 27)
    Me.cmdClose.TabIndex = 3
    Me.cmdClose.Text = "Close"
    Me.cmdClose.UseVisualStyleBackColor = True
    '
    'cmdRevert
    '
    Me.cmdRevert.Location = New System.Drawing.Point(259, 6)
    Me.cmdRevert.Name = "cmdRevert"
    Me.cmdRevert.Size = New System.Drawing.Size(96, 27)
    Me.cmdRevert.TabIndex = 2
    Me.cmdRevert.Text = "Revert"
    Me.cmdRevert.UseVisualStyleBackColor = True
    '
    'cmdPrevious
    '
    Me.cmdPrevious.Location = New System.Drawing.Point(370, 6)
    Me.cmdPrevious.Name = "cmdPrevious"
    Me.cmdPrevious.Size = New System.Drawing.Size(96, 27)
    Me.cmdPrevious.TabIndex = 1
    Me.cmdPrevious.Text = "&Previous"
    Me.cmdPrevious.UseVisualStyleBackColor = True
    '
    'cmdNext
    '
    Me.cmdNext.Location = New System.Drawing.Point(481, 6)
    Me.cmdNext.Name = "cmdNext"
    Me.cmdNext.Size = New System.Drawing.Size(96, 27)
    Me.cmdNext.TabIndex = 0
    Me.cmdNext.Text = "&Next"
    Me.cmdNext.UseVisualStyleBackColor = True
    '
    'frmDesignTraderApp
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(726, 500)
    Me.Controls.Add(Me.ButtonPanel1)
    Me.Controls.Add(Me.epl)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmDesignTraderApp"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    Me.Text = "frmDesignTraderApp"
    Me.ButtonPanel1.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents epl As CDBNETCL.EditPanel
  Friend WithEvents ButtonPanel1 As CDBNETCL.ButtonPanel
  Friend WithEvents cmdNext As System.Windows.Forms.Button
  Friend WithEvents cmdPrevious As System.Windows.Forms.Button
  Friend WithEvents cmdRevert As System.Windows.Forms.Button
  Friend WithEvents cmdClose As System.Windows.Forms.Button
End Class
