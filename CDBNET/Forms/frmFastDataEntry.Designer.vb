<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmFastDataEntry
  Inherits CDBNETCL.ThemedForm

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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmFastDataEntry))
    Me.bpl = New CDBNETCL.ButtonPanel()
    Me.cmdTest = New System.Windows.Forms.Button()
    Me.cmdOK = New System.Windows.Forms.Button()
    Me.cmdNext = New System.Windows.Forms.Button()
    Me.cmdCancel = New System.Windows.Forms.Button()
    Me.pnl = New System.Windows.Forms.Panel()
    Me.prgBar = New System.Windows.Forms.ProgressBar()
    Me.bpl.SuspendLayout()
    Me.SuspendLayout()
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdTest)
    Me.bpl.Controls.Add(Me.cmdOK)
    Me.bpl.Controls.Add(Me.cmdNext)
    Me.bpl.Controls.Add(Me.cmdCancel)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bpl.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bpl.Location = New System.Drawing.Point(0, 343)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(561, 39)
    Me.bpl.TabIndex = 1
    '
    'cmdTest
    '
    Me.cmdTest.Location = New System.Drawing.Point(66, 6)
    Me.cmdTest.Name = "cmdTest"
    Me.cmdTest.Size = New System.Drawing.Size(96, 27)
    Me.cmdTest.TabIndex = 0
    Me.cmdTest.Text = "&Test"
    Me.cmdTest.UseVisualStyleBackColor = True
    '
    'cmdOK
    '
    Me.cmdOK.Location = New System.Drawing.Point(177, 6)
    Me.cmdOK.Name = "cmdOK"
    Me.cmdOK.Size = New System.Drawing.Size(96, 27)
    Me.cmdOK.TabIndex = 1
    Me.cmdOK.Text = "&Submit"
    Me.cmdOK.UseVisualStyleBackColor = True
    '
    'cmdNext
    '
    Me.cmdNext.Location = New System.Drawing.Point(288, 6)
    Me.cmdNext.Name = "cmdNext"
    Me.cmdNext.Size = New System.Drawing.Size(96, 27)
    Me.cmdNext.TabIndex = 2
    Me.cmdNext.Text = "&Next"
    Me.cmdNext.UseVisualStyleBackColor = True
    '
    'cmdCancel
    '
    Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdCancel.Location = New System.Drawing.Point(399, 6)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(96, 27)
    Me.cmdCancel.TabIndex = 3
    Me.cmdCancel.Text = "&Cancel"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'pnl
    '
    Me.pnl.AutoScroll = True
    Me.pnl.Dock = System.Windows.Forms.DockStyle.Fill
    Me.pnl.Location = New System.Drawing.Point(0, 0)
    Me.pnl.Name = "pnl"
    Me.pnl.Size = New System.Drawing.Size(561, 343)
    Me.pnl.TabIndex = 0
    '
    'prgBar
    '
    Me.prgBar.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.prgBar.Location = New System.Drawing.Point(0, 329)
    Me.prgBar.Name = "prgBar"
    Me.prgBar.Size = New System.Drawing.Size(561, 14)
    Me.prgBar.Step = 1
    Me.prgBar.TabIndex = 0
    Me.prgBar.Visible = False
    '
    'frmFastDataEntry
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.cmdCancel
    Me.ClientSize = New System.Drawing.Size(561, 382)
    Me.Controls.Add(Me.prgBar)
    Me.Controls.Add(Me.pnl)
    Me.Controls.Add(Me.bpl)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmFastDataEntry"
    Me.ShowInTaskbar = False
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.Text = "Fast Data Entry"
    Me.bpl.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents bpl As CDBNETCL.ButtonPanel
  Friend WithEvents cmdOK As System.Windows.Forms.Button
  Friend WithEvents cmdNext As System.Windows.Forms.Button
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents pnl As System.Windows.Forms.Panel
  Friend WithEvents cmdTest As System.Windows.Forms.Button
  Friend WithEvents prgBar As System.Windows.Forms.ProgressBar
End Class
