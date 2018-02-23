<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmModuleContent
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmModuleContent))
    Me.cmdOK = New System.Windows.Forms.Button
    Me.cmdCancel = New System.Windows.Forms.Button
    Me.dgr = New CDBNETCL.DisplayGrid
    Me.bpl = New CDBNETCL.ButtonPanel
    Me.cmdRevert = New System.Windows.Forms.Button
    Me.bpl.SuspendLayout()
    Me.SuspendLayout()
    '
    'cmdOK
    '
    Me.cmdOK.DialogResult = System.Windows.Forms.DialogResult.OK
    Me.cmdOK.Location = New System.Drawing.Point(193, 6)
    Me.cmdOK.Margin = New System.Windows.Forms.Padding(2)
    Me.cmdOK.Name = "cmdOK"
    Me.cmdOK.Size = New System.Drawing.Size(96, 27)
    Me.cmdOK.TabIndex = 1
    Me.cmdOK.Text = "OK"
    Me.cmdOK.UseVisualStyleBackColor = True
    '
    'cmdCancel
    '
    Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdCancel.Location = New System.Drawing.Point(304, 6)
    Me.cmdCancel.Margin = New System.Windows.Forms.Padding(2)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(96, 27)
    Me.cmdCancel.TabIndex = 0
    Me.cmdCancel.Text = "Cancel"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'dgr
    '
    Me.dgr.AccessibleDescription = "Display List"
    Me.dgr.AccessibleName = "Display List"
    Me.dgr.AccessibleRole = System.Windows.Forms.AccessibleRole.Table
    Me.dgr.AllowDrop = True
    Me.dgr.AllowSorting = True
    Me.dgr.AutoSetHeight = False
    Me.dgr.AutoSetRowHeight = False
    Me.dgr.DisplayTitle = Nothing
    Me.dgr.Dock = System.Windows.Forms.DockStyle.Fill
    Me.dgr.Location = New System.Drawing.Point(0, 0)
    Me.dgr.Margin = New System.Windows.Forms.Padding(2)
    Me.dgr.MaxGridRows = 8
    Me.dgr.MultipleSelect = False
    Me.dgr.Name = "dgr"
    Me.dgr.ShowIfEmpty = False
    Me.dgr.Size = New System.Drawing.Size(482, 249)
    Me.dgr.TabIndex = 3
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdRevert)
    Me.bpl.Controls.Add(Me.cmdOK)
    Me.bpl.Controls.Add(Me.cmdCancel)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bpl.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bpl.Location = New System.Drawing.Point(0, 249)
    Me.bpl.Margin = New System.Windows.Forms.Padding(2)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(482, 39)
    Me.bpl.TabIndex = 2
    '
    'cmdRevert
    '
    Me.cmdRevert.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdRevert.Location = New System.Drawing.Point(82, 6)
    Me.cmdRevert.Margin = New System.Windows.Forms.Padding(2)
    Me.cmdRevert.Name = "cmdRevert"
    Me.cmdRevert.Size = New System.Drawing.Size(96, 27)
    Me.cmdRevert.TabIndex = 2
    Me.cmdRevert.Text = "&Revert"
    Me.cmdRevert.UseVisualStyleBackColor = True
    '
    'frmModuleContent
    '
    Me.AcceptButton = Me.cmdOK
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.cmdCancel
    Me.ClientSize = New System.Drawing.Size(482, 288)
    Me.Controls.Add(Me.dgr)
    Me.Controls.Add(Me.bpl)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Margin = New System.Windows.Forms.Padding(2)
    Me.Name = "frmModuleContent"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.Text = "Module Content"
    Me.bpl.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents bpl As CDBNETCL.ButtonPanel
  Friend WithEvents cmdOK As System.Windows.Forms.Button
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents dgr As CDBNETCL.DisplayGrid
  Friend WithEvents cmdRevert As System.Windows.Forms.Button
End Class
