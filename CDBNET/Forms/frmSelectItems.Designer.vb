<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSelectItems
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSelectItems))
    Me.cmdOK = New System.Windows.Forms.Button
    Me.cmdCancel = New System.Windows.Forms.Button
    Me.dgr = New CDBNETCL.DisplayGrid
    Me.bpl = New CDBNETCL.ButtonPanel
    Me.cmdSelectAll = New System.Windows.Forms.Button
    Me.cmdClearAll = New System.Windows.Forms.Button
    Me.bpl.SuspendLayout()
    Me.SuspendLayout()
    '
    'cmdOK
    '
    Me.cmdOK.DialogResult = System.Windows.Forms.DialogResult.OK
    Me.cmdOK.Location = New System.Drawing.Point(60, 6)
    Me.cmdOK.Name = "cmdOK"
    Me.cmdOK.Size = New System.Drawing.Size(94, 27)
    Me.cmdOK.TabIndex = 2
    Me.cmdOK.Text = "OK"
    Me.cmdOK.UseVisualStyleBackColor = True
    '
    'cmdCancel
    '
    Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdCancel.Location = New System.Drawing.Point(387, 6)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(94, 27)
    Me.cmdCancel.TabIndex = 0
    Me.cmdCancel.Text = "Cancel"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'dgr
    '
    Me.dgr.AutoSetHeight = False
    Me.dgr.AutoSetRowHeight = False
    Me.dgr.DisplayTitle = Nothing
    Me.dgr.Dock = System.Windows.Forms.DockStyle.Fill
    Me.dgr.Location = New System.Drawing.Point(0, 0)
    Me.dgr.MaxGridRows = 8
    Me.dgr.Name = "dgr"
    Me.dgr.ShowIfEmpty = False
    Me.dgr.Size = New System.Drawing.Size(541, 192)
    Me.dgr.TabIndex = 1
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdOK)
    Me.bpl.Controls.Add(Me.cmdSelectAll)
    Me.bpl.Controls.Add(Me.cmdClearAll)
    Me.bpl.Controls.Add(Me.cmdCancel)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bpl.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bpl.Location = New System.Drawing.Point(0, 192)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(541, 39)
    Me.bpl.TabIndex = 0
    '
    'cmdSelectAll
    '
    Me.cmdSelectAll.Location = New System.Drawing.Point(169, 6)
    Me.cmdSelectAll.Name = "cmdSelectAll"
    Me.cmdSelectAll.Size = New System.Drawing.Size(94, 27)
    Me.cmdSelectAll.TabIndex = 1
    Me.cmdSelectAll.Text = "Select All"
    Me.cmdSelectAll.UseVisualStyleBackColor = True
    '
    'cmdClearAll
    '
    Me.cmdClearAll.Location = New System.Drawing.Point(278, 6)
    Me.cmdClearAll.Name = "cmdClearAll"
    Me.cmdClearAll.Size = New System.Drawing.Size(94, 27)
    Me.cmdClearAll.TabIndex = 3
    Me.cmdClearAll.Text = "Clear All"
    Me.cmdClearAll.UseVisualStyleBackColor = True
    '
    'frmSelectItems
    '
    Me.AcceptButton = Me.cmdOK
    Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.cmdCancel
    Me.ClientSize = New System.Drawing.Size(541, 231)
    Me.Controls.Add(Me.dgr)
    Me.Controls.Add(Me.bpl)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmSelectItems"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.bpl.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents bpl As CDBNETCL.ButtonPanel
  Friend WithEvents cmdSelectAll As System.Windows.Forms.Button
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents dgr As CDBNETCL.DisplayGrid
  Friend WithEvents cmdOK As System.Windows.Forms.Button
  Friend WithEvents cmdClearAll As System.Windows.Forms.Button
End Class
