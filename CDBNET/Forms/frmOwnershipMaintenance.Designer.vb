<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmOwnershipMaintenance
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
    Me.components = New System.ComponentModel.Container()
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmOwnershipMaintenance))
    Me.tvw = New CDBNETCL.VistaTreeView()
    Me.dgr = New CDBNETCL.DisplayGrid()
    Me.bpl = New CDBNETCL.ButtonPanel()
    Me.cmdAmend = New System.Windows.Forms.Button()
    Me.cmdNew = New System.Windows.Forms.Button()
    Me.cmdDelete = New System.Windows.Forms.Button()
    Me.cmdOK = New System.Windows.Forms.Button()
    Me.Splitter1 = New System.Windows.Forms.Splitter()
    Me.bpl.SuspendLayout()
    Me.SuspendLayout()
    '
    'tvw
    '
    Me.tvw.BackColor = System.Drawing.Color.FromArgb(CType(CType(232, Byte), Integer), CType(CType(232, Byte), Integer), CType(CType(232, Byte), Integer))
    Me.tvw.Dock = System.Windows.Forms.DockStyle.Left
    Me.tvw.FontHotTracking = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tvw.Location = New System.Drawing.Point(0, 0)
    Me.tvw.Name = "tvw"
    Me.tvw.Size = New System.Drawing.Size(198, 467)
    Me.tvw.TabIndex = 0
    '
    'dgr
    '
    Me.dgr.AccessibleName = "Display Grid"
    Me.dgr.ActiveColumn = 0
    Me.dgr.AllowSorting = True
    Me.dgr.AutoSetHeight = False
    Me.dgr.AutoSetRowHeight = False
    Me.dgr.DisplayTitle = Nothing
    Me.dgr.Dock = System.Windows.Forms.DockStyle.Fill
    Me.dgr.Location = New System.Drawing.Point(203, 0)
    Me.dgr.MaintenanceDesc = Nothing
    Me.dgr.MaxGridRows = 10
    Me.dgr.MultipleSelect = False
    Me.dgr.Name = "dgr"
    Me.dgr.RowCount = 10
    Me.dgr.ShowIfEmpty = False
    Me.dgr.Size = New System.Drawing.Size(556, 467)
    Me.dgr.SuppressHyperLinkFormat = False
    Me.dgr.TabIndex = 1
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdAmend)
    Me.bpl.Controls.Add(Me.cmdNew)
    Me.bpl.Controls.Add(Me.cmdDelete)
    Me.bpl.Controls.Add(Me.cmdOK)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bpl.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bpl.Location = New System.Drawing.Point(0, 467)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(759, 39)
    Me.bpl.TabIndex = 2
    '
    'cmdAmend
    '
    Me.cmdAmend.Enabled = False
    Me.cmdAmend.Location = New System.Drawing.Point(165, 6)
    Me.cmdAmend.Name = "cmdAmend"
    Me.cmdAmend.Size = New System.Drawing.Size(96, 27)
    Me.cmdAmend.TabIndex = 0
    Me.cmdAmend.Text = "&Amend"
    Me.cmdAmend.UseVisualStyleBackColor = True
    '
    'cmdNew
    '
    Me.cmdNew.Enabled = False
    Me.cmdNew.Location = New System.Drawing.Point(276, 6)
    Me.cmdNew.Name = "cmdNew"
    Me.cmdNew.Size = New System.Drawing.Size(96, 27)
    Me.cmdNew.TabIndex = 1
    Me.cmdNew.Text = "&New"
    Me.cmdNew.UseVisualStyleBackColor = True
    '
    'cmdDelete
    '
    Me.cmdDelete.Enabled = False
    Me.cmdDelete.Location = New System.Drawing.Point(387, 6)
    Me.cmdDelete.Name = "cmdDelete"
    Me.cmdDelete.Size = New System.Drawing.Size(96, 27)
    Me.cmdDelete.TabIndex = 2
    Me.cmdDelete.Text = "&Delete"
    Me.cmdDelete.UseVisualStyleBackColor = True
    '
    'cmdOK
    '
    Me.cmdOK.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdOK.Location = New System.Drawing.Point(498, 6)
    Me.cmdOK.Name = "cmdOK"
    Me.cmdOK.Size = New System.Drawing.Size(96, 27)
    Me.cmdOK.TabIndex = 3
    Me.cmdOK.Text = "Close"
    Me.cmdOK.UseVisualStyleBackColor = True
    '
    'Splitter1
    '
    Me.Splitter1.Location = New System.Drawing.Point(198, 0)
    Me.Splitter1.Name = "Splitter1"
    Me.Splitter1.Size = New System.Drawing.Size(5, 467)
    Me.Splitter1.TabIndex = 3
    Me.Splitter1.TabStop = False
    '
    'frmOwnershipMaintenance
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.cmdOK
    Me.ClientSize = New System.Drawing.Size(759, 506)
    Me.Controls.Add(Me.dgr)
    Me.Controls.Add(Me.Splitter1)
    Me.Controls.Add(Me.tvw)
    Me.Controls.Add(Me.bpl)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmOwnershipMaintenance"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    Me.Text = "Ownership Maintenance"
    Me.bpl.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents tvw As CDBNETCL.VistaTreeView
  Friend WithEvents dgr As CDBNETCL.DisplayGrid
  Friend WithEvents bpl As CDBNETCL.ButtonPanel
  Friend WithEvents cmdOK As System.Windows.Forms.Button
  Friend WithEvents cmdAmend As System.Windows.Forms.Button
  Friend WithEvents cmdNew As System.Windows.Forms.Button
  Friend WithEvents cmdDelete As System.Windows.Forms.Button
  Friend WithEvents Splitter1 As System.Windows.Forms.Splitter
End Class
