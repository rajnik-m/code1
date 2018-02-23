<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDespatchTracking
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDespatchTracking))
    Me.lblPickingListNumber = New System.Windows.Forms.Label()
    Me.txtPickingListNumber = New System.Windows.Forms.TextBox()
    Me.lblWarehouse = New System.Windows.Forms.Label()
    Me.txtWarehouse = New System.Windows.Forms.TextBox()
    Me.txtWarehouseDesc = New System.Windows.Forms.TextBox()
    Me.epl = New CDBNETCL.EditPanel()
    Me.ToolStripTextBox1 = New System.Windows.Forms.ToolStripTextBox()
    Me.PanelEx1 = New CDBNETCL.PanelEx()
    Me.dgr = New CDBNETCL.DisplayGrid()
    Me.bpl = New CDBNETCL.ButtonPanel()
    Me.cmdClose = New System.Windows.Forms.Button()
    Me.cmdSave = New System.Windows.Forms.Button()
    Me.cmdClear = New System.Windows.Forms.Button()
    Me.spl = New System.Windows.Forms.SplitContainer()
    Me.PanelEx1.SuspendLayout()
    Me.bpl.SuspendLayout()
    CType(Me.spl, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.spl.Panel1.SuspendLayout()
    Me.spl.Panel2.SuspendLayout()
    Me.spl.SuspendLayout()
    Me.SuspendLayout()
    '
    'lblPickingListNumber
    '
    Me.lblPickingListNumber.AutoSize = True
    Me.lblPickingListNumber.Location = New System.Drawing.Point(14, 16)
    Me.lblPickingListNumber.Name = "lblPickingListNumber"
    Me.lblPickingListNumber.Size = New System.Drawing.Size(101, 13)
    Me.lblPickingListNumber.TabIndex = 0
    Me.lblPickingListNumber.Text = "PickingList Number:"
    '
    'txtPickingListNumber
    '
    Me.txtPickingListNumber.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
    Me.txtPickingListNumber.Location = New System.Drawing.Point(153, 14)
    Me.txtPickingListNumber.Name = "txtPickingListNumber"
    Me.txtPickingListNumber.Size = New System.Drawing.Size(81, 20)
    Me.txtPickingListNumber.TabIndex = 1
    '
    'lblWarehouse
    '
    Me.lblWarehouse.AutoSize = True
    Me.lblWarehouse.Location = New System.Drawing.Point(14, 43)
    Me.lblWarehouse.Name = "lblWarehouse"
    Me.lblWarehouse.Size = New System.Drawing.Size(65, 13)
    Me.lblWarehouse.TabIndex = 2
    Me.lblWarehouse.Text = "Warehouse:"
    '
    'txtWarehouse
    '
    Me.txtWarehouse.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
    Me.txtWarehouse.Location = New System.Drawing.Point(153, 40)
    Me.txtWarehouse.Name = "txtWarehouse"
    Me.txtWarehouse.ReadOnly = True
    Me.txtWarehouse.Size = New System.Drawing.Size(81, 20)
    Me.txtWarehouse.TabIndex = 16
    Me.txtWarehouse.TabStop = False
    '
    'txtWarehouseDesc
    '
    Me.txtWarehouseDesc.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
    Me.txtWarehouseDesc.Location = New System.Drawing.Point(240, 40)
    Me.txtWarehouseDesc.Name = "txtWarehouseDesc"
    Me.txtWarehouseDesc.ReadOnly = True
    Me.txtWarehouseDesc.Size = New System.Drawing.Size(240, 20)
    Me.txtWarehouseDesc.TabIndex = 17
    Me.txtWarehouseDesc.TabStop = False
    '
    'epl
    '
    Me.epl.AddressChanged = False
    Me.epl.BackColor = System.Drawing.Color.Transparent
    Me.epl.DataChanged = False
    Me.epl.Dock = System.Windows.Forms.DockStyle.Fill
    Me.epl.Location = New System.Drawing.Point(0, 0)
    Me.epl.Margin = New System.Windows.Forms.Padding(2)
    Me.epl.Name = "epl"
    Me.epl.Recipients = Nothing
    Me.epl.Size = New System.Drawing.Size(528, 236)
    Me.epl.SuppressDrawing = False
    Me.epl.TabIndex = 6
    Me.epl.TabSelectedIndex = 0
    '
    'ToolStripTextBox1
    '
    Me.ToolStripTextBox1.Name = "ToolStripTextBox1"
    Me.ToolStripTextBox1.Size = New System.Drawing.Size(100, 21)
    '
    'PanelEx1
    '
    Me.PanelEx1.BackColor = System.Drawing.Color.Transparent
    Me.PanelEx1.Controls.Add(Me.lblPickingListNumber)
    Me.PanelEx1.Controls.Add(Me.txtPickingListNumber)
    Me.PanelEx1.Controls.Add(Me.lblWarehouse)
    Me.PanelEx1.Controls.Add(Me.txtWarehouse)
    Me.PanelEx1.Controls.Add(Me.txtWarehouseDesc)
    Me.PanelEx1.Dock = System.Windows.Forms.DockStyle.Top
    Me.PanelEx1.Location = New System.Drawing.Point(0, 0)
    Me.PanelEx1.Name = "PanelEx1"
    Me.PanelEx1.Size = New System.Drawing.Size(528, 68)
    Me.PanelEx1.TabIndex = 8
    '
    'dgr
    '
    Me.dgr.AccessibleName = "Display Grid"
    Me.dgr.ActiveColumn = 0
    Me.dgr.AllowSorting = True
    Me.dgr.AutoSetHeight = True
    Me.dgr.AutoSetRowHeight = False
    Me.dgr.AutoSize = True
    Me.dgr.DisplayTitle = Nothing
    Me.dgr.Dock = System.Windows.Forms.DockStyle.Fill
    Me.dgr.Location = New System.Drawing.Point(0, 0)
    Me.dgr.MaintenanceDesc = Nothing
    Me.dgr.MaxGridRows = 8
    Me.dgr.MultipleSelect = False
    Me.dgr.Name = "dgr"
    Me.dgr.RowCount = 10
    Me.dgr.ShowIfEmpty = False
    Me.dgr.Size = New System.Drawing.Size(528, 124)
    Me.dgr.SuppressHyperLinkFormat = False
    Me.dgr.TabIndex = 2
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdClose)
    Me.bpl.Controls.Add(Me.cmdSave)
    Me.bpl.Controls.Add(Me.cmdClear)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bpl.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bpl.Location = New System.Drawing.Point(0, 432)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(528, 39)
    Me.bpl.TabIndex = 7
    '
    'cmdClose
    '
    Me.cmdClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdClose.Location = New System.Drawing.Point(105, 6)
    Me.cmdClose.Name = "cmdClose"
    Me.cmdClose.Size = New System.Drawing.Size(96, 27)
    Me.cmdClose.TabIndex = 2
    Me.cmdClose.Text = "Close"
    Me.cmdClose.UseVisualStyleBackColor = True
    '
    'cmdSave
    '
    Me.cmdSave.Enabled = False
    Me.cmdSave.Location = New System.Drawing.Point(216, 6)
    Me.cmdSave.Name = "cmdSave"
    Me.cmdSave.Size = New System.Drawing.Size(96, 27)
    Me.cmdSave.TabIndex = 1
    Me.cmdSave.Text = "&Save"
    Me.cmdSave.UseVisualStyleBackColor = True
    '
    'cmdClear
    '
    Me.cmdClear.Location = New System.Drawing.Point(327, 6)
    Me.cmdClear.Name = "cmdClear"
    Me.cmdClear.Size = New System.Drawing.Size(96, 27)
    Me.cmdClear.TabIndex = 0
    Me.cmdClear.Text = "&Clear"
    Me.cmdClear.UseVisualStyleBackColor = True
    '
    'spl
    '
    Me.spl.Dock = System.Windows.Forms.DockStyle.Fill
    Me.spl.FixedPanel = System.Windows.Forms.FixedPanel.Panel2
    Me.spl.Location = New System.Drawing.Point(0, 68)
    Me.spl.Margin = New System.Windows.Forms.Padding(1)
    Me.spl.Name = "spl"
    Me.spl.Orientation = System.Windows.Forms.Orientation.Horizontal
    '
    'spl.Panel1
    '
    Me.spl.Panel1.Controls.Add(Me.dgr)
    '
    'spl.Panel2
    '
    Me.spl.Panel2.Controls.Add(Me.epl)
    Me.spl.Panel2.Margin = New System.Windows.Forms.Padding(2)
    Me.spl.Size = New System.Drawing.Size(528, 364)
    Me.spl.SplitterDistance = 124
    Me.spl.TabIndex = 9
    '
    'frmDespatchTracking
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.cmdClose
    Me.ClientSize = New System.Drawing.Size(528, 471)
    Me.Controls.Add(Me.spl)
    Me.Controls.Add(Me.bpl)
    Me.Controls.Add(Me.PanelEx1)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.MinimumSize = New System.Drawing.Size(500, 480)
    Me.Name = "frmDespatchTracking"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.Text = "Despatch Tracking"
    Me.PanelEx1.ResumeLayout(False)
    Me.PanelEx1.PerformLayout()
    Me.bpl.ResumeLayout(False)
    Me.spl.Panel1.ResumeLayout(False)
    Me.spl.Panel1.PerformLayout()
    Me.spl.Panel2.ResumeLayout(False)
    CType(Me.spl, System.ComponentModel.ISupportInitialize).EndInit()
    Me.spl.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents lblPickingListNumber As System.Windows.Forms.Label
  Friend WithEvents txtPickingListNumber As System.Windows.Forms.TextBox
  Friend WithEvents lblWarehouse As System.Windows.Forms.Label
  Friend WithEvents txtWarehouse As System.Windows.Forms.TextBox
  Friend WithEvents txtWarehouseDesc As System.Windows.Forms.TextBox
  Friend WithEvents epl As CDBNETCL.EditPanel
  Friend WithEvents ToolStripTextBox1 As System.Windows.Forms.ToolStripTextBox
  Friend WithEvents PanelEx1 As CDBNETCL.PanelEx
  Friend WithEvents dgr As CDBNETCL.DisplayGrid
  Friend WithEvents bpl As CDBNETCL.ButtonPanel
  Friend WithEvents cmdClose As System.Windows.Forms.Button
  Friend WithEvents cmdSave As System.Windows.Forms.Button
  Friend WithEvents cmdClear As System.Windows.Forms.Button
  Friend WithEvents spl As System.Windows.Forms.SplitContainer
End Class
