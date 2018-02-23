<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmBatchProcessing
  Inherits CDBNETCL.MaintenanceParentForm

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
    Me.components = New System.ComponentModel.Container
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBatchProcessing))
    Me.tab = New CDBNETCL.TabControl
    Me.tbpOutstanding = New System.Windows.Forms.TabPage
    Me.tbpIncomplete = New System.Windows.Forms.TabPage
    Me.tbpDetailComplete = New System.Windows.Forms.TabPage
    Me.tbpPickingList = New System.Windows.Forms.TabPage
    Me.tbpConfirmStock = New System.Windows.Forms.TabPage
    Me.tbpChequeList = New System.Windows.Forms.TabPage
    Me.tbpCreateClaim = New System.Windows.Forms.TabPage
    Me.tbpPrintPayingInSlips = New System.Windows.Forms.TabPage
    Me.tbpPostToCashBook = New System.Windows.Forms.TabPage
    Me.tbpPostBatch = New System.Windows.Forms.TabPage
    Me.cmdRefresh = New System.Windows.Forms.Button
    Me.cmdDelete = New System.Windows.Forms.Button
    Me.cmdClose = New System.Windows.Forms.Button
    Me.cmdProcess = New System.Windows.Forms.Button
    Me.cmdDetails = New System.Windows.Forms.Button
    Me.bpl = New CDBNETCL.ButtonPanel
    Me.dgr = New CDBNETCL.DisplayGrid
    Me.ss = New System.Windows.Forms.StatusStrip
    Me.tssl = New System.Windows.Forms.ToolStripStatusLabel
    Me.tab.SuspendLayout()
    Me.bpl.SuspendLayout()
    Me.ss.SuspendLayout()
    Me.SuspendLayout()
    '
    'tab
    '
    Me.tab.Controls.Add(Me.tbpOutstanding)
    Me.tab.Controls.Add(Me.tbpIncomplete)
    Me.tab.Controls.Add(Me.tbpDetailComplete)
    Me.tab.Controls.Add(Me.tbpPickingList)
    Me.tab.Controls.Add(Me.tbpConfirmStock)
    Me.tab.Controls.Add(Me.tbpChequeList)
    Me.tab.Controls.Add(Me.tbpCreateClaim)
    Me.tab.Controls.Add(Me.tbpPrintPayingInSlips)
    Me.tab.Controls.Add(Me.tbpPostToCashBook)
    Me.tab.Controls.Add(Me.tbpPostBatch)
    Me.tab.Dock = System.Windows.Forms.DockStyle.Top
    Me.tab.DrawMode = System.Windows.Forms.TabDrawMode.OwnerDrawFixed
    Me.tab.ItemSize = New System.Drawing.Size(112, 22)
    Me.tab.Location = New System.Drawing.Point(0, 0)
    Me.tab.Name = "tab"
    Me.tab.SelectedIndex = 0
    Me.tab.Size = New System.Drawing.Size(1028, 22)
    Me.tab.SizeMode = System.Windows.Forms.TabSizeMode.Fixed
    Me.tab.TabIndex = 3
    '
    'tbpOutstanding
    '
    Me.tbpOutstanding.Location = New System.Drawing.Point(4, 26)
    Me.tbpOutstanding.Name = "tbpOutstanding"
    Me.tbpOutstanding.Size = New System.Drawing.Size(1020, 0)
    Me.tbpOutstanding.TabIndex = 0
    Me.tbpOutstanding.Text = "Outstanding"
    Me.tbpOutstanding.ToolTipText = "Outstanding Batches"
    Me.tbpOutstanding.UseVisualStyleBackColor = True
    '
    'tbpIncomplete
    '
    Me.tbpIncomplete.Location = New System.Drawing.Point(4, 24)
    Me.tbpIncomplete.Name = "tbpIncomplete"
    Me.tbpIncomplete.Size = New System.Drawing.Size(192, 72)
    Me.tbpIncomplete.TabIndex = 1
    Me.tbpIncomplete.Text = "Incomplete"
    Me.tbpIncomplete.ToolTipText = "Incomplete Batches"
    Me.tbpIncomplete.UseVisualStyleBackColor = True
    '
    'tbpDetailComplete
    '
    Me.tbpDetailComplete.Location = New System.Drawing.Point(4, 24)
    Me.tbpDetailComplete.Name = "tbpDetailComplete"
    Me.tbpDetailComplete.Size = New System.Drawing.Size(192, 72)
    Me.tbpDetailComplete.TabIndex = 2
    Me.tbpDetailComplete.Text = "Detail Complete"
    Me.tbpDetailComplete.ToolTipText = "Set Detail Complete"
    Me.tbpDetailComplete.UseVisualStyleBackColor = True
    '
    'tbpPickingList
    '
    Me.tbpPickingList.Location = New System.Drawing.Point(4, 24)
    Me.tbpPickingList.Name = "tbpPickingList"
    Me.tbpPickingList.Size = New System.Drawing.Size(192, 72)
    Me.tbpPickingList.TabIndex = 3
    Me.tbpPickingList.Text = "Picking List"
    Me.tbpPickingList.ToolTipText = "Picking List Production"
    Me.tbpPickingList.UseVisualStyleBackColor = True
    '
    'tbpConfirmStock
    '
    Me.tbpConfirmStock.Location = New System.Drawing.Point(4, 24)
    Me.tbpConfirmStock.Name = "tbpConfirmStock"
    Me.tbpConfirmStock.Size = New System.Drawing.Size(192, 72)
    Me.tbpConfirmStock.TabIndex = 4
    Me.tbpConfirmStock.Text = "Confirm Stock"
    Me.tbpConfirmStock.ToolTipText = "Confirm Stock Allocation"
    Me.tbpConfirmStock.UseVisualStyleBackColor = True
    '
    'tbpChequeList
    '
    Me.tbpChequeList.Location = New System.Drawing.Point(4, 24)
    Me.tbpChequeList.Name = "tbpChequeList"
    Me.tbpChequeList.Size = New System.Drawing.Size(192, 72)
    Me.tbpChequeList.TabIndex = 5
    Me.tbpChequeList.Text = "Cheque List"
    Me.tbpChequeList.ToolTipText = "Cheque List Production"
    Me.tbpChequeList.UseVisualStyleBackColor = True
    '
    'tbpCreateClaim
    '
    Me.tbpCreateClaim.Location = New System.Drawing.Point(4, 24)
    Me.tbpCreateClaim.Name = "tbpCreateClaim"
    Me.tbpCreateClaim.Size = New System.Drawing.Size(192, 72)
    Me.tbpCreateClaim.TabIndex = 6
    Me.tbpCreateClaim.Text = "Create Claim"
    Me.tbpCreateClaim.ToolTipText = "Create Claim File"
    Me.tbpCreateClaim.UseVisualStyleBackColor = True
    '
    'tbpPrintPayingInSlips
    '
    Me.tbpPrintPayingInSlips.Location = New System.Drawing.Point(4, 24)
    Me.tbpPrintPayingInSlips.Name = "tbpPrintPayingInSlips"
    Me.tbpPrintPayingInSlips.Size = New System.Drawing.Size(192, 72)
    Me.tbpPrintPayingInSlips.TabIndex = 7
    Me.tbpPrintPayingInSlips.Text = "Paying In Slips"
    Me.tbpPrintPayingInSlips.ToolTipText = "Print Paying In Slips"
    Me.tbpPrintPayingInSlips.UseVisualStyleBackColor = True
    '
    'tbpPostToCashBook
    '
    Me.tbpPostToCashBook.Location = New System.Drawing.Point(4, 24)
    Me.tbpPostToCashBook.Name = "tbpPostToCashBook"
    Me.tbpPostToCashBook.Size = New System.Drawing.Size(192, 72)
    Me.tbpPostToCashBook.TabIndex = 8
    Me.tbpPostToCashBook.Text = "Cash Book"
    Me.tbpPostToCashBook.ToolTipText = "Post to Cash Book"
    Me.tbpPostToCashBook.UseVisualStyleBackColor = True
    '
    'tbpPostBatch
    '
    Me.tbpPostBatch.Location = New System.Drawing.Point(4, 24)
    Me.tbpPostBatch.Name = "tbpPostBatch"
    Me.tbpPostBatch.Size = New System.Drawing.Size(192, 72)
    Me.tbpPostBatch.TabIndex = 9
    Me.tbpPostBatch.Text = "Post Batch"
    Me.tbpPostBatch.ToolTipText = "Post Batch"
    Me.tbpPostBatch.UseVisualStyleBackColor = True
    '
    'cmdRefresh
    '
    Me.cmdRefresh.Location = New System.Drawing.Point(577, 6)
    Me.cmdRefresh.Name = "cmdRefresh"
    Me.cmdRefresh.Size = New System.Drawing.Size(96, 27)
    Me.cmdRefresh.TabIndex = 1
    Me.cmdRefresh.Text = "&Refresh"
    Me.cmdRefresh.UseVisualStyleBackColor = True
    '
    'cmdDelete
    '
    Me.cmdDelete.Location = New System.Drawing.Point(466, 6)
    Me.cmdDelete.Name = "cmdDelete"
    Me.cmdDelete.Size = New System.Drawing.Size(96, 27)
    Me.cmdDelete.TabIndex = 2
    Me.cmdDelete.Text = "De&lete"
    Me.cmdDelete.UseVisualStyleBackColor = True
    '
    'cmdClose
    '
    Me.cmdClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdClose.Location = New System.Drawing.Point(688, 6)
    Me.cmdClose.Name = "cmdClose"
    Me.cmdClose.Size = New System.Drawing.Size(96, 27)
    Me.cmdClose.TabIndex = 0
    Me.cmdClose.Text = "&Close"
    Me.cmdClose.UseVisualStyleBackColor = True
    '
    'cmdProcess
    '
    Me.cmdProcess.Location = New System.Drawing.Point(355, 6)
    Me.cmdProcess.Name = "cmdProcess"
    Me.cmdProcess.Size = New System.Drawing.Size(96, 27)
    Me.cmdProcess.TabIndex = 3
    Me.cmdProcess.Text = "&Process"
    Me.cmdProcess.UseVisualStyleBackColor = True
    '
    'cmdDetails
    '
    Me.cmdDetails.Location = New System.Drawing.Point(244, 6)
    Me.cmdDetails.Name = "cmdDetails"
    Me.cmdDetails.Size = New System.Drawing.Size(96, 27)
    Me.cmdDetails.TabIndex = 4
    Me.cmdDetails.Text = "&Details"
    Me.cmdDetails.UseVisualStyleBackColor = True
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdDetails)
    Me.bpl.Controls.Add(Me.cmdProcess)
    Me.bpl.Controls.Add(Me.cmdDelete)
    Me.bpl.Controls.Add(Me.cmdRefresh)
    Me.bpl.Controls.Add(Me.cmdClose)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bpl.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bpl.Location = New System.Drawing.Point(0, 253)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(1028, 39)
    Me.bpl.TabIndex = 4
    '
    'dgr
    '
    Me.dgr.AccessibleDescription = "Display List"
    Me.dgr.AccessibleName = "Display List"
    Me.dgr.AccessibleRole = System.Windows.Forms.AccessibleRole.Table
    Me.dgr.AllowSorting = True
    Me.dgr.AutoSetHeight = False
    Me.dgr.AutoSetRowHeight = False
    Me.dgr.DisplayTitle = Nothing
    Me.dgr.Dock = System.Windows.Forms.DockStyle.Fill
    Me.dgr.Location = New System.Drawing.Point(0, 22)
    Me.dgr.MaintenanceDesc = Nothing
    Me.dgr.MaxGridRows = 8
    Me.dgr.MultipleSelect = False
    Me.dgr.Name = "dgr"
    Me.dgr.RowCount = 10
    Me.dgr.ShowIfEmpty = False
    Me.dgr.Size = New System.Drawing.Size(1028, 207)
    Me.dgr.TabIndex = 5
    '
    'ss
    '
    Me.ss.AutoSize = False
    Me.ss.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tssl})
    Me.ss.Location = New System.Drawing.Point(0, 229)
    Me.ss.Name = "ss"
    Me.ss.Size = New System.Drawing.Size(1028, 24)
    Me.ss.TabIndex = 6
    '
    'tssl
    '
    Me.tssl.Name = "tssl"
    Me.tssl.Size = New System.Drawing.Size(0, 19)
    Me.tssl.TextAlign = System.Drawing.ContentAlignment.BottomLeft
    '
    'frmBatchProcessing
    '
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
    Me.CancelButton = Me.cmdClose
    Me.ClientSize = New System.Drawing.Size(1028, 292)
    Me.Controls.Add(Me.dgr)
    Me.Controls.Add(Me.tab)
    Me.Controls.Add(Me.ss)
    Me.Controls.Add(Me.bpl)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmBatchProcessing"
    Me.Text = "Batch Processing"
    Me.tab.ResumeLayout(False)
    Me.bpl.ResumeLayout(False)
    Me.ss.ResumeLayout(False)
    Me.ss.PerformLayout()
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents tab As CDBNETCL.TabControl
  Friend WithEvents tbpOutstanding As System.Windows.Forms.TabPage
  Friend WithEvents tbpIncomplete As System.Windows.Forms.TabPage
  Friend WithEvents tbpPickingList As System.Windows.Forms.TabPage
  Friend WithEvents tbpConfirmStock As System.Windows.Forms.TabPage
  Friend WithEvents tbpChequeList As System.Windows.Forms.TabPage
  Friend WithEvents tbpCreateClaim As System.Windows.Forms.TabPage
  Friend WithEvents tbpPrintPayingInSlips As System.Windows.Forms.TabPage
  Friend WithEvents tbpPostToCashBook As System.Windows.Forms.TabPage
  Friend WithEvents tbpPostBatch As System.Windows.Forms.TabPage
  Friend WithEvents tbpDetailComplete As System.Windows.Forms.TabPage
  Friend WithEvents cmdRefresh As System.Windows.Forms.Button
  Friend WithEvents cmdDelete As System.Windows.Forms.Button
  Friend WithEvents cmdClose As System.Windows.Forms.Button
  Friend WithEvents cmdProcess As System.Windows.Forms.Button
  Friend WithEvents cmdDetails As System.Windows.Forms.Button
  Friend WithEvents bpl As CDBNETCL.ButtonPanel
  Friend WithEvents dgr As CDBNETCL.DisplayGrid
  Friend WithEvents ss As System.Windows.Forms.StatusStrip
  Friend WithEvents tssl As System.Windows.Forms.ToolStripStatusLabel

End Class
