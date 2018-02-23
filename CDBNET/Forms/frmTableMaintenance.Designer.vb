<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTableMaintenance
  Inherits CDBNETCL.PersistentForm

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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmTableMaintenance))
    Me.cboTables = New System.Windows.Forms.ComboBox()
    Me.bpl = New CDBNETCL.ButtonPanel()
    Me.cmdSubTable1 = New System.Windows.Forms.Button()
    Me.cmdSubTable2 = New System.Windows.Forms.Button()
    Me.cmdSubTable3 = New System.Windows.Forms.Button()
    Me.cmdSubTable4 = New System.Windows.Forms.Button()
    Me.cmdSave = New System.Windows.Forms.Button()
    Me.cmdDelete = New System.Windows.Forms.Button()
    Me.cmdAmend = New System.Windows.Forms.Button()
    Me.cmdMore = New System.Windows.Forms.Button()
    Me.cmdShowTable = New System.Windows.Forms.Button()
    Me.cmdNew = New System.Windows.Forms.Button()
    Me.cmdSelect = New System.Windows.Forms.Button()
    Me.cmdClose = New System.Windows.Forms.Button()
    Me.cmdExport = New System.Windows.Forms.Button()
    Me.pnl = New CDBNETCL.PanelEx()
    Me.lblGroups = New System.Windows.Forms.Label()
    Me.cboGroups = New System.Windows.Forms.ComboBox()
    Me.dgr = New CDBNETCL.DisplayGrid()
    Me.tlp = New System.Windows.Forms.TableLayoutPanel()
    Me.lblTableNotes = New System.Windows.Forms.Label()
    Me.txtTableNotes = New System.Windows.Forms.TextBox()
    Me.txtDefaultValues = New System.Windows.Forms.TextBox()
    Me.txtAdminNotes = New System.Windows.Forms.TextBox()
    Me.lblAdminNotes = New System.Windows.Forms.Label()
    Me.lblDefaultValues = New System.Windows.Forms.Label()
    Me.lblContents = New System.Windows.Forms.Label()
    Me.lblTableDesc = New System.Windows.Forms.Label()
    Me.bpl.SuspendLayout()
    Me.pnl.SuspendLayout()
    Me.tlp.SuspendLayout()
    Me.SuspendLayout()
    '
    'cboTables
    '
    Me.cboTables.FormattingEnabled = True
    Me.cboTables.Location = New System.Drawing.Point(135, 5)
    Me.cboTables.Name = "cboTables"
    Me.cboTables.Size = New System.Drawing.Size(247, 21)
    Me.cboTables.TabIndex = 0
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdSubTable1)
    Me.bpl.Controls.Add(Me.cmdSubTable2)
    Me.bpl.Controls.Add(Me.cmdSubTable3)
    Me.bpl.Controls.Add(Me.cmdSubTable4)
    Me.bpl.Controls.Add(Me.cmdSave)
    Me.bpl.Controls.Add(Me.cmdDelete)
    Me.bpl.Controls.Add(Me.cmdAmend)
    Me.bpl.Controls.Add(Me.cmdMore)
    Me.bpl.Controls.Add(Me.cmdShowTable)
    Me.bpl.Controls.Add(Me.cmdNew)
    Me.bpl.Controls.Add(Me.cmdSelect)
    Me.bpl.Controls.Add(Me.cmdClose)
    Me.bpl.Controls.Add(Me.cmdExport)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bpl.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bpl.Location = New System.Drawing.Point(0, 288)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(762, 39)
    Me.bpl.TabIndex = 2
    '
    'cmdSubTable1
    '
    Me.cmdSubTable1.Location = New System.Drawing.Point(4, 6)
    Me.cmdSubTable1.Name = "cmdSubTable1"
    Me.cmdSubTable1.Size = New System.Drawing.Size(57, 27)
    Me.cmdSubTable1.TabIndex = 9
    Me.cmdSubTable1.Text = "cmdSubTable1"
    Me.cmdSubTable1.UseVisualStyleBackColor = True
    '
    'cmdSubTable2
    '
    Me.cmdSubTable2.Location = New System.Drawing.Point(62, 6)
    Me.cmdSubTable2.Name = "cmdSubTable2"
    Me.cmdSubTable2.Size = New System.Drawing.Size(57, 27)
    Me.cmdSubTable2.TabIndex = 8
    Me.cmdSubTable2.Text = "cmdSubTable2"
    Me.cmdSubTable2.UseVisualStyleBackColor = True
    '
    'cmdSubTable3
    '
    Me.cmdSubTable3.Location = New System.Drawing.Point(120, 6)
    Me.cmdSubTable3.Name = "cmdSubTable3"
    Me.cmdSubTable3.Size = New System.Drawing.Size(57, 27)
    Me.cmdSubTable3.TabIndex = 10
    Me.cmdSubTable3.Text = "cmdSubTable3"
    Me.cmdSubTable3.UseVisualStyleBackColor = True
    '
    'cmdSubTable4
    '
    Me.cmdSubTable4.Location = New System.Drawing.Point(178, 6)
    Me.cmdSubTable4.Name = "cmdSubTable4"
    Me.cmdSubTable4.Size = New System.Drawing.Size(57, 27)
    Me.cmdSubTable4.TabIndex = 11
    Me.cmdSubTable4.Text = "cmdSubTable4"
    Me.cmdSubTable4.UseVisualStyleBackColor = True
    '
    'cmdSave
    '
    Me.cmdSave.Enabled = False
    Me.cmdSave.Location = New System.Drawing.Point(236, 6)
    Me.cmdSave.Name = "cmdSave"
    Me.cmdSave.Size = New System.Drawing.Size(57, 27)
    Me.cmdSave.TabIndex = 4
    Me.cmdSave.Text = "Sa&ve"
    Me.cmdSave.UseVisualStyleBackColor = True
    '
    'cmdDelete
    '
    Me.cmdDelete.Location = New System.Drawing.Point(294, 6)
    Me.cmdDelete.Name = "cmdDelete"
    Me.cmdDelete.Size = New System.Drawing.Size(57, 27)
    Me.cmdDelete.TabIndex = 6
    Me.cmdDelete.Text = "&Delete"
    Me.cmdDelete.UseVisualStyleBackColor = True
    '
    'cmdAmend
    '
    Me.cmdAmend.Location = New System.Drawing.Point(352, 6)
    Me.cmdAmend.Name = "cmdAmend"
    Me.cmdAmend.Size = New System.Drawing.Size(57, 27)
    Me.cmdAmend.TabIndex = 7
    Me.cmdAmend.Text = "&Amend"
    Me.cmdAmend.UseVisualStyleBackColor = True
    '
    'cmdMore
    '
    Me.cmdMore.Location = New System.Drawing.Point(410, 6)
    Me.cmdMore.Name = "cmdMore"
    Me.cmdMore.Size = New System.Drawing.Size(57, 27)
    Me.cmdMore.TabIndex = 12
    Me.cmdMore.Text = "&More"
    Me.cmdMore.UseVisualStyleBackColor = True
    Me.cmdMore.Visible = False
    '
    'cmdShowTable
    '
    Me.cmdShowTable.Location = New System.Drawing.Point(468, 6)
    Me.cmdShowTable.Name = "cmdShowTable"
    Me.cmdShowTable.Size = New System.Drawing.Size(57, 27)
    Me.cmdShowTable.TabIndex = 3
    Me.cmdShowTable.Text = "Show &Table"
    Me.cmdShowTable.UseVisualStyleBackColor = True
    '
    'cmdNew
    '
    Me.cmdNew.Location = New System.Drawing.Point(526, 6)
    Me.cmdNew.Name = "cmdNew"
    Me.cmdNew.Size = New System.Drawing.Size(57, 27)
    Me.cmdNew.TabIndex = 2
    Me.cmdNew.Text = "&New"
    Me.cmdNew.UseVisualStyleBackColor = True
    '
    'cmdSelect
    '
    Me.cmdSelect.Location = New System.Drawing.Point(584, 6)
    Me.cmdSelect.Name = "cmdSelect"
    Me.cmdSelect.Size = New System.Drawing.Size(57, 27)
    Me.cmdSelect.TabIndex = 1
    Me.cmdSelect.Text = "&Select"
    Me.cmdSelect.UseVisualStyleBackColor = True
    '
    'cmdClose
    '
    Me.cmdClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdClose.Location = New System.Drawing.Point(642, 6)
    Me.cmdClose.Name = "cmdClose"
    Me.cmdClose.Size = New System.Drawing.Size(57, 27)
    Me.cmdClose.TabIndex = 0
    Me.cmdClose.Text = "Close"
    Me.cmdClose.UseVisualStyleBackColor = True
    '
    'cmdExport
    '
    Me.cmdExport.Location = New System.Drawing.Point(700, 6)
    Me.cmdExport.Name = "cmdExport"
    Me.cmdExport.Size = New System.Drawing.Size(57, 27)
    Me.cmdExport.TabIndex = 5
    Me.cmdExport.Text = "Export"
    Me.cmdExport.UseVisualStyleBackColor = True
    '
    'pnl
    '
    Me.pnl.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
    Me.pnl.BackColor = System.Drawing.Color.Transparent
    Me.pnl.Controls.Add(Me.lblGroups)
    Me.pnl.Controls.Add(Me.cboGroups)
    Me.pnl.Controls.Add(Me.dgr)
    Me.pnl.Controls.Add(Me.tlp)
    Me.pnl.Controls.Add(Me.lblContents)
    Me.pnl.Controls.Add(Me.lblTableDesc)
    Me.pnl.Dock = System.Windows.Forms.DockStyle.Fill
    Me.pnl.Location = New System.Drawing.Point(0, 0)
    Me.pnl.Name = "pnl"
    Me.pnl.Size = New System.Drawing.Size(762, 327)
    Me.pnl.TabIndex = 3
    '
    'lblGroups
    '
    Me.lblGroups.AutoSize = True
    Me.lblGroups.Location = New System.Drawing.Point(412, 5)
    Me.lblGroups.Name = "lblGroups"
    Me.lblGroups.Size = New System.Drawing.Size(39, 13)
    Me.lblGroups.TabIndex = 20
    Me.lblGroups.Text = "Group:"
    '
    'cboGroups
    '
    Me.cboGroups.FormattingEnabled = True
    Me.cboGroups.Location = New System.Drawing.Point(470, 5)
    Me.cboGroups.Name = "cboGroups"
    Me.cboGroups.Size = New System.Drawing.Size(202, 21)
    Me.cboGroups.TabIndex = 4
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
    Me.dgr.Location = New System.Drawing.Point(0, 59)
    Me.dgr.MaintenanceDesc = Nothing
    Me.dgr.MaxGridRows = 6
    Me.dgr.MultipleSelect = False
    Me.dgr.Name = "dgr"
    Me.dgr.Padding = New System.Windows.Forms.Padding(0, 0, 0, 50)
    Me.dgr.RowCount = 10
    Me.dgr.ShowIfEmpty = False
    Me.dgr.Size = New System.Drawing.Size(762, 268)
    Me.dgr.SuppressHyperLinkFormat = False
    Me.dgr.TabIndex = 19
    Me.dgr.Visible = False
    '
    'tlp
    '
    Me.tlp.AutoSize = True
    Me.tlp.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
    Me.tlp.ColumnCount = 2
    Me.tlp.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
    Me.tlp.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
    Me.tlp.Controls.Add(Me.lblTableNotes, 0, 0)
    Me.tlp.Controls.Add(Me.txtTableNotes, 1, 0)
    Me.tlp.Controls.Add(Me.txtDefaultValues, 1, 1)
    Me.tlp.Controls.Add(Me.txtAdminNotes, 1, 2)
    Me.tlp.Controls.Add(Me.lblAdminNotes, 0, 2)
    Me.tlp.Controls.Add(Me.lblDefaultValues, 0, 1)
    Me.tlp.Dock = System.Windows.Forms.DockStyle.Fill
    Me.tlp.Location = New System.Drawing.Point(0, 59)
    Me.tlp.Name = "tlp"
    Me.tlp.Padding = New System.Windows.Forms.Padding(0, 0, 0, 50)
    Me.tlp.RowCount = 3
    Me.tlp.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33333!))
    Me.tlp.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33333!))
    Me.tlp.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33333!))
    Me.tlp.Size = New System.Drawing.Size(762, 268)
    Me.tlp.TabIndex = 18
    '
    'lblTableNotes
    '
    Me.lblTableNotes.Location = New System.Drawing.Point(3, 0)
    Me.lblTableNotes.Name = "lblTableNotes"
    Me.lblTableNotes.Size = New System.Drawing.Size(129, 26)
    Me.lblTableNotes.TabIndex = 8
    Me.lblTableNotes.Text = "Table Notes:"
    Me.lblTableNotes.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    '
    'txtTableNotes
    '
    Me.txtTableNotes.BackColor = System.Drawing.Color.White
    Me.txtTableNotes.Dock = System.Windows.Forms.DockStyle.Fill
    Me.txtTableNotes.Location = New System.Drawing.Point(138, 3)
    Me.txtTableNotes.Multiline = True
    Me.txtTableNotes.Name = "txtTableNotes"
    Me.txtTableNotes.ReadOnly = True
    Me.txtTableNotes.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
    Me.txtTableNotes.Size = New System.Drawing.Size(621, 66)
    Me.txtTableNotes.TabIndex = 10
    '
    'txtDefaultValues
    '
    Me.txtDefaultValues.BackColor = System.Drawing.Color.White
    Me.txtDefaultValues.Dock = System.Windows.Forms.DockStyle.Fill
    Me.txtDefaultValues.Location = New System.Drawing.Point(138, 75)
    Me.txtDefaultValues.Multiline = True
    Me.txtDefaultValues.Name = "txtDefaultValues"
    Me.txtDefaultValues.ReadOnly = True
    Me.txtDefaultValues.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
    Me.txtDefaultValues.Size = New System.Drawing.Size(621, 66)
    Me.txtDefaultValues.TabIndex = 11
    '
    'txtAdminNotes
    '
    Me.txtAdminNotes.Dock = System.Windows.Forms.DockStyle.Fill
    Me.txtAdminNotes.Location = New System.Drawing.Point(138, 147)
    Me.txtAdminNotes.Multiline = True
    Me.txtAdminNotes.Name = "txtAdminNotes"
    Me.txtAdminNotes.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
    Me.txtAdminNotes.Size = New System.Drawing.Size(621, 68)
    Me.txtAdminNotes.TabIndex = 13
    '
    'lblAdminNotes
    '
    Me.lblAdminNotes.Location = New System.Drawing.Point(3, 144)
    Me.lblAdminNotes.Name = "lblAdminNotes"
    Me.lblAdminNotes.Size = New System.Drawing.Size(113, 26)
    Me.lblAdminNotes.TabIndex = 12
    Me.lblAdminNotes.Text = "Administrator Notes:"
    Me.lblAdminNotes.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    '
    'lblDefaultValues
    '
    Me.lblDefaultValues.Location = New System.Drawing.Point(3, 72)
    Me.lblDefaultValues.Name = "lblDefaultValues"
    Me.lblDefaultValues.Size = New System.Drawing.Size(113, 26)
    Me.lblDefaultValues.TabIndex = 9
    Me.lblDefaultValues.Text = "Default Values:"
    Me.lblDefaultValues.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    '
    'lblContents
    '
    Me.lblContents.AutoSize = True
    Me.lblContents.Dock = System.Windows.Forms.DockStyle.Top
    Me.lblContents.Location = New System.Drawing.Point(0, 26)
    Me.lblContents.Name = "lblContents"
    Me.lblContents.Padding = New System.Windows.Forms.Padding(0, 10, 0, 10)
    Me.lblContents.Size = New System.Drawing.Size(0, 33)
    Me.lblContents.TabIndex = 17
    Me.lblContents.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    '
    'lblTableDesc
    '
    Me.lblTableDesc.Dock = System.Windows.Forms.DockStyle.Top
    Me.lblTableDesc.Location = New System.Drawing.Point(0, 0)
    Me.lblTableDesc.Name = "lblTableDesc"
    Me.lblTableDesc.Size = New System.Drawing.Size(762, 26)
    Me.lblTableDesc.TabIndex = 6
    Me.lblTableDesc.Text = "Table Description:"
    Me.lblTableDesc.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    '
    'frmTableMaintenance
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.cmdClose
    Me.ClientSize = New System.Drawing.Size(762, 327)
    Me.Controls.Add(Me.bpl)
    Me.Controls.Add(Me.cboTables)
    Me.Controls.Add(Me.pnl)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmTableMaintenance"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.Text = "frmTableMaintenance"
    Me.bpl.ResumeLayout(False)
    Me.pnl.ResumeLayout(False)
    Me.pnl.PerformLayout()
    Me.tlp.ResumeLayout(False)
    Me.tlp.PerformLayout()
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents cboTables As System.Windows.Forms.ComboBox
  Friend WithEvents bpl As CDBNETCL.ButtonPanel
  Friend WithEvents cmdSave As System.Windows.Forms.Button
  Friend WithEvents cmdShowTable As System.Windows.Forms.Button
  Friend WithEvents cmdNew As System.Windows.Forms.Button
  Friend WithEvents cmdSelect As System.Windows.Forms.Button
  Friend WithEvents cmdClose As System.Windows.Forms.Button
  Friend WithEvents pnl As CDBNETCL.PanelEx
  Friend WithEvents lblTableDesc As System.Windows.Forms.Label
  Friend WithEvents tlp As System.Windows.Forms.TableLayoutPanel
  Friend WithEvents lblTableNotes As System.Windows.Forms.Label
  Friend WithEvents txtTableNotes As System.Windows.Forms.TextBox
  Friend WithEvents txtDefaultValues As System.Windows.Forms.TextBox
  Friend WithEvents txtAdminNotes As System.Windows.Forms.TextBox
  Friend WithEvents lblAdminNotes As System.Windows.Forms.Label
  Friend WithEvents lblDefaultValues As System.Windows.Forms.Label
  Friend WithEvents lblContents As System.Windows.Forms.Label
  Friend WithEvents dgr As CDBNETCL.DisplayGrid
  Friend WithEvents cmdExport As System.Windows.Forms.Button
  Friend WithEvents cmdDelete As System.Windows.Forms.Button
  Friend WithEvents cmdAmend As System.Windows.Forms.Button
  Friend WithEvents cmdSubTable1 As System.Windows.Forms.Button
  Friend WithEvents cmdSubTable2 As System.Windows.Forms.Button
  Friend WithEvents cmdSubTable4 As System.Windows.Forms.Button
  Friend WithEvents cmdSubTable3 As System.Windows.Forms.Button
  Friend WithEvents lblGroups As System.Windows.Forms.Label
  Friend WithEvents cboGroups As System.Windows.Forms.ComboBox
  Friend WithEvents cmdMore As System.Windows.Forms.Button
End Class
