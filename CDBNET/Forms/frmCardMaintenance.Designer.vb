<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCardMaintenance
  Inherits MaintenanceParentForm

  'Form overrides dispose to clean up the component list.
  <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
    If disposing Then
      If Not (components Is Nothing) Then
        components.Dispose()
      End If
    End If
    MyBase.Dispose(disposing)
  End Sub

  Friend WithEvents splTop As System.Windows.Forms.SplitContainer
  Friend WithEvents splBottom As System.Windows.Forms.SplitContainer
  Friend WithEvents splRight As System.Windows.Forms.SplitContainer
  Friend WithEvents dgr As CDBNETCL.DisplayGrid
  Friend WithEvents epl As CDBNETCL.EditPanel
  Friend WithEvents bpl As CDBNETCL.ButtonPanel
  Friend WithEvents cmdCreateOrEdit As System.Windows.Forms.Button
  Friend WithEvents cmdLink1 As System.Windows.Forms.Button
  Friend WithEvents cmdLink2 As System.Windows.Forms.Button
  Friend WithEvents cmdDefault As System.Windows.Forms.Button
  Friend WithEvents cmdSave As System.Windows.Forms.Button
  Friend WithEvents cmdNew As System.Windows.Forms.Button
  Friend WithEvents cmdDelete As System.Windows.Forms.Button
  Friend WithEvents cmdClose As System.Windows.Forms.Button

  'Required by the Windows Form Designer
  Private components As System.ComponentModel.IContainer

  'NOTE: The following procedure is required by the Windows Form Designer
  'It can be modified using the Windows Form Designer.  
  'Do not modify it using the code editor.
  <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
    Me.components = New System.ComponentModel.Container
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCardMaintenance))
    Me.splTop = New System.Windows.Forms.SplitContainer
    Me.splBottom = New System.Windows.Forms.SplitContainer
    Me.sel = New CDBNETCL.TabSelector
    Me.splMaint = New System.Windows.Forms.SplitContainer
    Me.cboData = New System.Windows.Forms.ComboBox
    Me.splRight = New System.Windows.Forms.SplitContainer
    Me.dgr = New CDBNETCL.DisplayGrid
    Me.epl = New CDBNETCL.EditPanel
    Me.bpl = New CDBNETCL.ButtonPanel
    Me.cmdCreateOrEdit = New System.Windows.Forms.Button
    Me.cmdLink1 = New System.Windows.Forms.Button
    Me.cmdLink2 = New System.Windows.Forms.Button
    Me.cmdOther = New System.Windows.Forms.Button
    Me.cmdReply = New System.Windows.Forms.Button
    Me.cmdDefault = New System.Windows.Forms.Button
    Me.cmdSave = New System.Windows.Forms.Button
    Me.cmdNew = New System.Windows.Forms.Button
    Me.cmdDelete = New System.Windows.Forms.Button
    Me.cmdClose = New System.Windows.Forms.Button
    Me.tsImageList = New System.Windows.Forms.ImageList(Me.components)
    Me.cmdPrint = New System.Windows.Forms.Button
    Me.splTop.Panel2.SuspendLayout()
    Me.splTop.SuspendLayout()
    Me.splBottom.Panel1.SuspendLayout()
    Me.splBottom.Panel2.SuspendLayout()
    Me.splBottom.SuspendLayout()
    Me.splMaint.Panel1.SuspendLayout()
    Me.splMaint.Panel2.SuspendLayout()
    Me.splMaint.SuspendLayout()
    Me.splRight.Panel1.SuspendLayout()
    Me.splRight.Panel2.SuspendLayout()
    Me.splRight.SuspendLayout()
    Me.bpl.SuspendLayout()
    Me.SuspendLayout()
    '
    'splTop
    '
    Me.splTop.Dock = System.Windows.Forms.DockStyle.Fill
    Me.splTop.Location = New System.Drawing.Point(0, 0)
    Me.splTop.Name = "splTop"
    Me.splTop.Orientation = System.Windows.Forms.Orientation.Horizontal
    '
    'splTop.Panel2
    '
    Me.splTop.Panel2.Controls.Add(Me.splBottom)
    Me.splTop.Size = New System.Drawing.Size(834, 559)
    Me.splTop.SplitterDistance = 45
    Me.splTop.SplitterWidth = 8
    Me.splTop.TabIndex = 0
    '
    'splBottom
    '
    Me.splBottom.Dock = System.Windows.Forms.DockStyle.Fill
    Me.splBottom.Location = New System.Drawing.Point(0, 0)
    Me.splBottom.Name = "splBottom"
    '
    'splBottom.Panel1
    '
    Me.splBottom.Panel1.Controls.Add(Me.sel)
    '
    'splBottom.Panel2
    '
    Me.splBottom.Panel2.Controls.Add(Me.splMaint)
    Me.splBottom.Panel2.Controls.Add(Me.bpl)
    Me.splBottom.Size = New System.Drawing.Size(834, 506)
    Me.splBottom.SplitterDistance = 200
    Me.splBottom.SplitterWidth = 8
    Me.splBottom.TabIndex = 0
    '
    'sel
    '
    Me.sel.BackColor = System.Drawing.Color.Transparent
    Me.sel.Dock = System.Windows.Forms.DockStyle.Fill
    Me.sel.Location = New System.Drawing.Point(0, 0)
    Me.sel.Name = "sel"
    Me.sel.Padding = New System.Windows.Forms.Padding(2, 18, 2, 18)
    Me.sel.Size = New System.Drawing.Size(200, 506)
    Me.sel.TabIndex = 0
    Me.sel.TreeContextMenu = Nothing
    '
    'splMaint
    '
    Me.splMaint.Dock = System.Windows.Forms.DockStyle.Fill
    Me.splMaint.IsSplitterFixed = True
    Me.splMaint.Location = New System.Drawing.Point(0, 0)
    Me.splMaint.Name = "splMaint"
    Me.splMaint.Orientation = System.Windows.Forms.Orientation.Horizontal
    '
    'splMaint.Panel1
    '
    Me.splMaint.Panel1.Controls.Add(Me.cboData)
    Me.splMaint.Panel1.Padding = New System.Windows.Forms.Padding(4, 4, 4, 0)
    Me.splMaint.Panel1MinSize = 29
    '
    'splMaint.Panel2
    '
    Me.splMaint.Panel2.Controls.Add(Me.splRight)
    Me.splMaint.Size = New System.Drawing.Size(626, 467)
    Me.splMaint.SplitterDistance = 29
    Me.splMaint.SplitterWidth = 1
    Me.splMaint.TabIndex = 1
    '
    'cboData
    '
    Me.cboData.Dock = System.Windows.Forms.DockStyle.Fill
    Me.cboData.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cboData.FormattingEnabled = True
    Me.cboData.Location = New System.Drawing.Point(4, 4)
    Me.cboData.Name = "cboData"
    Me.cboData.Size = New System.Drawing.Size(618, 24)
    Me.cboData.TabIndex = 1
    '
    'splRight
    '
    Me.splRight.Dock = System.Windows.Forms.DockStyle.Fill
    Me.splRight.Location = New System.Drawing.Point(0, 0)
    Me.splRight.Name = "splRight"
    Me.splRight.Orientation = System.Windows.Forms.Orientation.Horizontal
    '
    'splRight.Panel1
    '
    Me.splRight.Panel1.Controls.Add(Me.dgr)
    '
    'splRight.Panel2
    '
    Me.splRight.Panel2.Controls.Add(Me.epl)
    Me.splRight.Size = New System.Drawing.Size(626, 437)
    Me.splRight.SplitterDistance = 172
    Me.splRight.SplitterWidth = 8
    Me.splRight.TabIndex = 0
    '
    'dgr
    '
    Me.dgr.AutoSetHeight = False
    Me.dgr.AutoSetRowHeight = False
    Me.dgr.DisplayTitle = Nothing
    Me.dgr.Dock = System.Windows.Forms.DockStyle.Fill
    Me.dgr.Location = New System.Drawing.Point(0, 0)
    Me.dgr.MaxGridRows = 8
    Me.dgr.MultipleSelect = True
    Me.dgr.Name = "dgr"
    Me.dgr.Padding = New System.Windows.Forms.Padding(4)
    Me.dgr.ShowIfEmpty = False
    Me.dgr.Size = New System.Drawing.Size(626, 172)
    Me.dgr.TabIndex = 0
    '
    'epl
    '
    Me.epl.AutoScroll = True
    Me.epl.BackColor = System.Drawing.Color.Transparent
    Me.epl.DataChanged = False
    Me.epl.Dock = System.Windows.Forms.DockStyle.Fill
    Me.epl.Location = New System.Drawing.Point(0, 0)
    Me.epl.Name = "epl"
    Me.epl.Recipients = Nothing
    Me.epl.Size = New System.Drawing.Size(626, 257)
    Me.epl.TabIndex = 0
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdCreateOrEdit)
    Me.bpl.Controls.Add(Me.cmdLink1)
    Me.bpl.Controls.Add(Me.cmdLink2)
    Me.bpl.Controls.Add(Me.cmdOther)
    Me.bpl.Controls.Add(Me.cmdReply)
    Me.bpl.Controls.Add(Me.cmdPrint)
    Me.bpl.Controls.Add(Me.cmdDefault)
    Me.bpl.Controls.Add(Me.cmdSave)
    Me.bpl.Controls.Add(Me.cmdNew)
    Me.bpl.Controls.Add(Me.cmdDelete)
    Me.bpl.Controls.Add(Me.cmdClose)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bpl.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bpl.Location = New System.Drawing.Point(0, 467)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(626, 39)
    Me.bpl.TabIndex = 1
    '
    'cmdCreateOrEdit
    '
    Me.cmdCreateOrEdit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdCreateOrEdit.Location = New System.Drawing.Point(0, 6)
    Me.cmdCreateOrEdit.Name = "cmdCreateOrEdit"
    Me.cmdCreateOrEdit.Size = New System.Drawing.Size(56, 27)
    Me.cmdCreateOrEdit.TabIndex = 0
    Me.cmdCreateOrEdit.Text = "Create"
    Me.cmdCreateOrEdit.Visible = False
    '
    'cmdLink1
    '
    Me.cmdLink1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdLink1.Location = New System.Drawing.Point(57, 6)
    Me.cmdLink1.Name = "cmdLink1"
    Me.cmdLink1.Size = New System.Drawing.Size(56, 27)
    Me.cmdLink1.TabIndex = 1
    Me.cmdLink1.Visible = False
    '
    'cmdLink2
    '
    Me.cmdLink2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdLink2.Location = New System.Drawing.Point(114, 6)
    Me.cmdLink2.Name = "cmdLink2"
    Me.cmdLink2.Size = New System.Drawing.Size(56, 27)
    Me.cmdLink2.TabIndex = 2
    Me.cmdLink2.Visible = False
    '
    'cmdOther
    '
    Me.cmdOther.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdOther.Location = New System.Drawing.Point(171, 6)
    Me.cmdOther.Name = "cmdOther"
    Me.cmdOther.Size = New System.Drawing.Size(56, 27)
    Me.cmdOther.TabIndex = 3
    Me.cmdOther.Visible = False
    '
    'cmdReply
    '
    Me.cmdReply.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdReply.Location = New System.Drawing.Point(228, 6)
    Me.cmdReply.Name = "cmdReply"
    Me.cmdReply.Size = New System.Drawing.Size(56, 27)
    Me.cmdReply.TabIndex = 9
    Me.cmdReply.Text = "Re&ply"
    Me.cmdReply.Visible = False
    '
    'cmdDefault
    '
    Me.cmdDefault.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdDefault.Location = New System.Drawing.Point(342, 6)
    Me.cmdDefault.Name = "cmdDefault"
    Me.cmdDefault.Size = New System.Drawing.Size(56, 27)
    Me.cmdDefault.TabIndex = 4
    Me.cmdDefault.Text = "&Default"
    '
    'cmdSave
    '
    Me.cmdSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdSave.Location = New System.Drawing.Point(399, 6)
    Me.cmdSave.Name = "cmdSave"
    Me.cmdSave.Size = New System.Drawing.Size(56, 27)
    Me.cmdSave.TabIndex = 5
    Me.cmdSave.Text = "&Save"
    '
    'cmdNew
    '
    Me.cmdNew.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdNew.Location = New System.Drawing.Point(456, 6)
    Me.cmdNew.Name = "cmdNew"
    Me.cmdNew.Size = New System.Drawing.Size(56, 27)
    Me.cmdNew.TabIndex = 6
    Me.cmdNew.Text = "&New"
    '
    'cmdDelete
    '
    Me.cmdDelete.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdDelete.Location = New System.Drawing.Point(513, 6)
    Me.cmdDelete.Name = "cmdDelete"
    Me.cmdDelete.Size = New System.Drawing.Size(56, 27)
    Me.cmdDelete.TabIndex = 7
    Me.cmdDelete.Text = "De&lete"
    '
    'cmdClose
    '
    Me.cmdClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdClose.Location = New System.Drawing.Point(570, 6)
    Me.cmdClose.Name = "cmdClose"
    Me.cmdClose.Size = New System.Drawing.Size(56, 27)
    Me.cmdClose.TabIndex = 8
    Me.cmdClose.Text = "Close"
    '
    'tsImageList
    '
    Me.tsImageList.ImageStream = CType(resources.GetObject("tsImageList.ImageStream"), System.Windows.Forms.ImageListStreamer)
    Me.tsImageList.TransparentColor = System.Drawing.Color.Transparent
    Me.tsImageList.Images.SetKeyName(0, "citgen")
    Me.tsImageList.Images.SetKeyName(1, "citCampaign")
    Me.tsImageList.Images.SetKeyName(2, "atH2HCollection")
    Me.tsImageList.Images.SetKeyName(3, "atMannedCollection")
    Me.tsImageList.Images.SetKeyName(4, "atSegment")
    Me.tsImageList.Images.SetKeyName(5, "atUnMannedCollection")
    '
    'cmdPrint
    '
    Me.cmdPrint.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdPrint.Location = New System.Drawing.Point(285, 6)
    Me.cmdPrint.Name = "cmdPrint"
    Me.cmdPrint.Size = New System.Drawing.Size(56, 27)
    Me.cmdPrint.TabIndex = 10
    Me.cmdPrint.Text = "&Print"
    Me.cmdPrint.Visible = False
    '
    'frmCardMaintenance
    '
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
    Me.ClientSize = New System.Drawing.Size(834, 559)
    Me.Controls.Add(Me.splTop)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmCardMaintenance"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.Text = "Record Maintenance"
    Me.TopMost = True
    Me.splTop.Panel2.ResumeLayout(False)
    Me.splTop.ResumeLayout(False)
    Me.splBottom.Panel1.ResumeLayout(False)
    Me.splBottom.Panel2.ResumeLayout(False)
    Me.splBottom.ResumeLayout(False)
    Me.splMaint.Panel1.ResumeLayout(False)
    Me.splMaint.Panel2.ResumeLayout(False)
    Me.splMaint.ResumeLayout(False)
    Me.splRight.Panel1.ResumeLayout(False)
    Me.splRight.Panel2.ResumeLayout(False)
    Me.splRight.ResumeLayout(False)
    Me.bpl.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents sel As CDBNETCL.TabSelector
  Friend WithEvents cmdOther As System.Windows.Forms.Button
  Friend WithEvents cmdReply As System.Windows.Forms.Button
  Friend WithEvents cboData As System.Windows.Forms.ComboBox
  Friend WithEvents splMaint As System.Windows.Forms.SplitContainer
  Friend WithEvents tsImageList As System.Windows.Forms.ImageList
  Friend WithEvents cmdPrint As System.Windows.Forms.Button

End Class