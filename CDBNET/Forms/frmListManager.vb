Public Class frmListManager
  Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

  Public Sub New()
    MyBase.New()

    'This call is required by the Windows Form Designer.
    InitializeComponent()

    'Add any initialization after the InitializeComponent() call
    InitialiseControls()
  End Sub

  'Form overrides dispose to clean up the component list.
  Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
    If disposing Then
      If Not (components Is Nothing) Then
        components.Dispose()
      End If
    End If
    MyBase.Dispose(disposing)
  End Sub

  'Required by the Windows Form Designer
  Private components As System.ComponentModel.IContainer

  'NOTE: The following procedure is required by the Windows Form Designer
  'It can be modified using the Windows Form Designer.  
  'Do not modify it using the code editor.
  Friend WithEvents sbr As System.Windows.Forms.StatusBar
  Friend WithEvents sbpRows As System.Windows.Forms.StatusBarPanel
  Friend WithEvents sbpSelected As System.Windows.Forms.StatusBarPanel
  Friend WithEvents sbpViewing As System.Windows.Forms.StatusBarPanel
  Friend WithEvents sbpSteps As System.Windows.Forms.StatusBarPanel
  Friend WithEvents sbrStatus As System.Windows.Forms.StatusBar
  Friend WithEvents mnuBar As System.Windows.Forms.MainMenu
  Friend WithEvents imgToolbar16 As System.Windows.Forms.ImageList
  Friend WithEvents imgToolbar32 As System.Windows.Forms.ImageList
  Friend WithEvents ttp As System.Windows.Forms.ToolTip
  Friend WithEvents pnlData As System.Windows.Forms.Panel
  Friend WithEvents dgr As CDBNETCL.DisplayGrid
  Friend WithEvents pnlSteps As System.Windows.Forms.Panel
  Friend WithEvents lblSteps As System.Windows.Forms.Label
  Friend WithEvents dgrSteps As CDBNETCL.DisplayGrid
  Friend WithEvents txtFilter As System.Windows.Forms.TextBox
  Friend WithEvents lblData As System.Windows.Forms.Label
  Friend WithEvents imbUp As CDBNETCL.ImageButton
  Friend WithEvents imbDelete As CDBNETCL.ImageButton
  Friend WithEvents imbDown As CDBNETCL.ImageButton
  Friend WithEvents sfd As System.Windows.Forms.SaveFileDialog
  Friend WithEvents tsp As System.Windows.Forms.ToolStrip
  Friend WithEvents mnu As System.Windows.Forms.MenuStrip
  Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents ViewToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents DataToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents ToolStripSeparator7 As System.Windows.Forms.ToolStripSeparator
  Friend WithEvents splTop As System.Windows.Forms.SplitContainer
  Friend WithEvents splBottom As System.Windows.Forms.SplitContainer
  Friend WithEvents ctxMenuStrip As System.Windows.Forms.ContextMenuStrip
  Friend WithEvents ctxsFilterSelection As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents ctxsExcludeColumnSelection As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents ctxsEnterFilter As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents ctxsFilterEmpty As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents ctxsFilterNotEmpty As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents ctxsClearColumnFilter As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents ctxsClearFilterOnAllColumns As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents pnlStepCommands As System.Windows.Forms.Panel
  <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
    Me.components = New System.ComponentModel.Container
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmListManager))
    Me.imgToolbar32 = New System.Windows.Forms.ImageList(Me.components)
    Me.imgToolbar16 = New System.Windows.Forms.ImageList(Me.components)
    Me.mnuBar = New System.Windows.Forms.MainMenu(Me.components)
    Me.sbr = New System.Windows.Forms.StatusBar
    Me.sbpRows = New System.Windows.Forms.StatusBarPanel
    Me.sbpSelected = New System.Windows.Forms.StatusBarPanel
    Me.sbpViewing = New System.Windows.Forms.StatusBarPanel
    Me.sbpSteps = New System.Windows.Forms.StatusBarPanel
    Me.sbrStatus = New System.Windows.Forms.StatusBar
    Me.ttp = New System.Windows.Forms.ToolTip(Me.components)
    Me.imbDelete = New CDBNETCL.ImageButton
    Me.imbDown = New CDBNETCL.ImageButton
    Me.imbUp = New CDBNETCL.ImageButton
    Me.pnlData = New System.Windows.Forms.Panel
    Me.dgr = New CDBNETCL.DisplayGrid
    Me.lblData = New System.Windows.Forms.Label
    Me.pnlSteps = New System.Windows.Forms.Panel
    Me.pnlStepCommands = New System.Windows.Forms.Panel
    Me.dgrSteps = New CDBNETCL.DisplayGrid
    Me.lblSteps = New System.Windows.Forms.Label
    Me.txtFilter = New System.Windows.Forms.TextBox
    Me.sfd = New System.Windows.Forms.SaveFileDialog
    Me.tsp = New System.Windows.Forms.ToolStrip
    Me.mnu = New System.Windows.Forms.MenuStrip
    Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
    Me.ViewToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
    Me.DataToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
    Me.ToolStripSeparator7 = New System.Windows.Forms.ToolStripSeparator
    Me.splTop = New System.Windows.Forms.SplitContainer
    Me.splBottom = New System.Windows.Forms.SplitContainer
    Me.ctxMenuStrip = New System.Windows.Forms.ContextMenuStrip(Me.components)
    Me.ctxsFilterSelection = New System.Windows.Forms.ToolStripMenuItem
    Me.ctxsExcludeColumnSelection = New System.Windows.Forms.ToolStripMenuItem
    Me.ctxsEnterFilter = New System.Windows.Forms.ToolStripMenuItem
    Me.ctxsFilterEmpty = New System.Windows.Forms.ToolStripMenuItem
    Me.ctxsFilterNotEmpty = New System.Windows.Forms.ToolStripMenuItem
    Me.ctxsClearColumnFilter = New System.Windows.Forms.ToolStripMenuItem
    Me.ctxsClearFilterOnAllColumns = New System.Windows.Forms.ToolStripMenuItem
    CType(Me.sbpRows, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.sbpSelected, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.sbpViewing, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.sbpSteps, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.pnlData.SuspendLayout()
    Me.pnlSteps.SuspendLayout()
    Me.pnlStepCommands.SuspendLayout()
    Me.mnu.SuspendLayout()
    Me.splTop.Panel1.SuspendLayout()
    Me.splTop.Panel2.SuspendLayout()
    Me.splTop.SuspendLayout()
    Me.splBottom.Panel1.SuspendLayout()
    Me.splBottom.Panel2.SuspendLayout()
    Me.splBottom.SuspendLayout()
    Me.ctxMenuStrip.SuspendLayout()
    Me.SuspendLayout()
    '
    'imgToolbar32
    '
    Me.imgToolbar32.ImageStream = CType(resources.GetObject("imgToolbar32.ImageStream"), System.Windows.Forms.ImageListStreamer)
    Me.imgToolbar32.TransparentColor = System.Drawing.Color.Transparent
    Me.imgToolbar32.Images.SetKeyName(0, "")
    Me.imgToolbar32.Images.SetKeyName(1, "")
    Me.imgToolbar32.Images.SetKeyName(2, "")
    Me.imgToolbar32.Images.SetKeyName(3, "")
    Me.imgToolbar32.Images.SetKeyName(4, "")
    Me.imgToolbar32.Images.SetKeyName(5, "")
    Me.imgToolbar32.Images.SetKeyName(6, "")
    Me.imgToolbar32.Images.SetKeyName(7, "")
    Me.imgToolbar32.Images.SetKeyName(8, "")
    Me.imgToolbar32.Images.SetKeyName(9, "")
    Me.imgToolbar32.Images.SetKeyName(10, "")
    Me.imgToolbar32.Images.SetKeyName(11, "")
    Me.imgToolbar32.Images.SetKeyName(12, "")
    Me.imgToolbar32.Images.SetKeyName(13, "")
    Me.imgToolbar32.Images.SetKeyName(14, "")
    Me.imgToolbar32.Images.SetKeyName(15, "")
    Me.imgToolbar32.Images.SetKeyName(16, "")
    '
    'imgToolbar16
    '
    Me.imgToolbar16.ImageStream = CType(resources.GetObject("imgToolbar16.ImageStream"), System.Windows.Forms.ImageListStreamer)
    Me.imgToolbar16.TransparentColor = System.Drawing.Color.Transparent
    Me.imgToolbar16.Images.SetKeyName(0, "")
    Me.imgToolbar16.Images.SetKeyName(1, "")
    Me.imgToolbar16.Images.SetKeyName(2, "")
    Me.imgToolbar16.Images.SetKeyName(3, "")
    Me.imgToolbar16.Images.SetKeyName(4, "")
    Me.imgToolbar16.Images.SetKeyName(5, "")
    Me.imgToolbar16.Images.SetKeyName(6, "")
    Me.imgToolbar16.Images.SetKeyName(7, "")
    Me.imgToolbar16.Images.SetKeyName(8, "")
    Me.imgToolbar16.Images.SetKeyName(9, "")
    Me.imgToolbar16.Images.SetKeyName(10, "")
    Me.imgToolbar16.Images.SetKeyName(11, "")
    Me.imgToolbar16.Images.SetKeyName(12, "")
    Me.imgToolbar16.Images.SetKeyName(13, "")
    Me.imgToolbar16.Images.SetKeyName(14, "")
    Me.imgToolbar16.Images.SetKeyName(15, "")
    Me.imgToolbar16.Images.SetKeyName(16, "")
    '
    'sbr
    '
    Me.sbr.Dock = System.Windows.Forms.DockStyle.Top
    Me.sbr.Location = New System.Drawing.Point(0, 75)
    Me.sbr.Name = "sbr"
    Me.sbr.Panels.AddRange(New System.Windows.Forms.StatusBarPanel() {Me.sbpRows, Me.sbpSelected, Me.sbpViewing, Me.sbpSteps})
    Me.sbr.ShowPanels = True
    Me.sbr.Size = New System.Drawing.Size(992, 22)
    Me.sbr.SizingGrip = False
    Me.sbr.TabIndex = 1
    '
    'sbpRows
    '
    Me.sbpRows.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Contents
    Me.sbpRows.Name = "sbpRows"
    Me.sbpRows.Text = "100 Rows"
    Me.sbpRows.Width = 74
    '
    'sbpSelected
    '
    Me.sbpSelected.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Contents
    Me.sbpSelected.Name = "sbpSelected"
    Me.sbpSelected.Text = "Selected Records 0"
    Me.sbpSelected.Width = 131
    '
    'sbpViewing
    '
    Me.sbpViewing.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Spring
    Me.sbpViewing.Name = "sbpViewing"
    Me.sbpViewing.Text = "Viewing"
    Me.sbpViewing.Width = 717
    '
    'sbpSteps
    '
    Me.sbpSteps.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Contents
    Me.sbpSteps.Name = "sbpSteps"
    Me.sbpSteps.Text = "No Steps"
    Me.sbpSteps.Width = 70
    '
    'sbrStatus
    '
    Me.sbrStatus.Dock = System.Windows.Forms.DockStyle.Top
    Me.sbrStatus.Location = New System.Drawing.Point(0, 51)
    Me.sbrStatus.Name = "sbrStatus"
    Me.sbrStatus.Size = New System.Drawing.Size(992, 24)
    Me.sbrStatus.SizingGrip = False
    Me.sbrStatus.TabIndex = 2
    Me.sbrStatus.Text = "No Current Filter"
    '
    'imbDelete
    '
    Me.imbDelete.BackColor = System.Drawing.Color.Transparent
    Me.imbDelete.DialogResult = System.Windows.Forms.DialogResult.OK
    Me.imbDelete.DrawBorder = False
    Me.imbDelete.Location = New System.Drawing.Point(0, 64)
    Me.imbDelete.Name = "imbDelete"
    Me.imbDelete.NormalImage = CType(resources.GetObject("imbDelete.NormalImage"), System.Drawing.Image)
    Me.imbDelete.ShowFocusRect = True
    Me.imbDelete.Size = New System.Drawing.Size(24, 24)
    Me.imbDelete.SizeMode = CDBNETCL.ImageButton.ImageButtonSizeMode.StretchImage
    Me.imbDelete.TabIndex = 2
    Me.imbDelete.TextAlign = System.Drawing.ContentAlignment.BottomCenter
    Me.ttp.SetToolTip(Me.imbDelete, "Delete Step")
    Me.imbDelete.TransparentColor = System.Drawing.Color.Transparent
    '
    'imbDown
    '
    Me.imbDown.BackColor = System.Drawing.Color.Transparent
    Me.imbDown.DialogResult = System.Windows.Forms.DialogResult.OK
    Me.imbDown.DrawBorder = False
    Me.imbDown.Location = New System.Drawing.Point(0, 32)
    Me.imbDown.Name = "imbDown"
    Me.imbDown.NormalImage = CType(resources.GetObject("imbDown.NormalImage"), System.Drawing.Image)
    Me.imbDown.ShowFocusRect = True
    Me.imbDown.Size = New System.Drawing.Size(24, 24)
    Me.imbDown.SizeMode = CDBNETCL.ImageButton.ImageButtonSizeMode.StretchImage
    Me.imbDown.TabIndex = 1
    Me.imbDown.TextAlign = System.Drawing.ContentAlignment.BottomCenter
    Me.ttp.SetToolTip(Me.imbDown, "Move Step Down")
    Me.imbDown.TransparentColor = System.Drawing.Color.Transparent
    '
    'imbUp
    '
    Me.imbUp.BackColor = System.Drawing.Color.Transparent
    Me.imbUp.DialogResult = System.Windows.Forms.DialogResult.OK
    Me.imbUp.DrawBorder = False
    Me.imbUp.Location = New System.Drawing.Point(0, 0)
    Me.imbUp.Name = "imbUp"
    Me.imbUp.NormalImage = CType(resources.GetObject("imbUp.NormalImage"), System.Drawing.Image)
    Me.imbUp.ShowFocusRect = True
    Me.imbUp.Size = New System.Drawing.Size(24, 24)
    Me.imbUp.SizeMode = CDBNETCL.ImageButton.ImageButtonSizeMode.StretchImage
    Me.imbUp.TabIndex = 0
    Me.imbUp.TextAlign = System.Drawing.ContentAlignment.BottomCenter
    Me.ttp.SetToolTip(Me.imbUp, "Move Step Up")
    Me.imbUp.TransparentColor = System.Drawing.Color.Transparent
    '
    'pnlData
    '
    Me.pnlData.Controls.Add(Me.dgr)
    Me.pnlData.Controls.Add(Me.lblData)
    Me.pnlData.Dock = System.Windows.Forms.DockStyle.Fill
    Me.pnlData.Location = New System.Drawing.Point(4, 4)
    Me.pnlData.Name = "pnlData"
    Me.pnlData.Size = New System.Drawing.Size(984, 276)
    Me.pnlData.TabIndex = 9
    '
    'dgr
    '
    Me.dgr.AllowSorting = True
    Me.dgr.AutoSetHeight = False
    Me.dgr.AutoSetRowHeight = False
    Me.dgr.DisplayTitle = Nothing
    Me.dgr.Dock = System.Windows.Forms.DockStyle.Fill
    Me.dgr.Location = New System.Drawing.Point(0, 16)
    Me.dgr.MaxGridRows = 8
    Me.dgr.MultipleSelect = False
    Me.dgr.Name = "dgr"
    Me.dgr.ShowIfEmpty = False
    Me.dgr.Size = New System.Drawing.Size(984, 260)
    Me.dgr.TabIndex = 1
    '
    'lblData
    '
    Me.lblData.Dock = System.Windows.Forms.DockStyle.Top
    Me.lblData.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.lblData.Location = New System.Drawing.Point(0, 0)
    Me.lblData.Name = "lblData"
    Me.lblData.Size = New System.Drawing.Size(984, 16)
    Me.lblData.TabIndex = 0
    Me.lblData.Text = "Processing Please Wait..."
    Me.lblData.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
    Me.lblData.Visible = False
    '
    'pnlSteps
    '
    Me.pnlSteps.Controls.Add(Me.pnlStepCommands)
    Me.pnlSteps.Controls.Add(Me.dgrSteps)
    Me.pnlSteps.Controls.Add(Me.lblSteps)
    Me.pnlSteps.Dock = System.Windows.Forms.DockStyle.Fill
    Me.pnlSteps.Location = New System.Drawing.Point(0, 0)
    Me.pnlSteps.Name = "pnlSteps"
    Me.pnlSteps.Size = New System.Drawing.Size(984, 85)
    Me.pnlSteps.TabIndex = 11
    '
    'pnlStepCommands
    '
    Me.pnlStepCommands.BackColor = System.Drawing.SystemColors.Control
    Me.pnlStepCommands.Controls.Add(Me.imbDelete)
    Me.pnlStepCommands.Controls.Add(Me.imbDown)
    Me.pnlStepCommands.Controls.Add(Me.imbUp)
    Me.pnlStepCommands.Dock = System.Windows.Forms.DockStyle.Right
    Me.pnlStepCommands.Location = New System.Drawing.Point(960, 16)
    Me.pnlStepCommands.Name = "pnlStepCommands"
    Me.pnlStepCommands.Size = New System.Drawing.Size(24, 69)
    Me.pnlStepCommands.TabIndex = 2
    '
    'dgrSteps
    '
    Me.dgrSteps.AllowSorting = True
    Me.dgrSteps.AutoSetHeight = False
    Me.dgrSteps.AutoSetRowHeight = False
    Me.dgrSteps.DisplayTitle = Nothing
    Me.dgrSteps.Dock = System.Windows.Forms.DockStyle.Fill
    Me.dgrSteps.Location = New System.Drawing.Point(0, 16)
    Me.dgrSteps.MaxGridRows = 8
    Me.dgrSteps.MultipleSelect = False
    Me.dgrSteps.Name = "dgrSteps"
    Me.dgrSteps.Padding = New System.Windows.Forms.Padding(0, 0, 24, 0)
    Me.dgrSteps.ShowIfEmpty = False
    Me.dgrSteps.Size = New System.Drawing.Size(984, 69)
    Me.dgrSteps.TabIndex = 1
    '
    'lblSteps
    '
    Me.lblSteps.Dock = System.Windows.Forms.DockStyle.Top
    Me.lblSteps.Location = New System.Drawing.Point(0, 0)
    Me.lblSteps.Name = "lblSteps"
    Me.lblSteps.Size = New System.Drawing.Size(984, 16)
    Me.lblSteps.TabIndex = 0
    Me.lblSteps.Text = "Steps"
    '
    'txtFilter
    '
    Me.txtFilter.Dock = System.Windows.Forms.DockStyle.Fill
    Me.txtFilter.Location = New System.Drawing.Point(0, 0)
    Me.txtFilter.Multiline = True
    Me.txtFilter.Name = "txtFilter"
    Me.txtFilter.Size = New System.Drawing.Size(984, 44)
    Me.txtFilter.TabIndex = 13
    '
    'tsp
    '
    Me.tsp.Location = New System.Drawing.Point(0, 26)
    Me.tsp.Name = "tsp"
    Me.tsp.Size = New System.Drawing.Size(992, 25)
    Me.tsp.TabIndex = 14
    '
    'mnu
    '
    Me.mnu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.ViewToolStripMenuItem, Me.DataToolStripMenuItem})
    Me.mnu.Location = New System.Drawing.Point(0, 0)
    Me.mnu.Name = "mnu"
    Me.mnu.Size = New System.Drawing.Size(992, 26)
    Me.mnu.TabIndex = 15
    '
    'FileToolStripMenuItem
    '
    Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
    Me.FileToolStripMenuItem.Size = New System.Drawing.Size(40, 22)
    Me.FileToolStripMenuItem.Text = "&File"
    '
    'ViewToolStripMenuItem
    '
    Me.ViewToolStripMenuItem.Name = "ViewToolStripMenuItem"
    Me.ViewToolStripMenuItem.Size = New System.Drawing.Size(49, 22)
    Me.ViewToolStripMenuItem.Text = "&View"
    '
    'DataToolStripMenuItem
    '
    Me.DataToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripSeparator7})
    Me.DataToolStripMenuItem.Name = "DataToolStripMenuItem"
    Me.DataToolStripMenuItem.Size = New System.Drawing.Size(51, 22)
    Me.DataToolStripMenuItem.Text = "&Data"
    '
    'ToolStripSeparator7
    '
    Me.ToolStripSeparator7.Name = "ToolStripSeparator7"
    Me.ToolStripSeparator7.Size = New System.Drawing.Size(57, 6)
    '
    'splTop
    '
    Me.splTop.Dock = System.Windows.Forms.DockStyle.Fill
    Me.splTop.Location = New System.Drawing.Point(0, 97)
    Me.splTop.Name = "splTop"
    Me.splTop.Orientation = System.Windows.Forms.Orientation.Horizontal
    '
    'splTop.Panel1
    '
    Me.splTop.Panel1.Controls.Add(Me.pnlData)
    Me.splTop.Panel1.Padding = New System.Windows.Forms.Padding(4)
    '
    'splTop.Panel2
    '
    Me.splTop.Panel2.Controls.Add(Me.splBottom)
    Me.splTop.Panel2.Padding = New System.Windows.Forms.Padding(4)
    Me.splTop.Size = New System.Drawing.Size(992, 437)
    Me.splTop.SplitterDistance = 284
    Me.splTop.SplitterWidth = 8
    Me.splTop.TabIndex = 16
    '
    'splBottom
    '
    Me.splBottom.Dock = System.Windows.Forms.DockStyle.Fill
    Me.splBottom.Location = New System.Drawing.Point(4, 4)
    Me.splBottom.Name = "splBottom"
    Me.splBottom.Orientation = System.Windows.Forms.Orientation.Horizontal
    '
    'splBottom.Panel1
    '
    Me.splBottom.Panel1.Controls.Add(Me.pnlSteps)
    '
    'splBottom.Panel2
    '
    Me.splBottom.Panel2.Controls.Add(Me.txtFilter)
    Me.splBottom.Size = New System.Drawing.Size(984, 137)
    Me.splBottom.SplitterDistance = 85
    Me.splBottom.SplitterWidth = 8
    Me.splBottom.TabIndex = 0
    '
    'ctxMenuStrip
    '
    Me.ctxMenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ctxsFilterSelection, Me.ctxsExcludeColumnSelection, Me.ctxsEnterFilter, Me.ctxsFilterEmpty, Me.ctxsFilterNotEmpty, Me.ctxsClearColumnFilter, Me.ctxsClearFilterOnAllColumns})
    Me.ctxMenuStrip.Name = "ctxMenuStrip"
    Me.ctxMenuStrip.Size = New System.Drawing.Size(300, 158)
    '
    'ctxsFilterSelection
    '
    Me.ctxsFilterSelection.Name = "ctxsFilterSelection"
    Me.ctxsFilterSelection.Size = New System.Drawing.Size(299, 22)
    Me.ctxsFilterSelection.Tag = "0"
    Me.ctxsFilterSelection.Text = "Filter By Column Selection"
    '
    'ctxsExcludeColumnSelection
    '
    Me.ctxsExcludeColumnSelection.Name = "ctxsExcludeColumnSelection"
    Me.ctxsExcludeColumnSelection.Size = New System.Drawing.Size(299, 22)
    Me.ctxsExcludeColumnSelection.Tag = "1"
    Me.ctxsExcludeColumnSelection.Text = "Filter Excluding Column Selection"
    '
    'ctxsEnterFilter
    '
    Me.ctxsEnterFilter.Name = "ctxsEnterFilter"
    Me.ctxsEnterFilter.Size = New System.Drawing.Size(299, 22)
    Me.ctxsEnterFilter.Tag = "2"
    Me.ctxsEnterFilter.Text = "Enter Filter..."
    '
    'ctxsFilterEmpty
    '
    Me.ctxsFilterEmpty.Name = "ctxsFilterEmpty"
    Me.ctxsFilterEmpty.Size = New System.Drawing.Size(299, 22)
    Me.ctxsFilterEmpty.Tag = "3"
    Me.ctxsFilterEmpty.Text = "Filter for Empty"
    '
    'ctxsFilterNotEmpty
    '
    Me.ctxsFilterNotEmpty.Name = "ctxsFilterNotEmpty"
    Me.ctxsFilterNotEmpty.Size = New System.Drawing.Size(299, 22)
    Me.ctxsFilterNotEmpty.Tag = "4"
    Me.ctxsFilterNotEmpty.Text = "Filter for Not Empty"
    '
    'ctxsClearColumnFilter
    '
    Me.ctxsClearColumnFilter.Name = "ctxsClearColumnFilter"
    Me.ctxsClearColumnFilter.Size = New System.Drawing.Size(299, 22)
    Me.ctxsClearColumnFilter.Tag = "5"
    Me.ctxsClearColumnFilter.Text = "Clear Column Filter"
    '
    'ctxsClearFilterOnAllColumns
    '
    Me.ctxsClearFilterOnAllColumns.Name = "ctxsClearFilterOnAllColumns"
    Me.ctxsClearFilterOnAllColumns.Size = New System.Drawing.Size(299, 22)
    Me.ctxsClearFilterOnAllColumns.Tag = "6"
    Me.ctxsClearFilterOnAllColumns.Text = "Clear Filter on All Columns"
    '
    'frmListManager
    '
    Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
    Me.ClientSize = New System.Drawing.Size(992, 534)
    Me.Controls.Add(Me.splTop)
    Me.Controls.Add(Me.sbr)
    Me.Controls.Add(Me.sbrStatus)
    Me.Controls.Add(Me.tsp)
    Me.Controls.Add(Me.mnu)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.MainMenuStrip = Me.mnu
    Me.Name = "frmListManager"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    CType(Me.sbpRows, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.sbpSelected, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.sbpViewing, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.sbpSteps, System.ComponentModel.ISupportInitialize).EndInit()
    Me.pnlData.ResumeLayout(False)
    Me.pnlSteps.ResumeLayout(False)
    Me.pnlStepCommands.ResumeLayout(False)
    Me.mnu.ResumeLayout(False)
    Me.mnu.PerformLayout()
    Me.splTop.Panel1.ResumeLayout(False)
    Me.splTop.Panel2.ResumeLayout(False)
    Me.splTop.ResumeLayout(False)
    Me.splBottom.Panel1.ResumeLayout(False)
    Me.splBottom.Panel2.ResumeLayout(False)
    Me.splBottom.Panel2.PerformLayout()
    Me.splBottom.ResumeLayout(False)
    Me.ctxMenuStrip.ResumeLayout(False)
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub

#End Region

  Private Enum PopupMenuItems
    ctxsFilterSelection
    ctxsFilterExcludingSelection
    ctxsEnterFilter
    ctxsFilterForEmpty
    ctxsFilterNotEmpty
    ctxsClearColumnFilter
    ctxsClearAllFilter
  End Enum

  Private Enum CommandIndexes
    'There are icons that match each of these items in both imgToolbar16 and imgToolbar32
    'They must be kept in synch so be careful
    cbiNew
    cbiOpenSteps
    cbiSaveSteps
    cbiDeleteSteps
    cbiRebuild
    cbiMail
    cbiSaveList
    cbiReport
    cbiSelect
    cbiRemove
    cbiReplace
    cbiViewSelected
    cbiViewUnSelected
    cbiSaveData
    cbiPrintData
    cbiViewSteps
    cbiViewData
    '-------------------   Items below here have no icons as yet
    cbiViewNameFields
    cbiViewDescriptions
    cbiViewIDNumbers
    cbiClear
    cbiSetDefaultAddress
    cbiUndo
    cbiExit
    '-------------------  Items below here are special items with no icons or menu items
    cbiSeparator
    cbiSelectView
  End Enum

  Private mvComboSelectionValid As Boolean
  Private mvInitialised As Boolean

  Private mvSSNumber As Integer                     'Selection Set Number
  Private mvCSNumber As Integer                     'Criteria Set Number
  Private mvNewCS As Boolean
  Private mvNewSS As Boolean

  Private mvFilterChanged As Boolean
  Private mvCurrentView As String
  Private mvCurrentViewDesc As String
  Private mvRestrictions As DBFields
  Private mvLastRestrictions As DBFields

  Private mvShowNameFields As Boolean               'Flag if name fields should be shown or hidden
  Private mvShowCodeDescriptions As Boolean         'Flag if code descriptions should be shown or hidden
  Private mvShowIDNumbers As Boolean                'Flag if id numbers should be shown or hidden
  Private mvViewingSelection As Boolean             'Flag if we are viewing the selected records or not
  Private mvFilter As Boolean                       'Filter has been selected
  Private mvFilterCount As Integer                  'Count of filtered records
  Private mvRecordsSelected As Integer              'Count of records selected
  Private mvDefaultIndex As Integer
  Private mvSequenceNumber As Integer
  Private mvViewsComboBox As ComboBox
  Private mvViewsToolStripComboBox As ToolStripComboBox
  Private mvMenuItems As New CollectionList(Of MenuToolbarCommand)
  Private mvBrowserMenu As New BrowserMenu(Nothing)

  Private Sub InitialiseControls()
#If DEBUG Then
    Static mvDebugTest As New DebugTest(Me)       'Include this statement so we can keep track of memory leakage of forms
#End If
    Me.BackColor = DisplayTheme.FormBackColor
    BuildMenus()
    BuildToolbar()
    mvViewsComboBox.ValueMember = "ViewName"
    mvViewsComboBox.DisplayMember = "ViewDescription"
    DataHelper.FillComboBox(mvViewsComboBox, CareServices.XMLLookupDataTypes.xldtListManagerViews, False)
    mvViewsComboBox.Width = 200
    dgr.Clear()
    mvSequenceNumber = 1
    Dim vList As ParameterList = DataHelper.InitListManager()
    mvSSNumber = vList.IntegerValue("SelectionSetNumber")
    mvNewSS = True
    mvCSNumber = 999999
    mvNewCS = True
    dgrSteps.Populate(DataHelper.GetCriteriaSetSteps(mvCSNumber))
    dgr.SetCellsEditable()
    dgr.SetRowHeaderVisible()
    dgr.ContextMenuStrip = ctxMenuStrip
    mvShowCodeDescriptions = True
    mvComboSelectionValid = True
    FormHelper.SelectComboBoxItem(mvViewsComboBox, "VContacts")
    mvDefaultIndex = mvViewsComboBox.SelectedIndex
    mvInitialised = True
  End Sub

  Private Sub BuildMenus()
    Me.FileToolStripMenuItem.Text = ControlText.mnuLMFile
    Me.ViewToolStripMenuItem.Text = ControlText.mnuLMView
    Me.DataToolStripMenuItem.Text = ControlText.mnuLMData
    With mvMenuItems
      .Add(CommandIndexes.cbiOpenSteps.ToString, New MenuToolbarCommand("OpenSteps", ControlText.MnuLMOpenSteps, CommandIndexes.cbiOpenSteps, , imgToolbar16.Images(CommandIndexes.cbiOpenSteps), ControlText.mnuLMOpenStepsTT))
      .Add(CommandIndexes.cbiSaveSteps.ToString, New MenuToolbarCommand("SaveSteps", ControlText.MnuLMSaveSteps, CommandIndexes.cbiSaveSteps, , imgToolbar16.Images(CommandIndexes.cbiSaveSteps), ControlText.mnuLMSaveStepsTT))
      .Add(CommandIndexes.cbiDeleteSteps.ToString, New MenuToolbarCommand("DeleteStep", ControlText.MnuLMDeleteStep, CommandIndexes.cbiDeleteSteps, , imgToolbar16.Images(CommandIndexes.cbiDeleteSteps), ControlText.mnuLMDeleteStepTT))
      .Add(CommandIndexes.cbiNew.ToString, New MenuToolbarCommand("NewList", ControlText.MnuLMNewList, CommandIndexes.cbiNew, , imgToolbar16.Images(CommandIndexes.cbiNew), ControlText.mnuLMNewListTT))
      .Add(CommandIndexes.cbiRebuild.ToString, New MenuToolbarCommand("RebuildList", ControlText.MnuLMRebuildList, CommandIndexes.cbiRebuild, , imgToolbar16.Images(CommandIndexes.cbiRebuild), ControlText.mnuLMRebuildListTT))
      .Add(CommandIndexes.cbiMail.ToString, New MenuToolbarCommand("MailList", ControlText.MnuLMMailList, CommandIndexes.cbiMail, , imgToolbar16.Images(CommandIndexes.cbiMail), ControlText.mnuLMMailListTT))
      .Add(CommandIndexes.cbiSaveList.ToString, New MenuToolbarCommand("SaveList", ControlText.MnuLMSaveList, CommandIndexes.cbiSaveList, , imgToolbar16.Images(CommandIndexes.cbiSaveList), ControlText.mnuLMSaveListTT))
      .Add(CommandIndexes.cbiReport.ToString, New MenuToolbarCommand("ReportList", ControlText.MnuLMReportList, CommandIndexes.cbiReport, , imgToolbar16.Images(CommandIndexes.cbiReport), ControlText.mnuLMReportListTT))
      .Add(CommandIndexes.cbiExit.ToString, New MenuToolbarCommand("Exit", ControlText.mnuLMExit, CommandIndexes.cbiExit))
      .Add(CommandIndexes.cbiViewSelected.ToString, New MenuToolbarCommand("ViewSelected", ControlText.MnuLMSelectedRecords, CommandIndexes.cbiViewSelected, , imgToolbar16.Images(CommandIndexes.cbiViewSelected), ControlText.mnuLMSelectedRecordsTT))
      .Add(CommandIndexes.cbiViewUnSelected.ToString, New MenuToolbarCommand("ViewUnSelected", ControlText.MnuLMUnSelectedRecords, CommandIndexes.cbiViewUnSelected, , imgToolbar16.Images(CommandIndexes.cbiViewUnSelected), ControlText.mnuLMUnSelectedRecordsTT))
      .Add(CommandIndexes.cbiViewData.ToString, New MenuToolbarCommand("ViewData", ControlText.MnuLMData, CommandIndexes.cbiViewData, , imgToolbar16.Images(CommandIndexes.cbiViewData), ControlText.mnuLMViewDataTT))
      .Add(CommandIndexes.cbiViewSteps.ToString, New MenuToolbarCommand("ViewSteps", ControlText.MnuLMSteps, CommandIndexes.cbiViewSteps, , imgToolbar16.Images(CommandIndexes.cbiViewSteps), ControlText.mnuLMViewStepsTT))
      .Add(CommandIndexes.cbiViewNameFields.ToString, New MenuToolbarCommand("ViewNameFields", ControlText.mnuLMNameFields, CommandIndexes.cbiViewNameFields))
      .Add(CommandIndexes.cbiViewDescriptions.ToString, New MenuToolbarCommand("ViewDescriptions", ControlText.mnuLMCodeDescriptions, CommandIndexes.cbiViewDescriptions))
      .Add(CommandIndexes.cbiViewIDNumbers.ToString, New MenuToolbarCommand("ViewNumbers", ControlText.mnuLMIDNumbers, CommandIndexes.cbiViewIDNumbers))
      .Add(CommandIndexes.cbiSelect.ToString, New MenuToolbarCommand("SelectRecords", ControlText.MnuLMSelectRecords, CommandIndexes.cbiSelect, , imgToolbar16.Images(CommandIndexes.cbiSelect), ControlText.mnuLMSelectRecordsTT))
      .Add(CommandIndexes.cbiRemove.ToString, New MenuToolbarCommand("RemoveRecords", ControlText.MnuLMRemoveRecords, CommandIndexes.cbiRemove, , imgToolbar16.Images(CommandIndexes.cbiRemove), ControlText.mnuLMRemoveRecordsTT))
      .Add(CommandIndexes.cbiReplace.ToString, New MenuToolbarCommand("ReplaceRecords", ControlText.MnuLMReplaceRecords, CommandIndexes.cbiReplace, , imgToolbar16.Images(CommandIndexes.cbiReplace), ControlText.mnuLMReplaceRecordsTT))
      .Add(CommandIndexes.cbiSetDefaultAddress.ToString, New MenuToolbarCommand("SetDefault", ControlText.mnuLMSetDefaultAddress, CommandIndexes.cbiSetDefaultAddress))
      .Add(CommandIndexes.cbiUndo.ToString, New MenuToolbarCommand("UndoLastChange", ControlText.mnuLMUndoLast, CommandIndexes.cbiUndo))
      .Add(CommandIndexes.cbiClear.ToString, New MenuToolbarCommand("ClearFilter", ControlText.mnuLMClearFilter, CommandIndexes.cbiClear))
      .Add(CommandIndexes.cbiSaveData.ToString, New MenuToolbarCommand("SaveAsFile", ControlText.MnuLMSaveAsFile, CommandIndexes.cbiSaveData, , imgToolbar16.Images(CommandIndexes.cbiSaveData), ControlText.mnuLMSaveAsFileTT))
      .Add(CommandIndexes.cbiPrintData.ToString, New MenuToolbarCommand("Print", ControlText.MnuLMPrint, CommandIndexes.cbiPrintData, , imgToolbar16.Images(CommandIndexes.cbiPrintData), ControlText.mnuLMPrintTT))
    End With
    For Each vItem As MenuToolbarCommand In mvMenuItems
      vItem.OnClick = AddressOf MenuHandler
    Next
    mvMenuItems(CommandIndexes.cbiReport).HideItem = True
    With FileToolStripMenuItem.DropDownItems
      .Add(mvMenuItems(CommandIndexes.cbiOpenSteps.ToString).MenuStripItem)
      .Add(mvMenuItems(CommandIndexes.cbiSaveSteps.ToString).MenuStripItem)
      .Add(mvMenuItems(CommandIndexes.cbiDeleteSteps.ToString).MenuStripItem)
      .Add(New ToolStripSeparator)
      .Add(mvMenuItems(CommandIndexes.cbiNew.ToString).MenuStripItem)
      .Add(mvMenuItems(CommandIndexes.cbiRebuild.ToString).MenuStripItem)
      .Add(mvMenuItems(CommandIndexes.cbiMail.ToString).MenuStripItem)
      .Add(mvMenuItems(CommandIndexes.cbiSaveList.ToString).MenuStripItem)
      If Not mvMenuItems(CommandIndexes.cbiReport).HideItem Then .Add(mvMenuItems(CommandIndexes.cbiReport.ToString).MenuStripItem)
      .Add(New ToolStripSeparator)
      .Add(mvMenuItems(CommandIndexes.cbiExit.ToString).MenuStripItem)
    End With
    With ViewToolStripMenuItem.DropDownItems
      .Add(mvMenuItems(CommandIndexes.cbiViewSelected.ToString).MenuStripItem)
      .Add(mvMenuItems(CommandIndexes.cbiViewUnSelected.ToString).MenuStripItem(True))
      .Add(New ToolStripSeparator)
      .Add(mvMenuItems(CommandIndexes.cbiViewData.ToString).MenuStripItem(True))
      .Add(mvMenuItems(CommandIndexes.cbiViewSteps.ToString).MenuStripItem(True))
      .Add(New ToolStripSeparator)
      .Add(mvMenuItems(CommandIndexes.cbiViewNameFields.ToString).MenuStripItem)
      .Add(mvMenuItems(CommandIndexes.cbiViewDescriptions.ToString).MenuStripItem(True))
      .Add(mvMenuItems(CommandIndexes.cbiViewIDNumbers.ToString).MenuStripItem)
    End With
    With DataToolStripMenuItem.DropDownItems
      .Add(mvMenuItems(CommandIndexes.cbiSelect.ToString).MenuStripItem)
      .Add(mvMenuItems(CommandIndexes.cbiRemove.ToString).MenuStripItem)
      .Add(mvMenuItems(CommandIndexes.cbiReplace.ToString).MenuStripItem)
      .Add(New ToolStripSeparator)
      .Add(mvMenuItems(CommandIndexes.cbiSetDefaultAddress.ToString).MenuStripItem)
      .Add(New ToolStripSeparator)
      .Add(mvMenuItems(CommandIndexes.cbiUndo.ToString).MenuStripItem)
      .Add(mvMenuItems(CommandIndexes.cbiClear.ToString).MenuStripItem)
      .Add(New ToolStripSeparator)
      .Add(mvMenuItems(CommandIndexes.cbiSaveData.ToString).MenuStripItem)
      .Add(mvMenuItems(CommandIndexes.cbiPrintData.ToString).MenuStripItem)
    End With
  End Sub

  Private Sub BuildToolbar()
    tsp.ImageList = imgToolbar32
    tsp.ImageScalingSize = New Size(32, 32)
    tsp.Items.Add(mvMenuItems(CommandIndexes.cbiNew.ToString).ToolStripButton)
    tsp.Items.Add(mvMenuItems(CommandIndexes.cbiOpenSteps.ToString).ToolStripButton)
    tsp.Items.Add(mvMenuItems(CommandIndexes.cbiSaveSteps.ToString).ToolStripButton)
    tsp.Items.Add(mvMenuItems(CommandIndexes.cbiDeleteSteps.ToString).ToolStripButton)
    tsp.Items.Add(New ToolStripSeparator)
    tsp.Items.Add(mvMenuItems(CommandIndexes.cbiRebuild.ToString).ToolStripButton)
    tsp.Items.Add(mvMenuItems(CommandIndexes.cbiMail.ToString).ToolStripButton)
    tsp.Items.Add(mvMenuItems(CommandIndexes.cbiSaveList.ToString).ToolStripButton)

    Dim vLabel As New ToolStripLabel
    vLabel.Text = ControlText.lblDataView
    vLabel.Visible = True
    tsp.Items.Add(vLabel)
    mvViewsToolStripComboBox = New ToolStripComboBox
    mvViewsComboBox = CType(mvViewsToolStripComboBox.Control, ComboBox)
    mvViewsComboBox.DropDownStyle = ComboBoxStyle.DropDownList
    AddHandler mvViewsComboBox.SelectedIndexChanged, AddressOf ViewsSelectedIndexChanged
    mvViewsToolStripComboBox.Visible = True
    tsp.Items.Add(mvViewsToolStripComboBox)

    tsp.Items.Add(mvMenuItems(CommandIndexes.cbiSelect.ToString).ToolStripButton)
    tsp.Items.Add(mvMenuItems(CommandIndexes.cbiRemove.ToString).ToolStripButton)
    tsp.Items.Add(mvMenuItems(CommandIndexes.cbiReplace.ToString).ToolStripButton)
    tsp.Items.Add(mvMenuItems(CommandIndexes.cbiViewSelected.ToString).ToolStripButton)
    tsp.Items.Add(mvMenuItems(CommandIndexes.cbiViewUnSelected.ToString).ToolStripButton)
    tsp.Items.Add(New ToolStripSeparator)
    tsp.Items.Add(mvMenuItems(CommandIndexes.cbiSaveData.ToString).ToolStripButton)
    tsp.Items.Add(mvMenuItems(CommandIndexes.cbiPrintData.ToString).ToolStripButton)
    tsp.Items.Add(mvMenuItems(CommandIndexes.cbiViewSteps.ToString).ToolStripButton(True))
    tsp.Items.Add(mvMenuItems(CommandIndexes.cbiViewData.ToString).ToolStripButton(True))
  End Sub
  Private Sub MenuHandler(ByVal sender As System.Object, ByVal e As System.EventArgs)
    Dim vToolStripItem As ToolStripItem = DirectCast(sender, ToolStripItem)
    Dim vCommand As MenuToolbarCommand = DirectCast(vToolStripItem.Tag, MenuToolbarCommand)
    ProcessMenuItem(CType(vCommand.CommandID, CommandIndexes))
  End Sub
  Private Sub ProcessMenuItem(ByVal pCommand As CommandIndexes)
    Dim vCursor As New BusyCursor
    Try
      Select Case pCommand
        Case CommandIndexes.cbiNew
          NewList()
        Case CommandIndexes.cbiOpenSteps
          GetSteps()
        Case CommandIndexes.cbiSaveSteps
          SaveSteps()
        Case CommandIndexes.cbiDeleteSteps
          DeleteStep()
        Case CommandIndexes.cbiRebuild
          If ShowQuestion(QuestionMessages.qmRebuildList, MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.OK Then DoRebuild()
        Case CommandIndexes.cbiMail
          MailSelection()
        Case CommandIndexes.cbiSaveList
          SaveSelection()
        Case CommandIndexes.cbiReport
          'Not implemented in this release
        Case CommandIndexes.cbiSelect
          ProcessAction("S")
        Case CommandIndexes.cbiRemove
          ProcessAction("R")
        Case CommandIndexes.cbiReplace
          ProcessAction("P")
        Case CommandIndexes.cbiViewSelected
          SetViewingSelection(True)
        Case CommandIndexes.cbiViewUnSelected
          SetViewingSelection(False)
        Case CommandIndexes.cbiSaveData
          If CheckForAllRecordsInList() Then SaveAsFile()
        Case CommandIndexes.cbiPrintData
          If CheckForAllRecordsInList() Then dgr.Print()
        Case CommandIndexes.cbiViewSteps
          splTop.Panel2Collapsed = Not splTop.Panel2Collapsed
          mvMenuItems(CommandIndexes.cbiViewData.ToString).CheckToolStripItem(mnu, tsp, Not splTop.Panel1Collapsed)
          mvMenuItems(CommandIndexes.cbiViewSteps.ToString).CheckToolStripItem(mnu, tsp, Not splTop.Panel2Collapsed)
        Case CommandIndexes.cbiViewData
          splTop.Panel1Collapsed = Not splTop.Panel1Collapsed
          mvMenuItems(CommandIndexes.cbiViewData.ToString).CheckToolStripItem(mnu, tsp, Not splTop.Panel1Collapsed)
          mvMenuItems(CommandIndexes.cbiViewSteps.ToString).CheckToolStripItem(mnu, tsp, Not splTop.Panel2Collapsed)
        Case CommandIndexes.cbiViewNameFields
          mvShowNameFields = Not mvShowNameFields
          ShowViewData(mvCurrentView, False)
        Case CommandIndexes.cbiViewDescriptions
          mvShowCodeDescriptions = Not mvShowCodeDescriptions
          ShowViewData(mvCurrentView, False)
        Case CommandIndexes.cbiViewIDNumbers
          mvShowIDNumbers = Not mvShowIDNumbers
          ShowViewData(mvCurrentView, False)
        Case CommandIndexes.cbiClear
          ClearCurrentFilter()
        Case CommandIndexes.cbiSetDefaultAddress
          ProcessAction("D")
        Case CommandIndexes.cbiUndo
          UndoLastFilterChange()
        Case CommandIndexes.cbiExit
          Me.Close()
      End Select
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Private Sub frmListManager_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
    Me.Text = ControlText.frmListManager
    lblSteps.Text = ControlText.LblSteps
    lblData.Text = ControlText.LblProcessingWait
    ctxsFilterSelection.Text = ControlText.mnuLMPColumnSelection
    ctxsExcludeColumnSelection.Text = ControlText.mnuLMPExColumnSelection
    ctxsEnterFilter.Text = ControlText.mnuLMPEnterFilter
    ctxsFilterEmpty.Text = ControlText.mnuLMPFilterEmpty
    ctxsFilterNotEmpty.Text = ControlText.mnuLMPFilterNotEmpty
    ctxsClearColumnFilter.Text = ControlText.mnuLMPClearColumn
    ctxsClearFilterOnAllColumns.Text = ControlText.MnuLMPClearAll
    ShowSelected(0)
    imbUp.Enabled = False
    imbDown.Enabled = False
    imbDelete.Enabled = False
  End Sub
  Private Sub frmListManager_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
    DataHelper.TerminateListManager(mvSSNumber, mvCSNumber)
  End Sub
  Private Sub ctxMenuStrip_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ctxsFilterSelection.Click, ctxsExcludeColumnSelection.Click, ctxsEnterFilter.Click, ctxsFilterEmpty.Click, ctxsFilterNotEmpty.Click, ctxsClearColumnFilter.Click, ctxsClearFilterOnAllColumns.Click
    Dim vCursor As New BusyCursor
    Try
      Dim vMenuItem As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)
      Dim vPopupMenuItem As PopupMenuItems = (CType(vMenuItem.Tag, PopupMenuItems))
      Dim vOperator As DBField.FieldWhereOperators
      Dim vName As String = dgr.ColumnName(dgr.ActiveColumn)
      Dim vValue As String = dgr.GetValue(dgr.ActiveRow, dgr.ActiveColumn)
      Dim vFieldType As DBField.FieldTypes = dgr.GetDataType(dgr.ActiveColumn)
      Dim vSelectedText As String = dgr.SelectedText()
      Dim vDone As Boolean

      If vSelectedText.Length > 0 Then
        If vValue.StartsWith(vSelectedText) Then
          vValue = vSelectedText & "*"
        ElseIf vValue.EndsWith(vSelectedText) Then
          vValue = "*" & vSelectedText
        Else
          vValue = "*" & vSelectedText & "*"
        End If
        vOperator = DBField.FieldWhereOperators.fwoLike
      Else
        If vName = "address" Or vName = "house_name" Then
          vOperator = DBField.FieldWhereOperators.fwoLike
        Else
          vOperator = DBField.FieldWhereOperators.fwoEqual
        End If
      End If
      Select Case vPopupMenuItem
        Case PopupMenuItems.ctxsFilterSelection
          SetColumnFilter(vName, vFieldType, vValue, vOperator)
        Case PopupMenuItems.ctxsFilterExcludingSelection
          If vOperator = DBField.FieldWhereOperators.fwoLike Then
            vOperator = DBField.FieldWhereOperators.fwoNotLike
          Else
            vOperator = DBField.FieldWhereOperators.fwoNotEqual
          End If
          SetColumnFilter(vName, vFieldType, vValue, vOperator)
        Case PopupMenuItems.ctxsEnterFilter
          Dim vList As New ParameterList(True)
          vList("TableName") = mvCurrentView
          vList("FieldName") = vName
          Dim vFieldChar As String
          Select Case vFieldType
            Case DBField.FieldTypes.cftDate
              vFieldChar = "D"
            Case DBField.FieldTypes.cftInteger, DBField.FieldTypes.cftLong
              vFieldChar = "I"
            Case DBField.FieldTypes.cftMemo
              vFieldChar = "M"
            Case DBField.FieldTypes.cftNumeric
              vFieldChar = "N"
            Case DBField.FieldTypes.cftTime
              vFieldChar = "T"
            Case Else
              vFieldChar = "C"
          End Select
          vList("FieldType") = vFieldChar
          vList("FieldNameDesc") = dgr.ColumnHeading(dgr.ActiveColumn)
          Dim vParams As ParameterList = DataHelper.GetMaintenanceData(vList)
          'The field name returned by this call may be different than we started with e.g address_valid_from changes to valid_from
          'make sure we continue to use the original field name
          vParams("AttributeName") = vName

          If vParams.ContainsKey("RestrictionAttribute") AndAlso vParams("RestrictionAttribute").Length > 0 Then
            Dim vRestrictCol As Integer = dgr.GetColumn(vParams("RestrictionAttribute"))
            If vRestrictCol = 0 Then
              ShowInformationMessage(String.Format(InformationMessages.imCannotEnterFilter, vParams("RestrictionAttribute")))
              vDone = True
            Else
              vParams("RestrictionValue") = dgr.GetValue(dgr.ActiveRow, vRestrictCol)
            End If
          End If
          If Not vDone Then
            Dim vForm As New frmFilter(vParams, vFieldType, mvRestrictions, mvLastRestrictions)
            If vForm.ShowDialog(Me) = Windows.Forms.DialogResult.Cancel Then vDone = True
          End If
        Case PopupMenuItems.ctxsFilterForEmpty
          vOperator = DBField.FieldWhereOperators.fwoEqual
          SetColumnFilter(vName, vFieldType, "", vOperator)
        Case PopupMenuItems.ctxsFilterNotEmpty
          vOperator = DBField.FieldWhereOperators.fwoNotEqual
          SetColumnFilter(vName, vFieldType, "", vOperator)
        Case PopupMenuItems.ctxsClearColumnFilter
          SetColumnFilter(vName, vFieldType, "", DBField.FieldWhereOperators.fwoEqual)
          mvRestrictions.Remove(vName)
        Case PopupMenuItems.ctxsClearAllFilter
          ClearCurrentFilter()
          vDone = True
      End Select
      If Not vDone Then ShowViewData(mvCurrentView, False)
      dgr.Focus()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub
  Private Sub ViewsSelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    If mvComboSelectionValid Then
      Try
        mvRestrictions = New DBFields
        mvLastRestrictions = New DBFields
        mvFilterChanged = False
        If mvViewsComboBox.SelectedIndex >= 0 Then
          Dim vRow As DataRowView = CType(mvViewsComboBox.SelectedItem, DataRowView)
          mvViewsToolStripComboBox.ToolTipText = vRow.Item("Notes").ToString
          mvCurrentView = vRow.Item("ViewName").ToString
          mvCurrentViewDesc = vRow.Item("ViewDescription").ToString
          ShowViewData(mvCurrentView, True)
        Else
          mvCurrentView = ""
          mvCurrentViewDesc = ""
          mvViewsToolStripComboBox.ToolTipText = ""
          dgr.Clear()
        End If
        If splTop.Panel1Collapsed = False Then dgr.Focus()
        Exit Sub
      Catch vException As Exception
        DataHelper.HandleException(vException)
      End Try
    End If
  End Sub
  Private Sub dgrSteps_RowSelected(ByVal pSender As Object, ByVal pRow As Integer, ByVal pDataRow As Integer) Handles dgrSteps.RowSelected
    Dim vFilter As String

    imbUp.Enabled = pRow > 0
    imbDown.Enabled = pRow < dgrSteps.RowCount - 1
    vFilter = dgrSteps.GetValue(pRow, dgrSteps.GetColumn("FilterSQL"))
    vFilter = Replace$(vFilter, "AND", vbCrLf & "AND")
    txtFilter.Text = vFilter
  End Sub
  Private Sub imbDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles imbDelete.Click
    DeleteStep()
  End Sub
  Private Sub imbUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles imbUp.Click
    Dim vCursor As New BusyCursor
    Try
      Dim vRow As Integer = dgrSteps.CurrentRow
      dgrSteps.MoveRow(False)
      dgrSteps.SelectRow(vRow - 1)
      ResequenceSteps(vRow)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub
  Private Sub imbDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles imbDown.Click
    Dim vCursor As New BusyCursor
    Try
      Dim vRow As Integer = dgrSteps.CurrentRow
      dgrSteps.MoveRow(True)
      dgrSteps.SelectRow(vRow + 1)
      ResequenceSteps(vRow)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Private Sub ShowViewData(ByVal pViewName As String, ByVal pClearSelection As Boolean, Optional ByVal pAllRecords As Boolean = False)
    ShowStepInfo()
    Dim vRetainLeftColumn As Boolean
    If pClearSelection Then
      mvRestrictions.Clear()
      mvLastRestrictions.Clear()
      mvFilterChanged = False
      mvFilter = False
    Else
      vRetainLeftColumn = True
    End If
    If pViewName.Length > 0 Then
      StartProcessing()
      ShowStatus(InformationMessages.imReadingData)
      mvFilter = mvRestrictions.Count > 0
      Dim vFilter As String = mvRestrictions.WhereClause
      Dim vHideColumns As String = ""
      If Not mvShowIDNumbers Then vHideColumns = "Numbers"
      If Not mvShowNameFields Then vHideColumns &= "Names"
      If Not mvShowCodeDescriptions Then vHideColumns &= "Descriptions"
      dgr.HeaderLines = 2
      dgr.Populate(DataHelper.GetListManagerData(mvSSNumber, pViewName, vHideColumns, vFilter, pAllRecords, mvViewingSelection), , , vRetainLeftColumn)
      dgr.SetColumnVisible("contact_number", mvShowIDNumbers)
      If mvFilter Then
        Dim vMaxRecords As Long = 1000
        If pAllRecords Then vMaxRecords = 32767
        If dgr.RowCount < vMaxRecords Then
          mvFilterCount = dgr.RowCount
        Else
          ShowStatus(InformationMessages.imCountingRecords)
          Dim vList As ParameterList = DataHelper.GetListManagerCount(mvSSNumber, pViewName, vFilter, mvViewingSelection)
          mvFilterCount = CInt(vList("RecordCount"))
        End If
      Else
        mvFilterCount = 0
      End If
      For Each vField As DBField In mvRestrictions
        If vField.Name.IndexOf("#") < 0 Then dgr.HighLightColumn(vField.Name, True)
      Next
      'dgr.LockDateColumns()
      EndProcessing()
    End If
    ShowRecordInfo()
    SetMenuToolbarItems()
  End Sub
  Private Sub SetColumnFilter(ByVal pName As String, ByVal pFieldType As DBField.FieldTypes, ByVal pValue As String, ByVal pOperator As DBField.FieldWhereOperators)
    SaveRestrictions()
    If mvRestrictions.ContainsKey(pName) Then
      mvRestrictions(pName).Value = pValue
    Else
      mvRestrictions.Add(pName, pFieldType, pValue)
    End If
    mvRestrictions(pName).WhereOperator = pOperator
    Dim vIndex As Integer = 1
    While mvRestrictions.ContainsKey(pName & "#" & vIndex)
      mvRestrictions.Remove(pName & "#" & vIndex)
      vIndex = vIndex + 1
    End While
  End Sub
  Private Sub ProcessAction(ByVal pAction As String)
    Dim vFilter As String

    vFilter = mvRestrictions.WhereClause      'TODO Handle Alias "x"
    StartProcessing()
    ProcessStep(mvSequenceNumber, pAction, mvCurrentView, mvCurrentViewDesc, vFilter, True)
    'SetStepWidths()
    mvSequenceNumber = mvSequenceNumber + 1
    ShowViewData(mvCurrentView, True)
    EndProcessing()
  End Sub
  Private Sub ProcessStep(ByVal pStepNumber As Integer, ByVal pAction As String, ByVal pView As String, ByVal pViewDesc As String, ByVal pFilter As String, ByVal pNewStep As Boolean)
    Select Case pAction
      Case "R"
        ShowStatus(GetInformationMessage(InformationMessages.imLMStepRemoving, pStepNumber.ToString))
      Case "S"
        ShowStatus(GetInformationMessage(InformationMessages.imLMStepAdding, pStepNumber.ToString))
      Case "P"
        ShowStatus(String.Format(GetInformationMessage(InformationMessages.imLMStepReplacing, pStepNumber.ToString)))
      Case "D"
        ShowStatus(String.Format(GetInformationMessage(InformationMessages.imLMDefaultAddress, pStepNumber.ToString)))
    End Select
    Dim vList As ParameterList = DataHelper.ProcessListManagerStep(mvSSNumber, pView, pFilter, pAction)
    mvRecordsSelected = vList.IntegerValue("RecordCount")
    ShowSelected(mvRecordsSelected)
    If pNewStep Then
      Dim vItems As New ParameterList(False)
      vItems("SequenceNumber") = pStepNumber.ToString
      vItems("SelectActionDesc") = GetActionDesc(pAction)
      vItems("ViewDesc") = pViewDesc
      vItems("FilterSQL") = pFilter
      vItems("RecordCount") = mvRecordsSelected.ToString
      vItems("SelectAction") = pAction
      vItems("ViewName") = pView
      dgrSteps.AddDataRow(vItems)
    Else
      With dgrSteps
        .SetValue(pStepNumber - 1, "RecordCount", mvRecordsSelected.ToString)
      End With
    End If
  End Sub

  Private Function CheckForAllRecordsInList() As Boolean
    Dim vResult As DialogResult = Windows.Forms.DialogResult.Yes
    If dgr.RowCount > 0 Then
      If mvFilterCount > dgr.RowCount Then
        If mvFilterCount < 32768 Then
          vResult = ShowQuestion(QuestionMessages.qmLoadAllRecords, MessageBoxButtons.YesNoCancel, dgr.RowCount.ToString, mvFilterCount.ToString)
        Else
          vResult = ShowQuestion(QuestionMessages.qmLoadMaxRecords, MessageBoxButtons.YesNoCancel, dgr.RowCount.ToString, mvFilterCount.ToString)
        End If
        If vResult = Windows.Forms.DialogResult.Yes Then ShowViewData(mvCurrentView, False, True)
      End If
    End If
    CheckForAllRecordsInList = (vResult <> Windows.Forms.DialogResult.Cancel)
  End Function
  Private Function GetActionDesc(ByVal pSelectAction As String) As String
    Select Case pSelectAction
      Case "S"
        Return InformationMessages.imListManagerSelected
      Case "R"
        Return InformationMessages.imListManagerRemoved
      Case "P"
        Return InformationMessages.imListManagerReplaced
      Case "D"
        Return InformationMessages.imListManagerSetDefault
      Case Else
        Return ""
    End Select
  End Function

  Private Sub NewList()
    'Process setting up a new list
    Dim vCursor As New BusyCursor
    Try
      ClearSelection()
      mvViewingSelection = False
      ShowViewData(mvCurrentView, True)
      txtFilter.Text = ""
      mvSequenceNumber = 1
      'Now clear the steps and start with a new criteria set number
      mvCSNumber = 999999
      dgrSteps.Populate(DataHelper.GetCriteriaSetSteps(mvCSNumber))
      lblSteps.Text = ControlText.lblSteps
      SetMenuToolbarItems()
      mvNewCS = True
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub
  Private Sub GetSteps()
    'Process selecting a list of steps (criteria set)
    Dim vCursor As New BusyCursor
    Try
      Dim vList As New ParameterList(True)
      vList("ListManager") = "Y"
      Dim vSI As New frmSimpleFinder
      vSI.Init(CareServices.XMLLookupDataTypes.xldtCriteriaSets, False, vList)
      If vSI.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
        Dim vResult As DialogResult
        If dgrSteps.RowCount > 0 Then
          vResult = ShowQuestion(QuestionMessages.qmReplaceSteps, MessageBoxButtons.YesNoCancel)
        Else
          vResult = Windows.Forms.DialogResult.Yes
        End If
        If vResult <> Windows.Forms.DialogResult.Cancel Then
          If vResult = Windows.Forms.DialogResult.Yes Then
            mvCSNumber = CInt(vSI.ResultValue)
            lblSteps.Text = vSI.ResultDescription
          End If
          SetSteps(CInt(vSI.ResultValue), vResult = Windows.Forms.DialogResult.No)
        End If
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub
  Private Sub SaveSelection()
    Dim vCursor As New BusyCursor
    Try
      Dim vCancel As Boolean
      If Not mvNewSS Then
        If ShowQuestion(QuestionMessages.qmUpdateSelectionSet, MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.Cancel Then vCancel = True
      Else
        Dim vForm As frmCardMaintenance = New frmCardMaintenance(CareServices.XMLMaintenanceControlTypes.xmctSelectionSet, mvSSNumber, Nothing)
        If vForm.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
          mvNewSS = False
        Else
          vCancel = True
        End If
      End If
      If Not vCancel Then
        ShowStatus(InformationMessages.imSavingSelection)
        Dim vList As New ParameterList(True)
        vList.IntegerValue("SelectionSetNumber") = mvSSNumber
        DataHelper.SaveListManagerSelection(vList)
      End If
      ShowRecordInfo()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Private Sub MailSelection()
    Dim vForm As frmCardMaintenance = New frmCardMaintenance(CareServices.XMLMaintenanceControlTypes.xmctMailingOptions, 0, Nothing)
    If vForm.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
      Dim vList As ParameterList = vForm.ReturnList
      vList.IntegerValue("SelectionSetNumber") = mvSSNumber
      Dim vFileName As String = vList("MailingFilename")
      vList.Remove("MailingFilename")
      Dim vProcessMailMerge As Boolean = False
      If vList("StandardDocument").Length > 0 Then
        vFileName = DataHelper.GetTempFile(".csv")
        vProcessMailMerge = True
      End If
      Dim vPreviewEMail As Boolean = vList("PreviewEMail") = "Y"
      vList.Remove("PreviewEMail")
      Dim vEMailDocument As String = vList("EMailDocument")
      vList.Remove("EMailDocument")
      If DataHelper.GetListManagerMailingFile(vList, vFileName) Then
        Dim vEMailFileName As String = ""
        Dim vRowCount As Integer
        Dim vEmailRowCount As Integer
        Dim vDialogResult As DialogResult = Windows.Forms.DialogResult.OK
        If vList("GenerateEMail") = "Y" Then
          DataHelper.SplitFile(vFileName, vEMailFileName, vRowCount, vEmailRowCount)
          If vEmailRowCount = 0 Then
            ShowInformationMessage(InformationMessages.imNoEMailRecords)
          Else
            If vPreviewEMail Then
              Dim vSDList As New ParameterList(True)
              vSDList("StandardDocument") = vEMailDocument
              Dim vRow As DataRow = DataHelper.GetLookupData(CareServices.XMLLookupDataTypes.xldtStandardDocuments, vSDList).Rows(0)
              Dim vApplication As ExternalApplication = GetDocumentApplication(vRow.Item("DocFileExtension").ToString)
              vDialogResult = vApplication.MergeStandardDocument(vRow.Item("StandardDocument").ToString, vRow.Item("DocFileExtension").ToString, vEMailFileName, False)
            End If
            If vDialogResult = Windows.Forms.DialogResult.OK Then
              Dim vMailList As New ParameterList(True)
              vMailList("StandardDocument") = vEMailDocument
              vMailList("EMailAddress") = vList("EMailAddress")
              vMailList("Name") = vList("Name")
              Dim vEmailReturn As ParameterList = DataHelper.ProcessBulkEMail(vEMailFileName, vMailList)
              ShowInformationMessage(InformationMessages.imEMailComplete, vEmailReturn("Count"))
            Else
              If vList("NoMailingHistory") = "Y" Then
                ShowInformationMessage(InformationMessages.imEmailCancelled)
              Else
                ShowInformationMessage(InformationMessages.imEmailCancelledMHCreated)
              End If
            End If
          End If
        Else
          vRowCount = 1     'Don't actually know how many rows in this case but have to assume there are some
        End If
        If vDialogResult = Windows.Forms.DialogResult.OK AndAlso vRowCount > 0 Then
          If vProcessMailMerge Then
            Dim vSDList As New ParameterList(True)
            vSDList("StandardDocument") = vList("StandardDocument")
            Dim vRow As DataRow = DataHelper.GetLookupData(CareServices.XMLLookupDataTypes.xldtStandardDocuments, vSDList).Rows(0)
            Dim vApplication As ExternalApplication = GetDocumentApplication(vRow.Item("DocFileExtension").ToString)
            vApplication.MergeStandardDocument(vRow.Item("StandardDocument").ToString, vRow.Item("DocFileExtension").ToString, vFileName, False)
          Else
            ShowInformationMessage(InformationMessages.imMailingComplete, vFileName)
          End If
        End If
      End If
    End If
  End Sub

  Private Sub SaveSteps()
    Dim vCancel As Boolean
    If Not mvNewCS Then
      If ShowQuestion(QuestionMessages.qmUpdateSteps, MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.OK Then
        DataHelper.DeleteSelectionSteps(mvCSNumber, 0)
      Else
        vCancel = True
      End If
    Else
      Dim vForm As frmCardMaintenance = New frmCardMaintenance(CareServices.XMLMaintenanceControlTypes.xmctCriterialSet, 0, Nothing)
      If vForm.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
        Dim vList As ParameterList = vForm.ReturnList
        mvCSNumber = vList.IntegerValue("CriteriaSetNumber")
        lblSteps.Text = vList("CriteriaSetDesc")
        mvNewCS = False
      Else
        vCancel = True
      End If
    End If
    If Not vCancel Then
      Dim vList As New ParameterList(True)
      For vRow As Integer = 0 To dgrSteps.RowCount - 1
        vList("CriteriaSetNumber") = mvCSNumber.ToString
        vList("SequenceNumber") = CStr(vRow + 1)
        vList("ViewName") = dgrSteps.GetValue(vRow, "ViewName")
        vList("Filter") = dgrSteps.GetValue(vRow, "FilterSQL")
        vList("SelectAction") = dgrSteps.GetValue(vRow, "SelectAction")
        vList("RecordCount") = dgrSteps.GetValue(vRow, "RecordCount")
        DataHelper.AddSelectionStep(vList)
      Next
    End If
  End Sub
  Private Sub DoRebuild()
    'Process a rebuild of the list
    Dim vCursor As New BusyCursor
    Try
      ClearSelection()
      StartProcessing()
      With dgrSteps
        For vRow As Integer = 0 To dgrSteps.RowCount - 1
          ProcessStep(vRow + 1, .GetValue(vRow, "SelectAction"), .GetValue(vRow, "ViewName"), .GetValue(vRow, "ViewDesc"), .GetValue(vRow, "FilterSQL"), False)
        Next
      End With
      ShowViewData(mvCurrentView, True)
      EndProcessing()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub
  Private Sub DeleteStep()
    Dim vCursor As New BusyCursor
    Try
      If ShowQuestion(QuestionMessages.qmDeleteStep, MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.OK Then
        Dim vRow As Integer = dgrSteps.CurrentRow
        dgrSteps.DeleteRow(vRow)
        ResequenceSteps(vRow)
        txtFilter.Text = ""
        If dgrSteps.RowCount <= vRow Then vRow = dgrSteps.RowCount - 1
        dgrSteps.SelectRow(vRow)
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub
  Private Sub SaveAsFile()
    Dim vCursor As New BusyCursor
    Try
      With sfd
        .Title = ControlText.dlgTitleSaveListAs
        .Filter = "Spreadsheet Files (*.xls)|*.xls|Tab Separated Files (*.tsv)|*.tsv"
        .DefaultExt = "xls"
        .FileName = ""
        .CheckPathExists = True
        .OverwritePrompt = True
        If .ShowDialog(Me) = Windows.Forms.DialogResult.OK Then dgr.SaveList(.FileName)
      End With
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Private Sub SetSteps(ByVal pCSNumber As Integer, ByVal pAdd As Boolean)
    Dim vRows As Integer = dgrSteps.RowCount
    If pAdd Then
      dgrSteps.AddData(DataHelper.GetCriteriaSetSteps(pCSNumber))
    Else
      dgrSteps.Populate(DataHelper.GetCriteriaSetSteps(pCSNumber))                '
    End If
    dgrSteps.SelectRow(0)
    If pAdd Then
      ResequenceSteps(vRows)
    Else
      mvSequenceNumber = dgrSteps.RowCount + 1
      mvNewCS = False
      SetInvalidView()
    End If
  End Sub
  Private Sub ResequenceSteps(ByVal pFromRow As Integer)
    With dgrSteps
      For vRow As Integer = 0 To .RowCount - 1
        .SetValue(vRow, "SequenceNumber", CStr(vRow + 1))
        If vRow >= pFromRow Then .SetValue(vRow, "RecordCount", "")
      Next
      mvSequenceNumber = .RowCount + 1
    End With
    SetInvalidView()
  End Sub

  Private Sub SetInvalidView()
    dgr.Clear()
    ClearSelection()
  End Sub
  Private Sub ClearSelection()
    ShowStatus(InformationMessages.imClearingSelected)
    If mvSSNumber > 0 Then DataHelper.TerminateListManager(mvSSNumber, mvCSNumber)
    Dim vList As ParameterList = DataHelper.InitListManager()
    mvSSNumber = vList.IntegerValue("SelectionSetNumber")
    mvNewSS = True
    ShowSelected(0)
    mvRecordsSelected = 0
    SetMenuToolbarItems()
    ShowRecordInfo()
  End Sub
  Private Sub UndoLastFilterChange()
    mvRestrictions.Clear()
    mvRestrictions = mvLastRestrictions.Clone
    mvFilterChanged = False
    ShowViewData(mvCurrentView, False)
  End Sub
  Private Sub ClearCurrentFilter()
    SaveRestrictions()
    mvRestrictions.Clear()
    ShowViewData(mvCurrentView, False)
  End Sub
  Private Sub SaveRestrictions()
    mvLastRestrictions = mvRestrictions.Clone
    mvFilterChanged = True
  End Sub
  Private Sub StartProcessing()
    If dgr.Visible Then
      dgr.Visible = False
      lblData.Dock = DockStyle.Fill
      lblData.Visible = True
      lblData.BringToFront()
      lblData.Refresh()
    End If
  End Sub
  Private Sub EndProcessing()
    lblData.Visible = False
    dgr.Visible = True
    dgr.BringToFront()
  End Sub
  Private Sub ShowRecordInfo()
    If dgr.RowCount > 0 Then
      If mvFilterCount > 0 Then
        sbpRows.Text = String.Format(InformationMessages.imListManagerNofNRecords, dgr.RowCount, mvFilterCount)
      Else
        sbpRows.Text = String.Format(InformationMessages.imListManagerRows, dgr.RowCount)
      End If
      If mvViewingSelection Then
        sbpViewing.Text = InformationMessages.imListManagerViewingSelected
      Else
        sbpViewing.Text = InformationMessages.imListmanagerViewingNotSelected
      End If
    Else
      sbpRows.Text = InformationMessages.imListManagerNoRecords
      sbpViewing.Text = ""
    End If
    If mvFilter Then
      ShowStatus(String.Format(InformationMessages.imListManagerFilter, mvRestrictions.WhereClause))
    Else
      ShowStatus(InformationMessages.imListManagerNoFilter)
    End If
  End Sub
  Private Sub ShowSelected(ByVal pCount As Integer)
    If pCount > 0 Then
      sbpSelected.Text = String.Format(ControlText.sbpListCount, pCount)
    Else
      sbpSelected.Text = ControlText.sbpListEmpty
    End If
  End Sub
  Private Sub ShowStatus(ByVal pMsg As String)
    sbrStatus.Text = pMsg
  End Sub
  Private Sub ShowStepInfo()
    If dgrSteps.RowCount > 0 Then
      sbpSteps.Text = String.Format(InformationMessages.imListManagerSteps, dgrSteps.RowCount)
    Else
      sbpSteps.Text = InformationMessages.imListManagerNoSteps
    End If
  End Sub
  Private Sub SetMenuToolbarItems()
    mvMenuItems(CommandIndexes.cbiSelect.ToString).EnableToolStripItem(mnu, tsp, mvFilter And Not mvViewingSelection)
    mvMenuItems(CommandIndexes.cbiRemove.ToString).EnableToolStripItem(mnu, tsp, mvFilter And mvViewingSelection)
    mvMenuItems(CommandIndexes.cbiReplace.ToString).EnableToolStripItem(mnu, tsp, mvFilter And mvViewingSelection)

    mvMenuItems(CommandIndexes.cbiViewSelected.ToString).CheckToolStripItem(mnu, tsp, mvViewingSelection)
    mvMenuItems(CommandIndexes.cbiViewUnSelected.ToString).CheckToolStripItem(mnu, tsp, Not mvViewingSelection)
    Dim vSelected As Boolean = mvRecordsSelected > 0
    If mvViewingSelection Then
      mvMenuItems(CommandIndexes.cbiViewUnSelected.ToString).EnableToolStripItem(mnu, tsp, True)
      mvMenuItems(CommandIndexes.cbiViewSelected.ToString).EnableToolStripItem(mnu, tsp, False)
    Else
      mvMenuItems(CommandIndexes.cbiViewUnSelected.ToString).EnableToolStripItem(mnu, tsp, False)
      mvMenuItems(CommandIndexes.cbiViewSelected.ToString).EnableToolStripItem(mnu, tsp, vSelected)
    End If
    mvMenuItems(CommandIndexes.cbiSetDefaultAddress.ToString).EnableToolStripItem(mnu, Nothing, vSelected)

    mvMenuItems(CommandIndexes.cbiMail.ToString).EnableToolStripItem(mnu, tsp, vSelected)
    mvMenuItems(CommandIndexes.cbiSaveList.ToString).EnableToolStripItem(mnu, tsp, vSelected)
    If Not mvMenuItems(CommandIndexes.cbiReport).HideItem Then mvMenuItems(CommandIndexes.cbiReport.ToString).EnableToolStripItem(mnu, tsp, vSelected)

    Dim vData As Boolean = dgr.RowCount > 1
    mvMenuItems(CommandIndexes.cbiSaveData.ToString).EnableToolStripItem(mnu, tsp, vData)
    mvMenuItems(CommandIndexes.cbiPrintData.ToString).EnableToolStripItem(mnu, tsp, vData)

    mvMenuItems(CommandIndexes.cbiViewNameFields.ToString).CheckToolStripItem(mnu, Nothing, mvShowNameFields)
    mvMenuItems(CommandIndexes.cbiViewDescriptions.ToString).CheckToolStripItem(mnu, Nothing, mvShowCodeDescriptions)
    mvMenuItems(CommandIndexes.cbiViewIDNumbers.ToString).CheckToolStripItem(mnu, Nothing, mvShowIDNumbers)

    Dim vSteps As Boolean = dgrSteps.RowCount > 0
    mvMenuItems(CommandIndexes.cbiDeleteSteps.ToString).EnableToolStripItem(mnu, tsp, vSteps)
    mvMenuItems(CommandIndexes.cbiRebuild.ToString).EnableToolStripItem(mnu, tsp, vSteps)
    mvMenuItems(CommandIndexes.cbiSaveSteps.ToString).EnableToolStripItem(mnu, tsp, vSteps)
    If mvInitialised Then
      imbUp.Enabled = vSteps
      imbDown.Enabled = vSteps
      imbDelete.Enabled = vSteps
    End If
  End Sub
  Private Sub SetViewingSelection(ByVal pViewingSelection As Boolean)
    If pViewingSelection Then
      mvViewingSelection = True
    Else
      mvViewingSelection = False
    End If
    If mvViewsComboBox.SelectedIndex = -1 Then
      mvViewsComboBox.SelectedIndex = mvDefaultIndex
    Else
      ShowViewData(mvCurrentView, True)
    End If
  End Sub

  Private Sub dgr_ContactSelected(ByVal pSender As Object, ByVal pRow As Integer, ByVal pContactNumber As Integer) Handles dgr.ContactSelected
    Try
      FormHelper.ShowCardIndex(CareServices.XMLContactDataSelectionTypes.xcdtNone, pContactNumber)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub dgr_RowHeaderClicked(ByVal sender As Object, ByVal pRow As Integer) Handles dgr.RowHeaderClicked
    Dim vCol As Integer = dgr.GetColumn("contact_number")
    If vCol >= 0 Then
      dgr.ContextMenuStrip = mvBrowserMenu
    End If
    mvBrowserMenu.EntityType = HistoryEntityTypes.hetContacts
    mvBrowserMenu.ItemNumber = CInt(dgr.GetValue(pRow, vCol))
  End Sub
End Class
