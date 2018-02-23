Imports System.Runtime.InteropServices
Imports Advanced.LanguageExtensions

Public Class frmCardSet
  Inherits MaintenanceParentForm
  Implements IDashboardTabContainer

  Implements IPanelVisibility
  Implements IMainForm

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
  Friend WithEvents hdr As CDBNETCL.FormHeader
  Friend WithEvents sel As CDBNETCL.TabSelector
  Friend WithEvents dgr As CDBNETCL.DisplayGrid
  Friend WithEvents dpl As CDBNETCL.DisplayPanel
  Friend WithEvents splTop As System.Windows.Forms.SplitContainer
  Friend WithEvents splBottom As System.Windows.Forms.SplitContainer
  Friend WithEvents splRight As System.Windows.Forms.SplitContainer
  Friend WithEvents dgrMenuStrip As System.Windows.Forms.ContextMenuStrip
  Friend WithEvents dgrMenuNew As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents dgrMenuEdit As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents bpl As CDBNETCL.ButtonPanel
  Friend WithEvents cmdSave As System.Windows.Forms.Button
  Friend WithEvents cmdNew As System.Windows.Forms.Button
  Friend WithEvents cmdDelete As System.Windows.Forms.Button
  Friend WithEvents dts As CDBNETCL.DisplayTabSet
  Friend WithEvents dgrMenuNewDocument As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents dgr0MenuStrip As System.Windows.Forms.ContextMenuStrip
  Friend WithEvents dgr0MenuNew As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents dgr0MenuEdit As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents dgr1MenuStrip As System.Windows.Forms.ContextMenuStrip
  Friend WithEvents dgr1MenuNew As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents dgr1MenuEdit As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents dplMenuStrip As System.Windows.Forms.ContextMenuStrip
  Friend WithEvents dplMenuNew As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents dplMenuEdit As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents dplMenuCustomise As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents dplMenuRevert As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents dgr2MenuStrip As System.Windows.Forms.ContextMenuStrip
  Friend WithEvents dgr2MenuNew As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents dgr2MenuEdit As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents dgr3MenuStrip As System.Windows.Forms.ContextMenuStrip
  Friend WithEvents dgr3MenuNew As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents dgr3MenuEdit As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents dgr4MenuStrip As System.Windows.Forms.ContextMenuStrip
  Friend WithEvents dgr4MenuNew As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents dgr4MenuEdit As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents dgr5MenuStrip As System.Windows.Forms.ContextMenuStrip
  Friend WithEvents dgr5MenuNew As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents dgr5MenuEdit As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents dgr5MenuDelete As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents gtb As CDBNETCL.GridToolbar
  Friend WithEvents splTab As System.Windows.Forms.Splitter
  Friend WithEvents pnlDisplay As System.Windows.Forms.Panel
  <System.Diagnostics.DebuggerStepThrough()>
  Private Sub InitializeComponent()
    Me.components = New System.ComponentModel.Container()
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCardSet))
    Me.sel = New CDBNETCL.TabSelector()
    Me.dgr = New CDBNETCL.DisplayGrid()
    Me.pnlDisplay = New System.Windows.Forms.Panel()
    Me.dts = New CDBNETCL.DisplayTabSet()
    Me.splTab = New System.Windows.Forms.Splitter()
    Me.dpl = New CDBNETCL.DisplayPanel()
    Me.bpl = New CDBNETCL.ButtonPanel()
    Me.cmdSave = New System.Windows.Forms.Button()
    Me.cmdNew = New System.Windows.Forms.Button()
    Me.cmdDelete = New System.Windows.Forms.Button()
    Me.gtb = New CDBNETCL.GridToolbar()
    Me.splTop = New System.Windows.Forms.SplitContainer()
    Me.hdr = New CDBNETCL.FormHeader()
    Me.splBottom = New System.Windows.Forms.SplitContainer()
    Me.splRight = New System.Windows.Forms.SplitContainer()
    Me.dgrMenuStrip = New System.Windows.Forms.ContextMenuStrip(Me.components)
    Me.dgrMenuNew = New System.Windows.Forms.ToolStripMenuItem()
    Me.dgrMenuEdit = New System.Windows.Forms.ToolStripMenuItem()
    Me.dgrMenuNewDocument = New System.Windows.Forms.ToolStripMenuItem()
    Me.dgr0MenuStrip = New System.Windows.Forms.ContextMenuStrip(Me.components)
    Me.dgr0MenuNew = New System.Windows.Forms.ToolStripMenuItem()
    Me.dgr0MenuEdit = New System.Windows.Forms.ToolStripMenuItem()
    Me.dgr1MenuStrip = New System.Windows.Forms.ContextMenuStrip(Me.components)
    Me.dgr1MenuNew = New System.Windows.Forms.ToolStripMenuItem()
    Me.dgr1MenuEdit = New System.Windows.Forms.ToolStripMenuItem()
    Me.dplMenuStrip = New System.Windows.Forms.ContextMenuStrip(Me.components)
    Me.dplMenuNew = New System.Windows.Forms.ToolStripMenuItem()
    Me.dplMenuEdit = New System.Windows.Forms.ToolStripMenuItem()
    Me.dplMenuCustomise = New System.Windows.Forms.ToolStripMenuItem()
    Me.dplMenuRevert = New System.Windows.Forms.ToolStripMenuItem()
    Me.dgr2MenuStrip = New System.Windows.Forms.ContextMenuStrip(Me.components)
    Me.dgr2MenuNew = New System.Windows.Forms.ToolStripMenuItem()
    Me.dgr2MenuEdit = New System.Windows.Forms.ToolStripMenuItem()
    Me.dgr3MenuStrip = New System.Windows.Forms.ContextMenuStrip(Me.components)
    Me.dgr3MenuNew = New System.Windows.Forms.ToolStripMenuItem()
    Me.dgr3MenuEdit = New System.Windows.Forms.ToolStripMenuItem()
    Me.dgr4MenuStrip = New System.Windows.Forms.ContextMenuStrip(Me.components)
    Me.dgr4MenuNew = New System.Windows.Forms.ToolStripMenuItem()
    Me.dgr4MenuEdit = New System.Windows.Forms.ToolStripMenuItem()
    Me.dgr5MenuStrip = New System.Windows.Forms.ContextMenuStrip(Me.components)
    Me.dgr5MenuNew = New System.Windows.Forms.ToolStripMenuItem()
    Me.dgr5MenuEdit = New System.Windows.Forms.ToolStripMenuItem()
    Me.dgr5MenuDelete = New System.Windows.Forms.ToolStripMenuItem()
    Me.pnlDisplay.SuspendLayout()
    Me.bpl.SuspendLayout()
    CType(Me.splTop, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.splTop.Panel1.SuspendLayout()
    Me.splTop.Panel2.SuspendLayout()
    Me.splTop.SuspendLayout()
    CType(Me.splBottom, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.splBottom.Panel1.SuspendLayout()
    Me.splBottom.Panel2.SuspendLayout()
    Me.splBottom.SuspendLayout()
    CType(Me.splRight, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.splRight.Panel1.SuspendLayout()
    Me.splRight.Panel2.SuspendLayout()
    Me.splRight.SuspendLayout()
    Me.dgrMenuStrip.SuspendLayout()
    Me.dgr0MenuStrip.SuspendLayout()
    Me.dgr1MenuStrip.SuspendLayout()
    Me.dplMenuStrip.SuspendLayout()
    Me.dgr2MenuStrip.SuspendLayout()
    Me.dgr3MenuStrip.SuspendLayout()
    Me.dgr4MenuStrip.SuspendLayout()
    Me.dgr5MenuStrip.SuspendLayout()
    Me.SuspendLayout()
    '
    'sel
    '
    Me.sel.AccessibleName = "Selection Panel"
    Me.sel.BackColor = System.Drawing.Color.Transparent
    Me.sel.Dock = System.Windows.Forms.DockStyle.Fill
    Me.sel.Location = New System.Drawing.Point(0, 0)
    Me.sel.Name = "sel"
    Me.sel.Padding = New System.Windows.Forms.Padding(2, 18, 2, 18)
    Me.sel.Size = New System.Drawing.Size(170, 485)
    Me.sel.TabIndex = 0
    Me.sel.TreeContextMenu = Nothing
    '
    'dgr
    '
    Me.dgr.AccessibleDescription = "Display List"
    Me.dgr.AccessibleName = "Display List"
    Me.dgr.AccessibleRole = System.Windows.Forms.AccessibleRole.Table
    Me.dgr.ActiveColumn = 0
    Me.dgr.AllowColumnResize = True
    Me.dgr.AllowSorting = True
    Me.dgr.AutoSetHeight = False
    Me.dgr.AutoSetRowHeight = False
    Me.dgr.DisplayTitle = Nothing
    Me.dgr.Dock = System.Windows.Forms.DockStyle.Top
    Me.dgr.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!)
    Me.dgr.Location = New System.Drawing.Point(0, 0)
    Me.dgr.MaintenanceDesc = Nothing
    Me.dgr.MaxGridRows = 8
    Me.dgr.MultipleSelect = False
    Me.dgr.Name = "dgr"
    Me.dgr.Padding = New System.Windows.Forms.Padding(0, 0, 8, 0)
    Me.dgr.RowCount = 10
    Me.dgr.ShowIfEmpty = False
    Me.dgr.Size = New System.Drawing.Size(651, 183)
    Me.dgr.SuppressHyperLinkFormat = False
    Me.dgr.TabIndex = 0
    '
    'pnlDisplay
    '
    Me.pnlDisplay.BackColor = System.Drawing.SystemColors.Control
    Me.pnlDisplay.Controls.Add(Me.dts)
    Me.pnlDisplay.Controls.Add(Me.splTab)
    Me.pnlDisplay.Controls.Add(Me.dpl)
    Me.pnlDisplay.Controls.Add(Me.bpl)
    Me.pnlDisplay.Dock = System.Windows.Forms.DockStyle.Fill
    Me.pnlDisplay.Location = New System.Drawing.Point(0, 20)
    Me.pnlDisplay.MinimumSize = New System.Drawing.Size(20, 20)
    Me.pnlDisplay.Name = "pnlDisplay"
    Me.pnlDisplay.Padding = New System.Windows.Forms.Padding(0, 0, 8, 0)
    Me.pnlDisplay.Size = New System.Drawing.Size(651, 272)
    Me.pnlDisplay.TabIndex = 9
    '
    'dts
    '
    Me.dts.AccessibleName = "Detail Tabs"
    Me.dts.BackColor = System.Drawing.SystemColors.Control
    Me.dts.Dock = System.Windows.Forms.DockStyle.Fill
    Me.dts.Location = New System.Drawing.Point(0, 96)
    Me.dts.Margin = New System.Windows.Forms.Padding(2)
    Me.dts.MinimumSize = New System.Drawing.Size(20, 20)
    Me.dts.Name = "dts"
    Me.dts.PanelVisible = True
    Me.dts.Size = New System.Drawing.Size(643, 137)
    Me.dts.TabDock = System.Windows.Forms.DockStyle.None
    Me.dts.TabIndex = 1
    Me.dts.TabVisible = True
    '
    'splTab
    '
    Me.splTab.BackColor = System.Drawing.SystemColors.Control
    Me.splTab.Dock = System.Windows.Forms.DockStyle.Top
    Me.splTab.Location = New System.Drawing.Point(0, 88)
    Me.splTab.Name = "splTab"
    Me.splTab.Size = New System.Drawing.Size(643, 8)
    Me.splTab.TabIndex = 3
    Me.splTab.TabStop = False
    '
    'dpl
    '
    Me.dpl.AccessibleName = "Display Panel"
    Me.dpl.BackColor = System.Drawing.Color.Transparent
    Me.dpl.DataSelectionType = CDBNETCL.CareNetServices.XMLContactDataSelectionTypes.xcdtNone
    Me.dpl.Dock = System.Windows.Forms.DockStyle.Top
    Me.dpl.Location = New System.Drawing.Point(0, 0)
    Me.dpl.Name = "dpl"
    Me.dpl.Size = New System.Drawing.Size(643, 88)
    Me.dpl.TabIndex = 0
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdSave)
    Me.bpl.Controls.Add(Me.cmdNew)
    Me.bpl.Controls.Add(Me.cmdDelete)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bpl.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bpl.Location = New System.Drawing.Point(0, 233)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(643, 39)
    Me.bpl.TabIndex = 2
    Me.bpl.Visible = False
    '
    'cmdSave
    '
    Me.cmdSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdSave.Location = New System.Drawing.Point(162, 6)
    Me.cmdSave.Name = "cmdSave"
    Me.cmdSave.Size = New System.Drawing.Size(96, 27)
    Me.cmdSave.TabIndex = 0
    Me.cmdSave.Text = "&Save"
    '
    'cmdNew
    '
    Me.cmdNew.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdNew.Location = New System.Drawing.Point(273, 6)
    Me.cmdNew.Name = "cmdNew"
    Me.cmdNew.Size = New System.Drawing.Size(96, 27)
    Me.cmdNew.TabIndex = 1
    Me.cmdNew.Text = "&New"
    '
    'cmdDelete
    '
    Me.cmdDelete.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdDelete.Location = New System.Drawing.Point(384, 6)
    Me.cmdDelete.Name = "cmdDelete"
    Me.cmdDelete.Size = New System.Drawing.Size(96, 27)
    Me.cmdDelete.TabIndex = 2
    Me.cmdDelete.Text = "De&lete"
    '
    'gtb
    '
    Me.gtb.Dock = System.Windows.Forms.DockStyle.Top
    Me.gtb.Location = New System.Drawing.Point(0, 0)
    Me.gtb.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
    Me.gtb.Name = "gtb"
    Me.gtb.Size = New System.Drawing.Size(651, 20)
    Me.gtb.TabIndex = 3
    '
    'splTop
    '
    Me.splTop.Dock = System.Windows.Forms.DockStyle.Fill
    Me.splTop.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
    Me.splTop.IsSplitterFixed = True
    Me.splTop.Location = New System.Drawing.Point(0, 0)
    Me.splTop.Name = "splTop"
    Me.splTop.Orientation = System.Windows.Forms.Orientation.Horizontal
    '
    'splTop.Panel1
    '
    Me.splTop.Panel1.AccessibleName = "Splitter for Header"
    Me.splTop.Panel1.Controls.Add(Me.hdr)
    '
    'splTop.Panel2
    '
    Me.splTop.Panel2.Controls.Add(Me.splBottom)
    Me.splTop.Size = New System.Drawing.Size(829, 571)
    Me.splTop.SplitterDistance = 78
    Me.splTop.SplitterWidth = 8
    Me.splTop.TabIndex = 10
    '
    'hdr
    '
    Me.hdr.DisplayContextMenuStrip = False
    Me.hdr.Dock = System.Windows.Forms.DockStyle.Top
    Me.hdr.Location = New System.Drawing.Point(0, 0)
    Me.hdr.Name = "hdr"
    Me.hdr.Size = New System.Drawing.Size(829, 72)
    Me.hdr.TabIndex = 0
    '
    'splBottom
    '
    Me.splBottom.AccessibleDescription = "Splitter for Selection Panel"
    Me.splBottom.AccessibleName = ""
    Me.splBottom.AccessibleRole = System.Windows.Forms.AccessibleRole.None
    Me.splBottom.Dock = System.Windows.Forms.DockStyle.Fill
    Me.splBottom.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
    Me.splBottom.Location = New System.Drawing.Point(0, 0)
    Me.splBottom.Name = "splBottom"
    '
    'splBottom.Panel1
    '
    Me.splBottom.Panel1.Controls.Add(Me.sel)
    '
    'splBottom.Panel2
    '
    Me.splBottom.Panel2.Controls.Add(Me.splRight)
    Me.splBottom.Size = New System.Drawing.Size(829, 485)
    Me.splBottom.SplitterDistance = 170
    Me.splBottom.SplitterWidth = 8
    Me.splBottom.TabIndex = 11
    '
    'splRight
    '
    Me.splRight.AccessibleDescription = "Splitter for Display List"
    Me.splRight.AccessibleName = ""
    Me.splRight.Dock = System.Windows.Forms.DockStyle.Fill
    Me.splRight.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
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
    Me.splRight.Panel2.Controls.Add(Me.pnlDisplay)
    Me.splRight.Panel2.Controls.Add(Me.gtb)
    Me.splRight.Size = New System.Drawing.Size(651, 485)
    Me.splRight.SplitterDistance = 185
    Me.splRight.SplitterWidth = 8
    Me.splRight.TabIndex = 1
    '
    'dgrMenuStrip
    '
    Me.dgrMenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.dgrMenuNew, Me.dgrMenuEdit, Me.dgrMenuNewDocument})
    Me.dgrMenuStrip.Name = "dgrMenuStrip"
    Me.dgrMenuStrip.Size = New System.Drawing.Size(167, 70)
    '
    'dgrMenuNew
    '
    Me.dgrMenuNew.Name = "dgrMenuNew"
    Me.dgrMenuNew.Size = New System.Drawing.Size(166, 22)
    Me.dgrMenuNew.Text = "&New..."
    '
    'dgrMenuEdit
    '
    Me.dgrMenuEdit.Name = "dgrMenuEdit"
    Me.dgrMenuEdit.Size = New System.Drawing.Size(166, 22)
    Me.dgrMenuEdit.Text = "&Edit..."
    '
    'dgrMenuNewDocument
    '
    Me.dgrMenuNewDocument.Name = "dgrMenuNewDocument"
    Me.dgrMenuNewDocument.Size = New System.Drawing.Size(166, 22)
    Me.dgrMenuNewDocument.Text = "New &Document..."
    '
    'dgr0MenuStrip
    '
    Me.dgr0MenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.dgr0MenuNew, Me.dgr0MenuEdit})
    Me.dgr0MenuStrip.Name = "dgrMenuStrip"
    Me.dgr0MenuStrip.Size = New System.Drawing.Size(153, 70)
    '
    'dgr0MenuNew
    '
    Me.dgr0MenuNew.Name = "dgr0MenuNew"
    Me.dgr0MenuNew.Size = New System.Drawing.Size(152, 22)
    Me.dgr0MenuNew.Text = "&New..."
    '
    'dgr0MenuEdit
    '
    Me.dgr0MenuEdit.Name = "dgr0MenuEdit"
    Me.dgr0MenuEdit.Size = New System.Drawing.Size(152, 22)
    Me.dgr0MenuEdit.Text = "&Edit..."
    '
    'dgr1MenuStrip
    '
    Me.dgr1MenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.dgr1MenuNew, Me.dgr1MenuEdit})
    Me.dgr1MenuStrip.Name = "dgrMenuStrip"
    Me.dgr1MenuStrip.Size = New System.Drawing.Size(108, 48)
    '
    'dgr1MenuNew
    '
    Me.dgr1MenuNew.Name = "dgr1MenuNew"
    Me.dgr1MenuNew.Size = New System.Drawing.Size(107, 22)
    Me.dgr1MenuNew.Text = "&New..."
    '
    'dgr1MenuEdit
    '
    Me.dgr1MenuEdit.Name = "dgr1MenuEdit"
    Me.dgr1MenuEdit.Size = New System.Drawing.Size(107, 22)
    Me.dgr1MenuEdit.Text = "&Edit..."
    '
    'dplMenuStrip
    '
    Me.dplMenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.dplMenuNew, Me.dplMenuEdit, Me.dplMenuCustomise, Me.dplMenuRevert})
    Me.dplMenuStrip.Name = "dgrMenuStrip"
    Me.dplMenuStrip.Size = New System.Drawing.Size(131, 92)
    '
    'dplMenuNew
    '
    Me.dplMenuNew.Name = "dplMenuNew"
    Me.dplMenuNew.Size = New System.Drawing.Size(130, 22)
    Me.dplMenuNew.Text = "&New..."
    '
    'dplMenuEdit
    '
    Me.dplMenuEdit.Name = "dplMenuEdit"
    Me.dplMenuEdit.Size = New System.Drawing.Size(130, 22)
    Me.dplMenuEdit.Text = "&Edit..."
    '
    'dplMenuCustomise
    '
    Me.dplMenuCustomise.Name = "dplMenuCustomise"
    Me.dplMenuCustomise.Size = New System.Drawing.Size(130, 22)
    Me.dplMenuCustomise.Text = "Customise"
    '
    'dplMenuRevert
    '
    Me.dplMenuRevert.Name = "dplMenuRevert"
    Me.dplMenuRevert.Size = New System.Drawing.Size(130, 22)
    Me.dplMenuRevert.Text = "Revert"
    '
    'dgr2MenuStrip
    '
    Me.dgr2MenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.dgr2MenuNew, Me.dgr2MenuEdit})
    Me.dgr2MenuStrip.Name = "dgrMenuStrip"
    Me.dgr2MenuStrip.Size = New System.Drawing.Size(108, 48)
    '
    'dgr2MenuNew
    '
    Me.dgr2MenuNew.Name = "dgr2MenuNew"
    Me.dgr2MenuNew.Size = New System.Drawing.Size(107, 22)
    Me.dgr2MenuNew.Text = "&New..."
    '
    'dgr2MenuEdit
    '
    Me.dgr2MenuEdit.Name = "dgr2MenuEdit"
    Me.dgr2MenuEdit.Size = New System.Drawing.Size(107, 22)
    Me.dgr2MenuEdit.Text = "&Edit..."
    '
    'dgr3MenuStrip
    '
    Me.dgr3MenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.dgr3MenuNew, Me.dgr3MenuEdit})
    Me.dgr3MenuStrip.Name = "dgrMenuStrip"
    Me.dgr3MenuStrip.Size = New System.Drawing.Size(153, 70)
    '
    'dgr3MenuNew
    '
    Me.dgr3MenuNew.Name = "dgr3MenuNew"
    Me.dgr3MenuNew.Size = New System.Drawing.Size(152, 22)
    Me.dgr3MenuNew.Text = "&New..."
    '
    'dgr3MenuEdit
    '
    Me.dgr3MenuEdit.Name = "dgr3MenuEdit"
    Me.dgr3MenuEdit.Size = New System.Drawing.Size(152, 22)
    Me.dgr3MenuEdit.Text = "&Edit..."
    '
    'dgr4MenuStrip
    '
    Me.dgr4MenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.dgr4MenuNew, Me.dgr4MenuEdit})
    Me.dgr4MenuStrip.Name = "dgrMenuStrip"
    Me.dgr4MenuStrip.Size = New System.Drawing.Size(153, 70)
    '
    'dgr4MenuNew
    '
    Me.dgr4MenuNew.Name = "dgr4MenuNew"
    Me.dgr4MenuNew.Size = New System.Drawing.Size(152, 22)
    Me.dgr4MenuNew.Text = "&New..."
    '
    'dgr4MenuEdit
    '
    Me.dgr4MenuEdit.Name = "dgr4MenuEdit"
    Me.dgr4MenuEdit.Size = New System.Drawing.Size(152, 22)
    Me.dgr4MenuEdit.Text = "&Edit..."
    '
    'dgr5MenuStrip
    '
    Me.dgr5MenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.dgr5MenuNew, Me.dgr5MenuEdit, Me.dgr5MenuDelete})
    Me.dgr5MenuStrip.Name = "dgrMenuStrip"
    Me.dgr5MenuStrip.Size = New System.Drawing.Size(153, 70)
    '
    'dgr5MenuNew
    '
    Me.dgr5MenuNew.Name = "dgr5MenuNew"
    Me.dgr5MenuNew.Size = New System.Drawing.Size(152, 22)
    Me.dgr5MenuNew.Text = "&New..."
    '
    'dgr5MenuEdit
    '
    Me.dgr5MenuEdit.Name = "dgr5MenuEdit"
    Me.dgr5MenuEdit.Size = New System.Drawing.Size(152, 22)
    Me.dgr5MenuEdit.Text = "&Edit..."
    '
    'dgr5MenuDelete
    '
    Me.dgr5MenuDelete.Name = "dgr5MenuDelete"
    Me.dgr5MenuDelete.Size = New System.Drawing.Size(152, 22)
    Me.dgr5MenuDelete.Text = "&Delete..."
    '
    'frmCardSet
    '
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
    Me.AutoScroll = True
    Me.ClientSize = New System.Drawing.Size(829, 571)
    Me.Controls.Add(Me.splTop)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmCardSet"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.pnlDisplay.ResumeLayout(False)
    Me.bpl.ResumeLayout(False)
    Me.splTop.Panel1.ResumeLayout(False)
    Me.splTop.Panel2.ResumeLayout(False)
    CType(Me.splTop, System.ComponentModel.ISupportInitialize).EndInit()
    Me.splTop.ResumeLayout(False)
    Me.splBottom.Panel1.ResumeLayout(False)
    Me.splBottom.Panel2.ResumeLayout(False)
    CType(Me.splBottom, System.ComponentModel.ISupportInitialize).EndInit()
    Me.splBottom.ResumeLayout(False)
    Me.splRight.Panel1.ResumeLayout(False)
    Me.splRight.Panel2.ResumeLayout(False)
    CType(Me.splRight, System.ComponentModel.ISupportInitialize).EndInit()
    Me.splRight.ResumeLayout(False)
    Me.dgrMenuStrip.ResumeLayout(False)
    Me.dgr0MenuStrip.ResumeLayout(False)
    Me.dgr1MenuStrip.ResumeLayout(False)
    Me.dplMenuStrip.ResumeLayout(False)
    Me.dgr2MenuStrip.ResumeLayout(False)
    Me.dgr3MenuStrip.ResumeLayout(False)
    Me.dgr4MenuStrip.ResumeLayout(False)
    Me.dgr5MenuStrip.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub

#End Region

  Private mvContactInfo As ContactInfo
  Private mvDataType As CareServices.XMLContactDataSelectionTypes
  Private mvCustomFormNumber As Integer
  Private mvGroupID As String                   'Used for pages like Activity Groups
  Private mvReadOnlyPage As Boolean
  Private mvDataSet As DataSet
  Private mvDataSet2 As DataSet
  Private mvDataSet3 As DataSet
  Private mvDataSet4 As DataSet
  Private mvDataSet5 As DataSet
  Private mvDataSet6 As DataSet
  Private mvDataSet7 As DataSet
  Private mvGroupCode As String
  Private mvTabPages As Collection
  Private mvCustomPanel As EditPanel
  Private mvExamUnitSelector As ExamUnitSelector
  Private mvDashboard As DashboardTabControl
  Private mvCalendarView As CalendarView
  Private mvMarketingChart As ChartControl
  Private mvSelectedRow As Integer
  Private mvAllowUpdate As Boolean
  Private mvAllowInsert As Boolean
  Private mvInitialised As Boolean
  Private mvMeetingMenu As MeetingMenu
  Private mvMailingFilename As String
  Private mvMailingDocReqVal As Integer
  Private mvMailingDocOutgoing As Boolean
  Private mvMailingDocFulfilled As Boolean
  Private mvMailingNumber As Integer
  Private mvActiveChildControl As Control
  Private mvAccessibilityRoleReset As Boolean
  Private mvRightSplitterAccessibleRole As AccessibleRole
  Private mvGridAccessibleRole As AccessibleRole
  Private WithEvents mvActionMenu As ActionMenu
  Private WithEvents mvFinancialMenu As FinancialMenu
  Private WithEvents mvFundraisingPaymentMenu As FundraisingPaymentMenu
  Private WithEvents mvAnalysisFinancialMenu As FinancialAnalysisMenu
  Private WithEvents mvTransactionLinkMenu As TransactionLinkMenu
  Private WithEvents mvJournalLinkMenu As JournalLinkMenu
  Private WithEvents mvPurchaseOrderMenu As PurchaseOrderMenu
  Private WithEvents mvPurchaseOrderPaymentMenu As PurchaseOrderPaymentMenu
  Private WithEvents mvContactEventDelegateMenu As ContactEventDelegateMenu
  Private WithEvents mvViewMailingDocumentMenu As MailingDocumentMenu
  Private WithEvents mvCustomiseMenu As CustomiseMenu
  Private WithEvents mvExamsCustomiseMenu As CustomiseMenu
  Private WithEvents mvPurchaseInvoiceChequeMenu As PurchaseInvoiceChequeMenu
  Private WithEvents mvNetworkTreeview As NetworkTreeView
  Private WithEvents mvFinancialSubMenu As FinancialSubMenu
  Private WithEvents mvFrmNewMember As frmNewMember
  Private mvMainMenu As MainMenu
  Private WithEvents mvServiceBookingMenu As ServiceBookingMenu
  Private WithEvents mvDocumentMenu As BaseDocumentMenu
  Private mvCareWebBrowser As CareWebBrowser
  Private WithEvents mvPurchaseInvoiceMenu As PurchaseInvoiceMenu
  Private mvGridCurrentRow As Integer
  Private mvGridCurrentRowPayments As Integer
  Private WithEvents mvContactExamCertificateMenu As ContactExamCertificatesMenu
  Private mvDgr0IndependentItemNames As List(Of String) 'Collection of names of tool strip items that does not depend associated data
  Private mvDgr0ContextCommandHelper As ContextCommandHelper 'Datarow of associated data for current data tab set data grid(0) e.g. for meeting node it may be for associated comms log for a link
#Region "Maintenance Parent Form Methods"

  Public Overrides ReadOnly Property SizeMaintenanceForm() As Boolean
    Get
      Return True
    End Get
  End Property

  Public Overrides ReadOnly Property ContactInfo() As ContactInfo
    Get
      Return mvContactInfo
    End Get
  End Property

  Public Overrides ReadOnly Property ContactDataType() As CareServices.XMLContactDataSelectionTypes
    Get
      Return mvDataType
    End Get
  End Property

  Public Overrides Sub RefreshData()
    RefreshCard()
    UpdateHeader()
  End Sub

  Public Overrides Sub RefreshData(ByVal pType As CareServices.XMLMaintenanceControlTypes)
    Dim vUpdateHeader As Boolean

    Select Case pType
      Case CareServices.XMLMaintenanceControlTypes.xmctAddressUsage,
           CareServices.XMLMaintenanceControlTypes.xmctRole,
           CareServices.XMLMaintenanceControlTypes.xmctDocumentLink,
           CareServices.XMLMaintenanceControlTypes.xmctDocumentTopic,
           CareServices.XMLMaintenanceControlTypes.xmctContactAccounts,
           CareServices.XMLMaintenanceControlTypes.xmctNone
        If pType = CareServices.XMLMaintenanceControlTypes.xmctContactAccounts Then RefreshCard()
        ProcessRowSelection(dgr.CurrentRow, dgr.CurrentDataRow)
      Case CareServices.XMLMaintenanceControlTypes.xmctStickyNote
        If mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactStickyNotes Then RefreshCard()
        vUpdateHeader = True
      Case CareServices.XMLMaintenanceControlTypes.xmctContact, CareServices.XMLMaintenanceControlTypes.xmctOrganisation,
           CareServices.XMLMaintenanceControlTypes.xmctAddresses, CareServices.XMLMaintenanceControlTypes.xmctNumber,
           CareServices.XMLMaintenanceControlTypes.xmctAction, CareServices.XMLMaintenanceControlTypes.xmctActivities
        vUpdateHeader = True
        RefreshCard()
      Case Else
        RefreshCard()
    End Select
    If vUpdateHeader Then UpdateHeader()
  End Sub

  Private Sub UpdateHeader()
    Dim vDataSet As DataSet = DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactHeaderInformation, mvContactInfo.ContactNumber)
    hdr.Populate(vDataSet, mvContactInfo)
  End Sub

#End Region

#Region "IMainForm Interface"

  Public ReadOnly Property MainMenu() As MainMenu Implements IMainForm.MainMenu
    Get
      Return mvMainMenu
    End Get
  End Property

#End Region

  Private Sub InitialiseControls()
    Try
      Me.SuspendLayout()
      mvMainMenu = MainHelper.AddMainMenu(Me)
      SetControlTheme()
      Me.cmdSave.Text = ControlText.CmdSave
      Me.cmdNew.Text = ControlText.CmdNew
      Me.cmdDelete.Text = ControlText.CmdDelete
      Me.dgrMenuNew.Text = ControlText.MnuBrowserNew
      Me.dgrMenuEdit.Text = ControlText.MnuBrowserEdit
      Me.dgrMenuNewDocument.Text = ControlText.MnuNewDocument
      Me.dgr0MenuNew.Text = ControlText.MnuBrowserNew
      Me.dgr0MenuEdit.Text = ControlText.MnuBrowserEdit
      Me.dgr1MenuNew.Text = ControlText.MnuBrowserNew
      Me.dgr1MenuEdit.Text = ControlText.MnuBrowserEdit
      Me.dplMenuNew.Text = ControlText.MnuBrowserNew
      Me.dplMenuEdit.Text = ControlText.MnuBrowserEdit
      Me.dplMenuCustomise.Text = ControlText.MnuDisplayListCustomise
      Me.dplMenuRevert.Text = ControlText.MnuDisplayListRevert
      Me.dgr2MenuNew.Text = ControlText.MnuBrowserNew
      Me.dgr2MenuEdit.Text = ControlText.MnuBrowserEdit
      Me.dgr3MenuNew.Text = ControlText.MnuBrowserNew
      Me.dgr3MenuEdit.Text = ControlText.MnuBrowserEdit
      Me.dgr4MenuNew.Text = ControlText.MnuBrowserNew
      Me.dgr4MenuEdit.Text = ControlText.MnuBrowserEdit
      Me.dgr5MenuNew.Text = ControlText.MnuBrowserNew
      Me.dgr5MenuEdit.Text = ControlText.MnuBrowserEdit
      Me.dgr5MenuDelete.Text = ControlText.MnuBrowserDelete

      Dim vMenuNewItems() As ToolStripMenuItem = {dgrMenuNew, dgr0MenuNew, dgr1MenuNew, dgr2MenuNew, dgr3MenuNew, dgr4MenuNew, dgr5MenuNew, dplMenuNew}
      Dim vMenuEditItems() As ToolStripMenuItem = {dgrMenuEdit, dgr0MenuEdit, dgr1MenuEdit, dgr2MenuEdit, dgr3MenuEdit, dgr4MenuEdit, dgr5MenuEdit, dplMenuEdit}
      For Each vMenuItem As ToolStripMenuItem In vMenuNewItems
        vMenuItem.Text = ControlText.MnuBrowserNew
        vMenuItem.Image = AppHelper.ImageProvider.NewOtherImages16.Images("New")
      Next
      For Each vMenuItem As ToolStripMenuItem In vMenuEditItems
        vMenuItem.Text = ControlText.MnuBrowserEdit
        vMenuItem.Image = AppHelper.ImageProvider.NewOtherImages16.Images("Edit")
      Next
      Me.dgr5MenuDelete.Text = ControlText.MnuBrowserDelete
      Me.dgr5MenuDelete.Image = AppHelper.ImageProvider.NewOtherImages16.Images("Delete")

      dgrMenuNewDocument.Text = ControlText.MnuNewDocument
      Me.dts.TabDock = DockStyle.Fill
      SetPanelVisibility()
    Finally
      Me.ResumeLayout()
    End Try
  End Sub

#Region "IPanelVisibility"

  Public Sub SetPanelVisibility() Implements IPanelVisibility.SetPanelVisibility
    splTop.Panel1Collapsed = Not MainHelper.ShowHeaderPanel
    splBottom.Panel1Collapsed = Not MainHelper.ShowSelectionPanel
  End Sub

  Public Property PanelHasFocus() As Boolean Implements IPanelVisibility.PanelHasFocus
    Get
      Return sel.Controls(0).Focused
    End Get
    Set(ByVal value As Boolean)
      If value AndAlso splBottom.Panel1Collapsed = False Then
        sel.Focus()
      Else
        If dgr.Visible Then
          dgr.Focus()
        ElseIf dpl.Visible Then
          dpl.Focus()
        Else
          Me.Focus()
        End If
      End If
    End Set
  End Property

#End Region

#Region "IDashboardTabContainer"

  Private Sub OpenDashboard() Implements CDBNETCL.IDashboardTabContainer.Open
    Dim vDashboard As DashboardTabControl = mvDashboard.CreateFromDatabase(Me)
    If vDashboard IsNot Nothing Then
      Dim vParent As Control = mvDashboard.Parent
      mvDashboard.Visible = False
      vDashboard.Visible = False
      Me.SuspendLayout()
      vParent.Controls.Remove(mvDashboard)
      mvDashboard = vDashboard
      vParent.Controls.Add(mvDashboard)
      Dim vEntityGroup As EntityGroup = Nothing
      mvDashboard.SetItemID(mvContactInfo.ContactNumber)
      Me.ResumeLayout()
      mvDashboard.Visible = True
    End If
  End Sub

  Private Sub SaveDashboard(ByVal pOptions As DashboardTabControl.SaveOptions) Implements CDBNETCL.IDashboardTabContainer.Save
    mvDashboard.SaveToDatabase(pOptions)
  End Sub

  Private Sub DeleteDashboard(ByVal pOptions As DashboardTabControl.DeleteOptions) Implements CDBNETCL.IDashboardTabContainer.Delete
    mvDashboard.DeleteFromDatabase(pOptions)
    OpenDashboard()     'This will either open another existing Dashboard or init a new Dashboard
  End Sub

  Private Sub NavigateHistoryItem(ByVal pHistoryEntityType As HistoryEntityTypes, ByVal pNumber As Integer) Implements CDBNETCL.IDashboardTabContainer.NavigateHistoryItem
    MainHelper.NavigateHistoryItem(pHistoryEntityType, pNumber)
  End Sub

  Private Sub ProcessSearch(ByVal pSearchText As String) Implements CDBNETCL.IDashboardTabContainer.ProcessSearch
    MainHelper.ProcessSearch(pSearchText)
  End Sub

  Private Sub SetBrowserMenu(ByVal pSender As Object) Implements CDBNETCL.IDashboardTabContainer.SetBrowserMenu
    MainHelper.SetBrowserMenu(pSender, Me)
  End Sub

  Private Sub HistoryItemSelected(ByVal pSender As Object, ByVal pHistoryItem As CDBNETCL.UserHistoryItem, ByVal pDescription As String, ByVal pList As CDBNETCL.ArrayListEx) Implements CDBNETCL.IDashboardTabContainer.HistoryItemSelected
    MainHelper.HistoryItemSelected(pSender, pHistoryItem, pDescription, pList)
  End Sub

  Private Sub ProcessEditing(ByVal pType As CDBNETCL.DashboardDisplayPanel.MaintenanceTypes) Implements CDBNETCL.IDashboardTabContainer.ProcessEditing
    Dim vForm As frmCardMaintenance = Nothing
    Dim vDataType As CareServices.XMLContactDataSelectionTypes
    Dim vCursor As New BusyCursor
    Try
      Select Case pType
        Case DashboardDisplayPanel.MaintenanceTypes.Addresses
          vDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactAddresses
        Case DashboardDisplayPanel.MaintenanceTypes.Contact
          vDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactInformation
        Case DashboardDisplayPanel.MaintenanceTypes.Communications
          vDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactCommsNumbers
        Case Else
          Return
      End Select
      Dim vDataSet As DataSet = DataHelper.GetContactData(vDataType, mvContactInfo.ContactNumber)
      vForm = New frmCardMaintenance(Me, mvContactInfo, vDataType, vDataSet, True, 0)
      ShowMaintenanceForm(vForm)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Private Sub ActionItemSelected(ByVal pSender As Object) Implements CDBNETCL.IDashboardTabContainer.ActionItemSelected
    MainHelper.ActionItemSelected(pSender)
  End Sub

  Private Sub SetActionMenu(ByVal pSender As Object) Implements CDBNETCL.IDashboardTabContainer.SetActionMenu
    MainHelper.SetActionMenu(pSender, Me)
  End Sub

  Private Sub CalendardItemDoubleClicked(ByVal pType As CalendarView.CalendarItemTypes, ByVal pDescription As String, ByVal pStart As Date, ByVal pEnd As Date, ByVal pUniqueID As Integer) Implements IDashboardTabContainer.CalendardItemDoubleClicked
    MainHelper.CalendarDoubleClicked(Me, pType, pDescription, pStart, pEnd, pUniqueID)
  End Sub

  Private Sub ContactSelected(ByVal pSender As Object, ByVal pContactNumber As Integer) Implements IDashboardTabContainer.ContactSelected
    FormHelper.ShowContactCardIndex(pContactNumber)
  End Sub

  Public Sub NavigateNewSelectionSet(pContactNumbers As String) Implements CDBNETCL.IDashboardTabContainer.NavigateNewSelectionSet
    FormHelper.CreateNewSelectionSet(pContactNumbers)
  End Sub

#End Region

  Public Overrides Sub SetControlTheme()
    MyBase.SetControlTheme()
    pnlDisplay.BackColor = DisplayTheme.FormBackColor
    splBottom.BackColor = DisplayTheme.SplitterBackColor
    splBottom.SplitterWidth = DisplayTheme.SplitterWidth
    splRight.BackColor = DisplayTheme.SplitterBackColor
    splRight.SplitterWidth = DisplayTheme.SplitterWidth
  End Sub

  Public Overridable Sub Init(ByVal pDataSet As DataSet, ByVal pType As CareServices.XMLContactDataSelectionTypes, ByVal pContactInfo As ContactInfo, Optional ByVal pRetainPage As Boolean = False)
    Try
      Me.hdr.SuspendLayout()
      Me.sel.SuspendLayout()
      Me.dgr.SuspendLayout()
      Me.dpl.SuspendLayout()
      Me.splTop.SuspendLayout()
      Me.splBottom.Panel2.SuspendLayout()
      Me.SuspendLayout()
      For Each vControl As Control In Me.splBottom.Panel2.Controls
        If vControl.GetType.IsSubclassOf(GetType(Form)) Then
          DirectCast(vControl, Form).Close()
        End If
      Next vControl
      SettingsName = pContactInfo.ContactGroup & "_CardSet"
      sel.Init(pContactInfo)
      Dim vEntityGroup As EntityGroup
      Dim vPrefix As String = ""
      If DataHelper.ContactAndOrganisationGroups.ContainsKey(pContactInfo.ContactGroup) Then
        vEntityGroup = DataHelper.ContactAndOrganisationGroups(pContactInfo.ContactGroup)
        If vEntityGroup.ImageIndex < MainHelper.ImageProvider.NewTreeViewImages.Images.Count Then
          Me.Icon = Drawing.Icon.FromHandle(CType(MainHelper.ImageProvider.NewTreeViewImages.Images(vEntityGroup.ImageIndex), Bitmap).GetHicon)
          vPrefix = vEntityGroup.GroupName & ": "
        Else
          Me.Icon = vEntityGroup.Icon
          vPrefix = vEntityGroup.GroupName & ": "
        End If
      End If
      If mvDataType <> CareServices.XMLContactDataSelectionTypes.xcdtNone And pRetainPage Then
        'Leave mvdatatype as it is
      Else
        mvDataType = pType
        sel.SetSelectionType(mvDataType)
      End If
      mvContactInfo = pContactInfo
      If pDataSet IsNot Nothing Then hdr.Populate(pDataSet, mvContactInfo)
      splTop.SplitterDistance = hdr.Height

      splTop.SplitterWidth = 1
      splTop.BackColor = Color.Black
      Me.Text = vPrefix & mvContactInfo.ContactName
      dgr.AutoSetHeight = True
      dpl.AutoSetHeight = True
      mvInitialised = True
      hdr.DisplayContextMenuStrip = True
      RefreshCard()
      If mvContactInfo.ContactNumber > 0 Then
        UserHistory.AddContactHistoryNode(mvContactInfo.ContactNumber, mvContactInfo.ContactName, mvContactInfo.ContactGroup)
        MainHelper.SetStatusContact(mvContactInfo, True)
      End If
    Finally
      Me.hdr.ResumeLayout()
      Me.sel.ResumeLayout()
      Me.dgr.ResumeLayout()
      Me.dpl.ResumeLayout()
      Me.splTop.ResumeLayout()
      Me.splBottom.Panel2.ResumeLayout()
      Me.ResumeLayout()
    End Try
  End Sub

  Public ReadOnly Property CurrentRow() As Integer
    Get
      Return dgr.CurrentRow
    End Get
  End Property

  Public Sub RefreshCard()
    If Me.DesignMode Then Exit Sub
    Try
      Me.SuspendLayout()
      dgr.SuspendLayout()
      gtb.SuspendLayout()
      dpl.SuspendLayout()
      dts.SuspendLayout()
      bpl.SuspendLayout()
      dgr.ContextMenuStrip = Nothing
      dts.GridContextMenuStrip(0) = Nothing
      dts.GridContextMenuStrip(1) = Nothing
      dts.GridContextMenuStrip(2) = Nothing
      dpl.ContextMenuStrip = Nothing
      dgr.MaxGridRows = DisplayTheme.DefaultMaxGridRows
      dgrMenuNewDocument.Visible = False
      gtb.Visible = False
      Select Case mvDataType
        Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactNotes
          dgr.AutoSetRowHeight = True
          dpl.ShowAllText = True
        Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactStatusHistory
          dgr.AutoSetRowHeight = True
          dpl.ShowAllText = False
        Case Else
          dgr.AutoSetRowHeight = False
          dpl.ShowAllText = False
      End Select
      If dgr.MultipleSelect Then dgr.MultipleSelect = False
      Select Case mvDataType
        Case CareServices.XMLContactDataSelectionTypes.xcdtContactCustomFormData
          If mvCustomPanel Is Nothing Then
            mvCustomPanel = New EditPanel
            AddHandler mvCustomPanel.ContactSelected, AddressOf ContactSelectedHandler
            AddHandler mvCustomPanel.PopulateDisplayGrid, AddressOf PopulateDisplayGrid
            dts.Panel.Controls.Add(mvCustomPanel)
          Else
            mvCustomPanel.AutoScrollPosition = New Point(0, 0)
          End If
          dts.TabVisible = False
          bpl.Visible = IsEditableCustomForm(mvCustomFormNumber)
          mvCustomPanel.Visible = False
          mvCustomPanel.Init(New EditPanelInfo(CareServices.XMLMaintenanceControlTypes.xmctCustomForm, mvContactInfo, mvCustomFormNumber))
          mvCustomPanel.AutoSize = False
          dts.Height = mvCustomPanel.RequiredHeight
          mvCustomPanel.Height = dts.Height
          mvCustomPanel.Width = dts.Width
          mvCustomPanel.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
          mvCustomPanel.AutoScroll = True
        Case CareServices.XMLContactDataSelectionTypes.xcdtContactDocuments
          mvDocumentMenu = New DocumentMenu(Me)
          mvDocumentMenu.DocumentType = BaseDocumentMenu.DocumentTypes.ContactDocuments
          mvDocumentMenu.DocumentNumber = 0
        Case CType(CareNetServices.XMLContactDataSelectionTypes.xcdtContactMeetings, CareServices.XMLContactDataSelectionTypes)
          If mvMeetingMenu Is Nothing Then mvMeetingMenu = New MeetingMenu(Me)
          mvMeetingMenu.MeetingNumber = 0
        Case CareServices.XMLContactDataSelectionTypes.xcdtContactActions, CareServices.XMLContactDataSelectionTypes.xcdtContactLegacyActions
          mvActionMenu = New ActionMenu(Me)
          mvActionMenu.ContactInfo = mvContactInfo
          If mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactLegacyActions Then
            mvActionMenu.ActionType = ActionMenu.ActionTypes.LegacyActions
          Else
            mvActionMenu.ActionType = ActionMenu.ActionTypes.ContactActions
          End If
        Case CareServices.XMLContactDataSelectionTypes.xcdtContactCovenants, CareServices.XMLContactDataSelectionTypes.xcdtContactCreditCardAuthorities,
             CareServices.XMLContactDataSelectionTypes.xcdtContactDirectDebits, CareNetServices.XMLContactDataSelectionTypes.xcdtContactSubscriptions,
             CareServices.XMLContactDataSelectionTypes.xcdtContactEventBookings, CareServices.XMLContactDataSelectionTypes.xcdtContactGiftAidDeclarations,
             CareServices.XMLContactDataSelectionTypes.xcdtContactMembershipDetails, CareServices.XMLContactDataSelectionTypes.xcdtContactPaymentPlans,
             CareServices.XMLContactDataSelectionTypes.xcdtContactPledges, CareServices.XMLContactDataSelectionTypes.xcdtContactPostPGPledges,
             CareServices.XMLContactDataSelectionTypes.xcdtContactProcessedTransactions, CareServices.XMLContactDataSelectionTypes.xcdtContactStandingOrders,
             CareServices.XMLContactDataSelectionTypes.xcdtContactCreditCustomers, CareServices.XMLContactDataSelectionTypes.xcdtContactAppropriateCertificates,
             CareServices.XMLContactDataSelectionTypes.xcdtContactUnProcessedTransactions, CareServices.XMLContactDataSelectionTypes.xcdtContactCancelledProvisionalTrans,
             CareServices.XMLContactDataSelectionTypes.xcdtContactFundraising, CareNetServices.XMLContactDataSelectionTypes.xcdtContactDeliveryTransactions,
             CareNetServices.XMLContactDataSelectionTypes.xcdtContactSalesTransactions, CareNetServices.XMLContactDataSelectionTypes.xcdtContactEventRoomBookings,
             CareNetServices.XMLContactDataSelectionTypes.xcdtContactLoans, CareNetServices.XMLContactDataSelectionTypes.xcdtContactExamDetails,
             CareNetServices.XMLContactDataSelectionTypes.xcdtContactExamSummary,
             CareNetServices.XMLContactDataSelectionTypes.xcdtContactExamExemptions, CareNetServices.XMLContactDataSelectionTypes.xcdtContactDespatchNotes
          mvFinancialMenu = New FinancialMenu(Me, mvDataType, mvContactInfo)
          If mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactPaymentPlans Then
            mvFinancialSubMenu = New FinancialSubMenu(Me, mvDataType, mvContactInfo)
          End If
          Select Case mvDataType
            Case CareServices.XMLContactDataSelectionTypes.xcdtContactPaymentPlans,
                 CareServices.XMLContactDataSelectionTypes.xcdtContactPledges,
                 CareServices.XMLContactDataSelectionTypes.xcdtContactPostPGPledges,
                 CareServices.XMLContactDataSelectionTypes.xcdtContactMembershipDetails,
                 CareServices.XMLContactDataSelectionTypes.xcdtContactLoans
              mvTransactionLinkMenu = New TransactionLinkMenu(Me, mvDataType, mvContactInfo)
            Case CareServices.XMLContactDataSelectionTypes.xcdtContactProcessedTransactions
              mvAnalysisFinancialMenu = New FinancialAnalysisMenu(Me, mvDataType, mvContactInfo)
            Case CareServices.XMLContactDataSelectionTypes.xcdtContactFundraising
              mvTransactionLinkMenu = New TransactionLinkMenu(Me, mvDataType, mvContactInfo)
              mvFundraisingPaymentMenu = New FundraisingPaymentMenu(Me, mvDataType, mvContactInfo)
              mvActionMenu = New ActionMenu(Me)
              mvActionMenu.ContactInfo = mvContactInfo
              mvActionMenu.ActionType = ActionMenu.ActionTypes.FundraisingActions
          End Select
        Case CareServices.XMLContactDataSelectionTypes.xcdtContactXactionContributedTo,
             CareServices.XMLContactDataSelectionTypes.xcdtContactXactionHandledBy,
             CareServices.XMLContactDataSelectionTypes.xcdtContactXactionInMemoriamDonated,
             CareServices.XMLContactDataSelectionTypes.xcdtContactXactionInMemoriamReceived,
             CareServices.XMLContactDataSelectionTypes.xcdtContactXactionPaidInBy,
             CareServices.XMLContactDataSelectionTypes.xcdtContactXactionSentOnBehalfOf,
             CareServices.XMLContactDataSelectionTypes.xcdtContactLegacyBequests
          mvTransactionLinkMenu = New TransactionLinkMenu(Me, mvDataType, mvContactInfo)
        Case CareServices.XMLContactDataSelectionTypes.xcdtContactJournals
          mvJournalLinkMenu = New JournalLinkMenu(Me, mvDataType, mvContactInfo)
        Case CareServices.XMLContactDataSelectionTypes.xcdtContactPurchaseOrders
          mvPurchaseOrderMenu = New PurchaseOrderMenu(Me, mvDataType, mvContactInfo)
          mvPurchaseOrderPaymentMenu = New PurchaseOrderPaymentMenu(Me, mvDataType, mvContactInfo)
          mvGridCurrentRow = dgr.CurrentRow
          mvGridCurrentRowPayments = dts.DisplayGrid(1).CurrentRow
        Case CareServices.XMLContactDataSelectionTypes.xcdtContactEventDelegates
          mvContactEventDelegateMenu = New ContactEventDelegateMenu(Me, mvDataType, mvContactInfo)
        Case CareServices.XMLContactDataSelectionTypes.xcdtContactMailings
          mvViewMailingDocumentMenu = New MailingDocumentMenu(Me, mvDataType, mvContactInfo)
        Case CareServices.XMLContactDataSelectionTypes.xcdtContactPurchaseInvoices
          mvPurchaseInvoiceChequeMenu = New PurchaseInvoiceChequeMenu(Me, mvDataType, mvContactInfo)
          mvPurchaseInvoiceMenu = New PurchaseInvoiceMenu(Me, mvDataType, mvContactInfo)
        Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactServiceBookings
          mvServiceBookingMenu = New ServiceBookingMenu(Me)
          mvTransactionLinkMenu = New TransactionLinkMenu(Me, mvDataType, mvContactInfo)
        Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactExamCertificates
          mvContactExamCertificateMenu = New ContactExamCertificatesMenu()
          AddHandler mvContactExamCertificateMenu.ReprintSelected, Sub(sender As Object, e As ReprintSelectedEventArgs)
                                                                     Call New frmCertificateReprint(IntegerValue(dgr.GetValue(dgr.CurrentDataRow, "ContactExamCertId")), e.ReprintType).ShowDialog()
                                                                     ProcessRowSelection(dgr.CurrentRow, dgr.CurrentDataRow)
                                                                   End Sub
          AddHandler mvContactExamCertificateMenu.RecallRequested, Sub(sender As Object, e As EventArgs)
                                                                     ExamsDataHelper.RecallOrReinstateCertificate(IntegerValue(dgr.GetValue(dgr.CurrentDataRow, "ContactExamCertId")),
                                                                                                                  dgr.GetValue(dgr.CurrentDataRow, "IsCertificateRecalled").Equals("N", StringComparison.InvariantCultureIgnoreCase))
                                                                     Dim currentRow As Integer = dgr.CurrentRow
                                                                     RefreshData()
                                                                     dgr.SelectRow(currentRow)
                                                                     ProcessRowSelection(dgr.CurrentRow, dgr.CurrentDataRow)
                                                                   End Sub
          dgr.ContextMenuStrip = mvContactExamCertificateMenu
      End Select
      If mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactCustomFormData AndAlso mvCustomPanel.Controls.Count > 0 Then
        dts.PanelVisible = True
        'mvCustomPanel.Visible = True
      Else
        If mvCustomPanel IsNot Nothing Then mvCustomPanel.Visible = False
        bpl.Visible = False     'Only visible on custom forms
      End If

      If (mvContactInfo.OwnershipAccessLevel = ContactInfo.OwnershipAccessLevels.oalWrite AndAlso
          Not mvReadOnlyPage AndAlso
          DataHelper.UserInfo.AccessLevel > UserInfo.UserAccessLevel.ualReadOnly) OrElse
         mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactDocuments OrElse
         mvDataType = CareNetServices.XMLContactDataSelectionTypes.xcdtContactExamCertificates OrElse
         (mvDataType = CareNetServices.XMLContactDataSelectionTypes.xcdtContactInformation AndAlso
          mvContactInfo.Department = DataHelper.UserInfo.Department) Then
        If GetMaintenanceType(mvDataType) <> CareServices.XMLMaintenanceControlTypes.xmctNone Then
          dgr.ContextMenuStrip = dgrMenuStrip
          dgrMenuNew.Visible = True
          dgrMenuEdit.Visible = True
          dplMenuNew.Visible = True
          dplMenuEdit.Visible = True

          If AppValues.HasItemAccessRights(AppValues.AccessControlItems.aciDisplayListMaintenance) AndAlso mvDataType = CareNetServices.XMLContactDataSelectionTypes.xcdtContactInformation Then
            dplMenuCustomise.Visible = True
            dplMenuRevert.Visible = True
          Else
            dplMenuCustomise.Visible = False
            dplMenuRevert.Visible = False
          End If

          Select Case mvDataType
            Case CareServices.XMLContactDataSelectionTypes.xcdtContactActions, CareServices.XMLContactDataSelectionTypes.xcdtContactLegacyActions
              If mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactActions OrElse mvContactInfo.LegacyNumber > 0 Then
                dgr.ContextMenuStrip = mvActionMenu
                mvActionMenu.ActionNumber = 0
                dts.GridContextMenuStrip(0) = dgr0MenuStrip
                dts.GridContextMenuStrip(1) = dgr1MenuStrip
              Else
                dgr.ContextMenuStrip = Nothing
              End If
            Case CareServices.XMLContactDataSelectionTypes.xcdtContactAddresses
              dts.GridContextMenuStrip(1) = dgr1MenuStrip
            Case CareServices.XMLContactDataSelectionTypes.xcdtContactCommsNumbers
              dts.GridContextMenuStrip(0) = dgr0MenuStrip
            Case CareServices.XMLContactDataSelectionTypes.xcdtContactCreditCustomers
              If AppValues.HasItemAccessRights(AppValues.AccessControlItems.aciNewCreditCustomer) Then dgrMenuNew.Visible = True Else dgrMenuNew.Visible = False
              dts.GridContextMenuStrip(0) = mvFinancialMenu
            Case CareServices.XMLContactDataSelectionTypes.xcdtContactDocuments
              dgr.ContextMenuStrip = mvDocumentMenu
              dts.GridContextMenuStrip(0) = dgr0MenuStrip
              dts.GridContextMenuStrip(1) = dgr1MenuStrip
            Case CType(CareNetServices.XMLContactDataSelectionTypes.xcdtContactMeetings, CareServices.XMLContactDataSelectionTypes)
              dgr.ContextMenuStrip = mvMeetingMenu
              dts.GridContextMenuStrip(0) = dgr0MenuStrip
              dts.GridContextMenuStrip(1) = dgr1MenuStrip
            Case CareServices.XMLContactDataSelectionTypes.xcdtContactInformation
              dpl.ContextMenuStrip = dplMenuStrip
              dplMenuNew.Visible = False
              gtb.Visible = gtb.InitFromContextMenuStrip(dplMenuStrip) > 0
            Case CareServices.XMLContactDataSelectionTypes.xcdtContactLegacy
              dpl.ContextMenuStrip = dplMenuStrip
            Case CareServices.XMLContactDataSelectionTypes.xcdtContactLegacyAssets, CareServices.XMLContactDataSelectionTypes.xcdtContactLegacyBequests,
                 CareServices.XMLContactDataSelectionTypes.xcdtContactLegacyExpenses, CareServices.XMLContactDataSelectionTypes.xcdtContactLegacyLinks,
                 CareServices.XMLContactDataSelectionTypes.xcdtContactLegacyTaxCertificates
              If mvContactInfo.LegacyNumber <= 0 Then
                dgrMenuNew.Visible = False
                dgrMenuEdit.Visible = False
              Else
                If mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactLegacyBequests Then
                  dts.GridContextMenuStrip(0) = dgr0MenuStrip
                  dgr0MenuNew.Visible = False
                  dts.GridContextMenuStrip(1) = dgr1MenuStrip
                End If
              End If
            Case CareServices.XMLContactDataSelectionTypes.xcdtContactNotes, CareServices.XMLContactDataSelectionTypes.xcdtContactAppropriateCertificates
              dgrMenuNew.Visible = False
            Case CareServices.XMLContactDataSelectionTypes.xcdtContactPositions, CareNetServices.XMLContactDataSelectionTypes.xcdtContactViewOrganisations
              dts.GridContextMenuStrip(0) = dgr0MenuStrip
              dts.GridContextMenuStrip(1) = dgr1MenuStrip
              dts.GridContextMenuStrip(2) = dgr2MenuStrip
              dts.GridContextMenuStrip(3) = dgr3MenuStrip
              dts.GridContextMenuStrip(4) = dgr4MenuStrip
              dts.GridContextMenuStrip(5) = dgr5MenuStrip
            Case CareServices.XMLContactDataSelectionTypes.xcdtContactGiftAidDeclarations
              If mvContactInfo.ContactType = CDBNETCL.ContactInfo.ContactTypes.ctOrganisation Then
                mvFinancialMenu.Items(FinancialMenu.FinancialMenuItems.fmiNew).Enabled = False
              End If
            Case CareServices.XMLContactDataSelectionTypes.xcdtContactRegisteredUsers
              dgrMenuNew.Visible = True
            Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactCreditCards
              dgrMenuNew.Visible = False
          End Select
        End If
      End If

      Select Case mvDataType
        Case CareServices.XMLContactDataSelectionTypes.xcdtContactCovenants, CareServices.XMLContactDataSelectionTypes.xcdtContactCreditCardAuthorities,
             CareServices.XMLContactDataSelectionTypes.xcdtContactDirectDebits, CareNetServices.XMLContactDataSelectionTypes.xcdtContactSubscriptions,
             CareServices.XMLContactDataSelectionTypes.xcdtContactEventBookings, CareServices.XMLContactDataSelectionTypes.xcdtContactGiftAidDeclarations,
             CareServices.XMLContactDataSelectionTypes.xcdtContactMembershipDetails, CareServices.XMLContactDataSelectionTypes.xcdtContactPaymentPlans,
             CareServices.XMLContactDataSelectionTypes.xcdtContactPledges, CareServices.XMLContactDataSelectionTypes.xcdtContactPostPGPledges,
             CareServices.XMLContactDataSelectionTypes.xcdtContactProcessedTransactions, CareServices.XMLContactDataSelectionTypes.xcdtContactStandingOrders,
             CareServices.XMLContactDataSelectionTypes.xcdtContactAppropriateCertificates, CareServices.XMLContactDataSelectionTypes.xcdtContactUnProcessedTransactions,
             CareServices.XMLContactDataSelectionTypes.xcdtContactCancelledProvisionalTrans, CareServices.XMLContactDataSelectionTypes.xcdtContactFundraising,
             CareNetServices.XMLContactDataSelectionTypes.xcdtContactDeliveryTransactions, CareNetServices.XMLContactDataSelectionTypes.xcdtContactSalesTransactions,
             CareNetServices.XMLContactDataSelectionTypes.xcdtContactEventRoomBookings, CareNetServices.XMLContactDataSelectionTypes.xcdtContactLoans,
             CareNetServices.XMLContactDataSelectionTypes.xcdtContactExamDetails, CareNetServices.XMLContactDataSelectionTypes.xcdtContactExamExemptions,
             CareNetServices.XMLContactDataSelectionTypes.xcdtContactExamSummary,
             CareNetServices.XMLContactDataSelectionTypes.xcdtContactDespatchNotes
          dgr.ContextMenuStrip = mvFinancialMenu
          Select Case mvDataType
            Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactLoans
              dts.GridContextMenuStrip(1) = mvTransactionLinkMenu
            Case CareServices.XMLContactDataSelectionTypes.xcdtContactMembershipDetails
              dts.GridContextMenuStrip(1) = mvTransactionLinkMenu
            Case CareServices.XMLContactDataSelectionTypes.xcdtContactPaymentPlans
              dts.GridContextMenuStrip(1) = mvTransactionLinkMenu
              dts.GridContextMenuStrip(2) = mvFinancialSubMenu
            Case CareServices.XMLContactDataSelectionTypes.xcdtContactPledges, CareServices.XMLContactDataSelectionTypes.xcdtContactPostPGPledges
              dts.GridContextMenuStrip(0) = mvTransactionLinkMenu
            Case CareServices.XMLContactDataSelectionTypes.xcdtContactProcessedTransactions
              dts.GridContextMenuStrip(0) = mvAnalysisFinancialMenu
            Case CareServices.XMLContactDataSelectionTypes.xcdtContactFundraising
              dts.GridContextMenuStrip(0) = mvFundraisingPaymentMenu
              dts.GridContextMenuStrip(4) = mvActionMenu
          End Select
        Case CareServices.XMLContactDataSelectionTypes.xcdtContactXactionContributedTo,
             CareServices.XMLContactDataSelectionTypes.xcdtContactXactionHandledBy,
             CareServices.XMLContactDataSelectionTypes.xcdtContactXactionInMemoriamDonated,
             CareServices.XMLContactDataSelectionTypes.xcdtContactXactionInMemoriamReceived,
             CareServices.XMLContactDataSelectionTypes.xcdtContactXactionPaidInBy,
             CareServices.XMLContactDataSelectionTypes.xcdtContactXactionSentOnBehalfOf
          dgr.ContextMenuStrip = mvTransactionLinkMenu
        Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactLegacyBequests
          dts.GridContextMenuStrip(0) = mvTransactionLinkMenu

        Case CareServices.XMLContactDataSelectionTypes.xcdtContactJournals
          dgr.ContextMenuStrip = mvJournalLinkMenu
        Case CareServices.XMLContactDataSelectionTypes.xcdtContactPurchaseOrders
          dgr.ContextMenuStrip = mvPurchaseOrderMenu
          dts.GridContextMenuStrip(1) = mvPurchaseOrderPaymentMenu
        Case CareServices.XMLContactDataSelectionTypes.xcdtContactEventDelegates
          dgr.ContextMenuStrip = mvContactEventDelegateMenu
        Case CareServices.XMLContactDataSelectionTypes.xcdtContactMailings
          dgr.ContextMenuStrip = mvViewMailingDocumentMenu
        Case CareServices.XMLContactDataSelectionTypes.xcdtContactPurchaseInvoices
          dts.GridContextMenuStrip(1) = mvPurchaseInvoiceChequeMenu
          dgr.ContextMenuStrip = mvPurchaseInvoiceMenu
        Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactServiceBookings
          dgr.ContextMenuStrip = mvServiceBookingMenu
          dts.GridContextMenuStrip(0) = mvTransactionLinkMenu
      End Select

      Select Case mvDataType
        Case CareServices.XMLContactDataSelectionTypes.xcdtContactCovenants,
             CareServices.XMLContactDataSelectionTypes.xcdtContactCreditCardAuthorities,
             CareServices.XMLContactDataSelectionTypes.xcdtContactDirectDebits,
             CareServices.XMLContactDataSelectionTypes.xcdtContactMembershipDetails,
             CareServices.XMLContactDataSelectionTypes.xcdtContactPaymentPlans,
             CareServices.XMLContactDataSelectionTypes.xcdtContactPledges,
             CareServices.XMLContactDataSelectionTypes.xcdtContactPostPGPledges,
             CareServices.XMLContactDataSelectionTypes.xcdtContactStandingOrders,
             CareServices.XMLContactDataSelectionTypes.xcdtContactLoans
          dgr.MaxGridRows = DisplayTheme.CommittmentMaxGridRows
      End Select

      If mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactInformation OrElse mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactLegacy Then
        HideAdditionalControls(Nothing)
        HideDisplayGridPanel(True)
        mvDataSet = DataHelper.GetContactData(mvDataType, mvContactInfo.ContactNumber)
      ElseIf mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactAppointments Then
        HideAdditionalControls(mvCalendarView)
        HideDisplayGridPanel(True)
        If mvCalendarView Is Nothing Then
          mvCalendarView = New CalendarView
          AddHandler mvCalendarView.ItemDoubleClicked, AddressOf CalendarDoubleClickedHandler
          dts.Panel.Controls.Add(mvCalendarView)
          mvCalendarView.Dock = DockStyle.Fill
        End If
        dts.TabVisible = False
        mvCalendarView.Visible = True
        dts.PanelVisible = True
        mvDataSet = New DataSet ' DataHelper.GetContactData(mvDataType, mvContactInfo.ContactNumber)
        mvCalendarView.Init(mvContactInfo)
      ElseIf mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactDashboard Then
        HideAdditionalControls(mvDashboard)
        If mvDashboard Is Nothing Then
          mvDashboard = New DashboardTabControl
          mvDashboard.Visible = False
          dts.Panel.Controls.Add(mvDashboard)
          mvDashboard.Dock = DockStyle.Fill
          Dim vDashboardType As DashboardTypes = DashboardTypes.ContactDashboardType
          If mvContactInfo.ContactType = CDBNETCL.ContactInfo.ContactTypes.ctOrganisation Then vDashboardType = DashboardTypes.OrganisationDashboardType
          mvDashboard.Init(Me, vDashboardType, mvContactInfo.ContactGroup)
          OpenDashboard()
        Else
          mvDashboard.Visible = False
        End If
        HideDisplayGridPanel(True)
        dts.TabVisible = False
        dts.PanelVisible = True
        mvDataSet = New DataSet
        mvDashboard.SetItemID(mvContactInfo.ContactNumber)
        mvDashboard.Visible = True

        ' New Enum for Network Contact
      ElseIf mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactPositionLinks + 2 Then ' TODO: The new enum xcdtContactNetwork is not implemented in vb code
        HideAdditionalControls(mvNetworkTreeview)

        If dts.Panel.Controls.Contains(mvNetworkTreeview) Then
          dts.Panel.Controls.Remove(mvNetworkTreeview)
        End If
        mvNetworkTreeview = New NetworkTreeView(mvContactInfo, CareServices.XMLContactDataSelectionTypes.xcdtContactLinksTo)
        mvNetworkTreeview.Visible = False
        dts.Panel.Controls.Add(mvNetworkTreeview)
        mvNetworkTreeview.Dock = DockStyle.Fill
        mvNetworkTreeview.SetBrowserMenu()

        HideDisplayGridPanel(True)
        dts.TabVisible = False
        dts.PanelVisible = True
        mvDataSet = New DataSet
        mvNetworkTreeview.Visible = True

      ElseIf mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactScores Then
        HideAdditionalControls(mvMarketingChart)
        dgr.MultipleSelect = True
        mvDataSet = dgr.Populate(mvDataType, mvContactInfo.ContactNumber)
        HideDisplayGridPanel(False)
        If mvMarketingChart Is Nothing Then
          mvMarketingChart = New ChartControl
          dts.Panel.Controls.Add(mvMarketingChart)
          mvMarketingChart.Dock = DockStyle.Fill
        End If
        dts.TabVisible = False
        dts.PanelVisible = True
        mvMarketingChart.Visible = True
        Dim vSequences As String = AppValues.ConfigurationValue(AppValues.ConfigurationValues.profile_user_sequence)
        If vSequences.Length = 0 Then vSequences = "1|2|3|4"
        Dim vSequenceValues() As String = vSequences.Split("|"c)
        For Each vSequence As String In vSequenceValues
          Dim vIndex As Integer = IntegerValue(vSequence)
          dgr.AddToSelection(dgr.GetColumn("Sequence"), vIndex.ToString)
        Next
        mvMarketingChart.PopulateFromScores(mvDataSet, True)
      ElseIf mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactExamSummary OrElse mvDataType = CareNetServices.XMLContactDataSelectionTypes.xcdtContactExamDetails Then
        HideAdditionalControls(mvExamUnitSelector)
        mvDataSet = dgr.Populate(mvDataType, mvContactInfo.ContactNumber)
        If mvExamUnitSelector Is Nothing Then
          mvExamUnitSelector = New ExamUnitSelector
          dts.Panel.Controls.Add(mvExamUnitSelector)
          mvExamUnitSelector.Dock = DockStyle.Fill
          AddHandler mvExamUnitSelector.CustomiseCardSet, AddressOf dgr_CustomiseCardSet
          mvExamsCustomiseMenu = New CustomiseMenu
          mvExamUnitSelector.DisplayPanelContextMenu = mvExamsCustomiseMenu
        End If
        mvExamsCustomiseMenu.SetContext(mvDataType, mvContactInfo)
        dts.TabVisible = False
        dts.PanelVisible = True
        mvExamUnitSelector.SetContext(mvDataType, mvContactInfo)
        mvExamUnitSelector.Visible = True
        HideDisplayGridPanel(False)
      Else
        HideAdditionalControls(Nothing)
        If mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactCustomFormData Then
          Dim vList As New ParameterList(True)
          vList.IntegerValue("CustomForm") = mvCustomFormNumber
          mvDataSet = dgr.Populate(mvDataType, mvContactInfo.ContactNumber, vList)
          cmdDelete.Visible = False
          cmdNew.Visible = False
          cmdSave.Visible = False
          Dim vTable As DataTable = Nothing
          Dim vHideGrid As Boolean
          If mvDataSet.Tables.Contains("DataRow") AndAlso mvDataSet.Tables("DataRow").Columns.Contains("CustomFormUrl") Then
            'J1414: Display Web Browser instead of display panel
            Dim vDT As DataTable = mvDataSet.Tables("DataRow")
            Dim vCustomFormUrl As String = ""
            Dim vShowBrowserToolbar As Boolean = False
            If vDT IsNot Nothing Then
              For Each vRow As DataRow In vDT.Rows
                vCustomFormUrl = vRow.Item("CustomFormUrl").ToString
                vShowBrowserToolbar = BooleanValue(vRow.Item("ShowBrowserToolbar").ToString)
              Next
            End If
            mvCareWebBrowser = New CareWebBrowser
            mvCareWebBrowser.Init(vCustomFormUrl, vShowBrowserToolbar)
            mvCareWebBrowser.Dock = DockStyle.Fill
            'J1414: Clear any cached controls
            mvCustomPanel.Controls.Clear()
            mvCustomPanel.Controls.Add(mvCareWebBrowser)
            vHideGrid = True
            dts.PanelVisible = True
          ElseIf mvDataSet.Tables.Contains("Column") Then
            vHideGrid = False   'We have columns so show the grid
            vTable = mvDataSet.Tables("Column")
          Else
            vHideGrid = True    'There is no definition for the columns of this custom form which means there should not be a grid (grid_attribute_names null)
            vTable = mvDataSet.Tables("CustomFormInfo")
          End If
          HideDisplayGridPanel(vHideGrid)
          If vTable IsNot Nothing Then
            For Each vRow As DataRow In vTable.Rows
              If vRow.Item("Name").ToString = "AllowInsert" Then
                mvAllowInsert = BooleanValue(vRow.Item("Value").ToString) AndAlso vHideGrid = False
                cmdNew.Visible = mvAllowInsert
              ElseIf vRow.Item("Name").ToString = "AllowDelete" Then
                cmdDelete.Visible = BooleanValue(vRow.Item("Value").ToString) AndAlso vHideGrid = False
              ElseIf vRow.Item("Name").ToString = "AllowUpdate" Then
                mvAllowUpdate = BooleanValue(vRow.Item("Value").ToString)
              End If
            Next
          End If
          If mvContactInfo.OwnershipAccessLevel <> ContactInfo.OwnershipAccessLevels.oalWrite OrElse mvReadOnlyPage OrElse DataHelper.UserInfo.AccessLevel <= UserInfo.UserAccessLevel.ualReadOnly Then
            cmdNew.Visible = False
            cmdDelete.Visible = False
            cmdSave.Visible = False
          Else
            cmdSave.Visible = mvAllowInsert OrElse mvAllowUpdate
            cmdSave.Enabled = mvAllowUpdate
          End If
          If mvCustomPanel.Controls.Count > 0 Then mvCustomPanel.Visible = True
          bpl.RepositionButtons()
        ElseIf mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactCategories AndAlso mvGroupID.Length > 0 Then
          mvDataSet = DataHelper.GetContactActivityGroupData(mvContactInfo, mvGroupID)
          dgr.Populate(mvDataSet)
          HideDisplayGridPanel(False)
        ElseIf mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactLinksTo AndAlso mvGroupID.Length > 0 Then
          mvDataSet = DataHelper.GetContactRelationshipGroupData(mvContactInfo, mvGroupID)
          dgr.Populate(mvDataSet)
          HideDisplayGridPanel(False)
        ElseIf mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactMembershipDetails Then
          mvDataSet = dgr.Populate(mvDataType, mvContactInfo.ContactNumber)
          HideDisplayGridPanel(False)
          Dim vTabCount As Integer = 5
          If (AppValues.ConfigurationOption(AppValues.ConfigurationOptions.fp_pay_plan_amendment_history) AndAlso AppValues.ConfigurationOption(AppValues.ConfigurationOptions.option_journal, True)) Then vTabCount = 6
          If (BooleanValue(SelectRowItem(0, "BranchMembership")) = True AndAlso AppValues.ControlValue(AppValues.ControlTables.membership_controls, AppValues.ControlValues.organisation_group).Length > 0) Then
            Dim vDGR As DisplayGrid = dts.DisplayGrid(5)
            If vTabCount = 5 Then vDGR = dts.DisplayGrid(4)
            'Populate MembershipGroups
            vDGR.Populate(DataHelper.GetMembershipData(CareServices.XMLMembershipDataSelectionTypes.xmdtMembershipGroups, SelectRowItemNumber(0, "MembershipNumber")))
            vDGR.ShowIfEmpty = True
            vTabCount += 1
            vDGR = dts.DisplayGrid(6)
            If vTabCount = 6 Then vDGR = dts.DisplayGrid(5)
            'Populate MembershipHistoryDetails
            vDGR.Populate(DataHelper.GetMembershipData(CareServices.XMLMembershipDataSelectionTypes.xmdtMembershipGroupHistory, SelectRowItemNumber(0, "MembershipNumber")))
            vDGR.ShowIfEmpty = True
            vTabCount += 1
          End If
        Else
          Dim vParams As ParameterList = Nothing
          If mvDataType = CareNetServices.XMLContactDataSelectionTypes.xcdtContactDocuments Then
            vParams = New ParameterList(True)
            vParams.Add("IncludeEmailDocSource", "Y")
          ElseIf mvDataType = CareNetServices.XMLContactDataSelectionTypes.xcdtContactViewOrganisations Then
            vParams = New ParameterList(True)
            vParams.Add("OrganisationGroup", mvGroupID)
          End If
          mvDataSet = dgr.Populate(mvDataType, mvContactInfo.ContactNumber, vParams)
          If dgr.DataRowCount = 0 AndAlso mvDataType <> CareServices.XMLContactDataSelectionTypes.xcdtContactNotes Then
            dgrMenuEdit.Visible = False
          End If
          'Set MenuNew invisible if Contact is already registered
          If dgr.DataRowCount > 0 AndAlso mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactRegisteredUsers Then
            dgrMenuNew.Visible = False
          End If
          HideDisplayGridPanel(False)
        End If
        If dgr.DataRowCount = 0 AndAlso mvDataType <> CareServices.XMLContactDataSelectionTypes.xcdtContactNotes Then
          dgrMenuEdit.Visible = False
        End If
      End If
      dpl.Init(mvDataSet, False, False)
      If mvDataType <> CareServices.XMLContactDataSelectionTypes.xcdtContactScores Then ProcessRowSelection(0, 0)
      Me.ResumeLayout()
      splBottom.Panel2.AutoScroll = True
      dpl.ProcessResize = True
      dts.TabDock = DockStyle.Fill
    Finally
      bpl.ResumeLayout()
      dts.ResumeLayout()
      dpl.ResumeLayout()
      gtb.ResumeLayout()
      dgr.ResumeLayout()
      Me.ResumeLayout()
    End Try
  End Sub

  Private Sub mvNetworkTreeView_SetBrowserMenuContext(ByVal psender As Object) Handles mvNetworkTreeview.SetBrowserMenuContext
    MainHelper.SetBrowserMenu(psender, Me)
  End Sub

  Private Sub mvNetworkTreeView_SetBrowserProperty(ByVal psender As Object, ByVal pNodeInfo As CDBNETCL.TreeViewNodeInfo) Handles mvNetworkTreeview.SetBrowserProperty
    If CType(psender, TreeView) IsNot Nothing OrElse CType(psender, VistaTreeView) IsNot Nothing Then
      Dim vBrowserMenu As BrowserMenu = CType(CType(psender, VistaTreeView).ContextMenuStrip, BrowserMenu)
      If vBrowserMenu IsNot Nothing Then
        If Not pNodeInfo.IsFolder AndAlso pNodeInfo.AddressNumber = 0 Then
          vBrowserMenu.EntityType = HistoryEntityTypes.hetContacts
          vBrowserMenu.ItemNumber = pNodeInfo.ContactNumber
          vBrowserMenu.ItemDescription = pNodeInfo.ContactName
          vBrowserMenu.GroupCode = pNodeInfo.Group
        Else
          vBrowserMenu.EntityType = HistoryEntityTypes.hetNone
          vBrowserMenu.GroupCode = String.Empty
        End If
      End If
    End If
  End Sub

  Private Sub HideDisplayGridPanel(ByVal pHide As Boolean)
    Try
      Me.SuspendLayout()
      splRight.SuspendLayout()
      splTop.SuspendLayout()
      splBottom.SuspendLayout()
      dgr.SuspendLayout()
      splRight.Panel1Collapsed = pHide
      splRight.TabStop = Not pHide
      If Not pHide Then
        dgr.SetToolBarVisible()
        splRight.SplitterDistance = dgr.Height
      End If
    Finally
      dgr.ResumeLayout()
      splBottom.ResumeLayout()
      splTop.ResumeLayout()
      splRight.ResumeLayout()
      Me.ResumeLayout()
    End Try
  End Sub

  Private Sub HideAdditionalControls(ByVal pExcept As Control)
    Try
      Me.SuspendLayout()
      If mvCalendarView IsNot Nothing AndAlso pExcept IsNot mvCalendarView Then mvCalendarView.Visible = False
      If mvMarketingChart IsNot Nothing AndAlso pExcept IsNot mvMarketingChart Then mvMarketingChart.Visible = False
      If mvDashboard IsNot Nothing AndAlso pExcept IsNot mvDashboard Then mvDashboard.Visible = False
      If mvNetworkTreeview IsNot Nothing AndAlso pExcept IsNot mvNetworkTreeview Then mvNetworkTreeview.Visible = False
      If mvCareWebBrowser IsNot Nothing AndAlso pExcept IsNot mvCareWebBrowser Then mvCareWebBrowser = Nothing
      If mvExamUnitSelector IsNot Nothing AndAlso pExcept IsNot mvExamUnitSelector Then mvExamUnitSelector.Visible = False
    Finally
      Me.ResumeLayout()
    End Try
  End Sub

  Private Sub splRight_SplitterMoved(ByVal sender As Object, ByVal e As System.Windows.Forms.SplitterEventArgs) Handles splRight.SplitterMoved
    dgr.Height = splRight.SplitterDistance
  End Sub

  Private Sub ProcessRowSelection(ByVal pRow As Integer, ByVal pDataRow As Integer)
    Dim vTabCount As Integer
    Try
      Me.SuspendLayout()
      dpl.SuspendLayout()
      dts.ResumeLayout()
      dts.DisplayGrid(0).ResumeLayout()
      dts.DisplayGrid(1).ResumeLayout()
      dts.DisplayGrid(2).ResumeLayout()
      dts.DisplayGrid(3).ResumeLayout()
      dts.SetText("")

      If mvDataType = CareNetServices.XMLContactDataSelectionTypes.xcdtContactExamCertificates Then
        mvContactExamCertificateMenu.ItemsEnabled = False
      End If

      If dgr.DataRowCount > pRow OrElse mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactInformation OrElse mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactLegacy Then
        dpl.Populate(mvDataSet, pDataRow)
        mvDataSet2 = Nothing
        mvDataSet3 = Nothing
        If mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactLegacy Then
          Dim vDataRow As DataRow = DataHelper.GetRowFromDataSet(mvDataSet)
          If vDataRow IsNot Nothing Then
            dplMenuEdit.Visible = True
            dplMenuNew.Visible = False
            mvContactInfo.LegacyNumber = IntegerValue(vDataRow.Item("LegacyNumber").ToString)
          Else
            dplMenuEdit.Visible = False
            dplMenuNew.Visible = True
            mvContactInfo.LegacyNumber = 0
          End If
        End If
        dts.DisplayGrid(0).ShowIfEmpty = False
        dts.DisplayGrid(1).ShowIfEmpty = False
        Dim vNumber As Integer
        Select Case mvDataType
          Case CareServices.XMLContactDataSelectionTypes.xcdtContactCreditCustomers
            Dim vList As New ParameterList(True)
            Dim vRow As DataRow = GetDataRow(mvDataSet, pDataRow)
            vList("SalesLedgerAccount") = vRow.Item("SalesLedgerAccount").ToString
            vList("Company") = vRow.Item("Company").ToString
            vList("AllocationType") = mvFinancialMenu.DisplayTransactionsAllocationType
            mvDataSet2 = DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactSalesLedgerItems, mvContactInfo.ContactNumber, vList)
            dts.DisplayGrid(0).Populate(mvDataSet2)
            dts.DisplayGrid(0).ShowIfEmpty = True
            vTabCount = 2
            mvFinancialMenu.SetContext(mvDataType, GetDataRow(mvDataSet2, 0), mvContactInfo, mvReadOnlyPage)  'Menu is on the sub grid so always want to start with row 0 not the row number of the main grid
            dts.AddSubDisplayGrid(0)    'Add Sub Display Grid (with Splitter) if not already added
            RemoveHandler dts.SubDisplayGrid(0).RowSelected, AddressOf dgr_RowSelected
            AddHandler dts.SubDisplayGrid(0).RowSelected, AddressOf dgr_RowSelected

          Case CareServices.XMLContactDataSelectionTypes.xcdtContactActions, CareServices.XMLContactDataSelectionTypes.xcdtContactLegacyActions
            vNumber = SelectRowItemNumber(pDataRow, "ActionNumber")
            mvActionMenu.ActionNumber = vNumber
            mvActionMenu.ActionStatus = dgr.GetValue(pRow, "ActionStatus")
            mvActionMenu.MasterActionNumber = IntegerValue(dgr.GetValue(pRow, "MasterAction"))
            mvActionMenu.OutlookId = dgr.GetValue(pRow, "OutlookId")
            mvContactInfo.SelectedActionNumber = vNumber
            dts.TabPageText(0) = ControlText.TbpActionText
            dts.SetText(DataHelper.GetActionText(vNumber))
            mvDataSet2 = DataHelper.GetActionData(CareNetServices.XMLActionDataSelectionTypes.xadtActionSubjects, vNumber)
            dts.DisplayGrid(0).Populate(mvDataSet2)
            dts.DisplayGrid(0).ShowIfEmpty = True
            mvDataSet3 = DataHelper.GetActionData(CareNetServices.XMLActionDataSelectionTypes.xadtActionLinks, vNumber)
            dts.DisplayGrid(1).Populate(mvDataSet3)
            mvActionMenu.SetNotify(dts.DisplayGrid(1))
            Dim e As New System.ComponentModel.CancelEventArgs
            mvActionMenu.SetVisibleItems(e)
            dgr.SetToolBarVisible()
            vTabCount = 3

          Case CareServices.XMLContactDataSelectionTypes.xcdtContactAddresses
            Dim vList As New ParameterList(True)
            vNumber = SelectRowItemNumber(pDataRow, "AddressNumber")
            mvContactInfo.SelectedAddressNumber = vNumber
            vList.IntegerValue("AddressNumber") = vNumber
            dts.DisplayGrid(0).Populate(DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactAddressPositionAndOrganisation, mvContactInfo.ContactNumber, vList), "OrganisationNumber <> " & mvContactInfo.ContactNumber)
            If dts.DisplayGrid(0).DisplayTitle.Length = 0 Then dts.DisplayGrid(0).DisplayTitle = "Position and Organisation"
            If dts.DisplayGrid(0).RowCount = 0 Then dts.DisplayGrid(0).Clear()
            mvDataSet3 = DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactAddressUsages, mvContactInfo.ContactNumber, vList)
            dts.DisplayGrid(1).AutoSetRowHeight = True    'This will allow Notes column to show multiple lines
            dts.DisplayGrid(1).Populate(mvDataSet3)
            Dim vNotesCol As Integer = dts.DisplayGrid(1).GetColumn("Notes")
            If vNotesCol >= 0 Then dts.DisplayGrid(1).SetPreferredColumnWidth(dts.DisplayGrid(1).GetColumn("Notes"))
            dts.DisplayGrid(1).ShowIfEmpty = True
            vTabCount = 3

          Case CareServices.XMLContactDataSelectionTypes.xcdtContactBankAccounts
            vNumber = SelectRowItemNumber(pDataRow, "BankDetailsNumber")
            dts.DisplayGrid(0).Populate(DataHelper.GetBankAccountData(CareServices.XMLBankAccountDataSelectionTypes.xadtBACSAmendments, vNumber))
            vTabCount = 2

          Case CareServices.XMLContactDataSelectionTypes.xcdtContactCreditCardAuthorities
            mvFinancialMenu.SetContext(mvDataType, GetDataRow(mvDataSet, pDataRow), mvContactInfo, mvReadOnlyPage)

          Case CareServices.XMLContactDataSelectionTypes.xcdtContactCommsNumbers
            Dim vList As New ParameterList(True)
            vNumber = SelectRowItemNumber(pDataRow, "CommunicationNumber")
            mvContactInfo.SelectedCommunicationNumber = vNumber
            mvContactInfo.SelectedCommunicationDevice = SelectRowItem(pDataRow, "DeviceCode")
            vList.IntegerValue("CommunicationNumber") = vNumber
            mvDataSet2 = DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactCommunicationUsages, mvContactInfo.ContactNumber, vList)
            dts.DisplayGrid(0).Populate(mvDataSet2)
            dts.DisplayGrid(0).ShowIfEmpty = True
            vTabCount = 2

          Case CareServices.XMLContactDataSelectionTypes.xcdtContactCovenants
            vNumber = SelectRowItemNumber(pDataRow, "CovenantNumber")
            dts.DisplayGrid(0).Populate(DataHelper.GetCovenantData(CareServices.XMLCovenantDataSelectionTypes.xcdtCovenantGiftAidClaims, vNumber))
            dts.DisplayGrid(1).Populate(DataHelper.GetCovenantData(CareServices.XMLCovenantDataSelectionTypes.xcdtCovenantClaims, vNumber))
            vTabCount = 3
            mvFinancialMenu.SetContext(mvDataType, GetDataRow(mvDataSet, pDataRow), mvContactInfo, mvReadOnlyPage)

          Case CareServices.XMLContactDataSelectionTypes.xcdtContactCustomFormData
            Dim vList As ParameterList = New ParameterList(True)
            vList("Detail") = "Y"
            vList.IntegerValue("CustomForm") = mvCustomFormNumber
            vList("Values") = dgr.GetRowCSValues(pRow)
            Dim vRow As DataRow = Nothing
            Dim vTable As DataTable = DataHelper.GetTableFromDataSet(DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactCustomFormData, mvContactInfo.ContactNumber, vList))
            If vTable IsNot Nothing Then vRow = vTable.Rows(0)
            If Not vRow Is Nothing Then
              mvSelectedRow = pRow
              mvCustomPanel.Populate(vRow)
              If mvAllowUpdate = False AndAlso IsEditableCustomForm(mvCustomFormNumber) Then
                mvCustomPanel.EnableControls(mvCustomPanel, False)
                cmdSave.Enabled = False
              End If
              cmdDelete.Enabled = True
            Else
              ProcessNew()
            End If

          Case CareServices.XMLContactDataSelectionTypes.xcdtContactDespatchNotes
            vNumber = SelectRowItemNumber(pDataRow, "BatchNumber")
            Dim vTNumber As Integer = SelectRowItemNumber(pDataRow, "TransactionNumber")
            Dim vPLNumber As Integer = SelectRowItemNumber(pDataRow, "PickingListNumber")
            Dim vDNNumber As Integer = SelectRowItemNumber(pDataRow, "DespatchNoteNumber")
            ' BR 11665 - Added new parameter (vDNNumber)below 
            dts.DisplayGrid(0).Populate(DataHelper.GetTransactionData(CareServices.XMLTransactionDataSelectionTypes.xtdtDespatchStock, vNumber, vTNumber, vPLNumber, 0, vDNNumber))
            vTabCount = 2
            mvFinancialMenu.SetContext(mvDataType, GetDataRow(mvDataSet, pDataRow), mvContactInfo, mvReadOnlyPage)

          Case CareServices.XMLContactDataSelectionTypes.xcdtContactDirectDebits
            vNumber = SelectRowItemNumber(pDataRow, "DirectDebitNumber")
            dts.DisplayGrid(0).Populate(DataHelper.GetDirectDebitData(CareServices.XMLDirectDebitDataSelectionTypes.xbdtBACSAmendments, vNumber))
            vTabCount = 2
            mvFinancialMenu.SetContext(mvDataType, GetDataRow(mvDataSet, pDataRow), mvContactInfo, mvReadOnlyPage)

          Case CareServices.XMLContactDataSelectionTypes.xcdtContactDocuments
            vNumber = SelectRowItemNumber(pDataRow, "DocumentNumber")
            mvDocumentMenu.DocumentNumber = vNumber
            mvContactInfo.SelectedDocumentNumber = vNumber
            Dim e As New System.ComponentModel.CancelEventArgs
            mvDocumentMenu.SetVisibleItems(e)
            dgr.SetToolBarVisible()
            ShowDocumentDetails(vNumber)
            vTabCount = 4
          Case CareServices.XMLContactDataSelectionTypes.xcdtContactEventBookings, CareServices.XMLContactDataSelectionTypes.xcdtContactAppropriateCertificates,
               CareNetServices.XMLContactDataSelectionTypes.xcdtContactEventRoomBookings
            SetFinancialMenuToolbar(GetDataRow(mvDataSet, pDataRow))

          Case CareServices.XMLContactDataSelectionTypes.xcdtContactEventDelegates
            vNumber = SelectRowItemNumber(pDataRow, "DelegateNumber")
            mvContactEventDelegateMenu.SetContext(mvContactInfo, GetDataRow(mvDataSet, pDataRow))
            dts.DisplayGrid(0).Populate(DataHelper.GetDelegateData(CareServices.XMLDelegateDataSelectionTypes.xeddtActivities, vNumber))
            dts.DisplayGrid(1).Populate(DataHelper.GetDelegateData(CareServices.XMLDelegateDataSelectionTypes.xeddtLinks, vNumber))
            vTabCount = 3

          Case CareServices.XMLContactDataSelectionTypes.xcdtContactExamSummary
            vNumber = SelectRowItemNumber(pDataRow, "ExamStudentHeaderId")
            mvExamUnitSelector.InitForSummary(vNumber)

          Case CareServices.XMLContactDataSelectionTypes.xcdtContactExamDetails
            vNumber = SelectRowItemNumber(pDataRow, "ExamBookingId")
            Dim vSessionId As Integer = SelectRowItemNumber(pDataRow, "ExamSessionId")
            mvExamUnitSelector.InitForBooking(vSessionId, mvContactInfo.ContactNumber, vNumber, True)
            mvFinancialMenu.SetContext(mvDataType, GetDataRow(mvDataSet, pDataRow), mvContactInfo, mvReadOnlyPage)

          Case CareServices.XMLContactDataSelectionTypes.xcdtContactExamExemptions
            SetFinancialMenuToolbar(GetDataRow(mvDataSet, pDataRow))
            Dim vList As New ParameterList(True, True)
            vList.IntegerValue("ExamStudentExemptionId") = SelectRowItemNumber(pDataRow, "ExamStudentExemptionId")
            dts.DisplayGrid(0).Populate(ExamsDataHelper.GetExamData(ExamsAccess.XMLExamDataSelectionTypes.ExamStudentExemptionHistory, vList))
            vTabCount = 2

          Case CareServices.XMLContactDataSelectionTypes.xcdtContactFundraising
            vNumber = SelectRowItemNumber(pDataRow, "FundraisingRequestNumber")
            If AppValues.ControlValue(AppValues.ControlValues.fundraising_status).Length > 0 Then
              mvDataSet2 = DataHelper.GetFundraisingData(CareNetServices.XMLFundraisingDataSelectionTypes.xfdtPaymentSchedule, vNumber)
              dts.DisplayGrid(0).Populate(mvDataSet2)
              dts.DisplayGrid(0).ShowIfEmpty = True
              dts.AddSubDisplayGrid(0)    'Add Sub Display Grid (with Splitter) if not already added
              RemoveHandler dts.SubDisplayGrid(0).RowSelected, AddressOf dgr_RowSelected
              AddHandler dts.SubDisplayGrid(0).RowSelected, AddressOf dgr_RowSelected
            Else
              dts.DisplayGrid(0).Clear()
            End If
            dts.DisplayGrid(1).Populate(DataHelper.GetFundraisingData(CareNetServices.XMLFundraisingDataSelectionTypes.xfdtTargets, vNumber))
            dts.DisplayGrid(2).Populate(DataHelper.GetFundraisingData(CareNetServices.XMLFundraisingDataSelectionTypes.xfdtExpectedAmountHistory, vNumber))
            dts.DisplayGrid(3).Populate(DataHelper.GetFundraisingData(CareNetServices.XMLFundraisingDataSelectionTypes.xfdtRequestStatusHistory, vNumber))
            vTabCount = 5
            If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.option_actions) Then
              mvDataSet4 = DataHelper.GetFundraisingData(CareNetServices.XMLFundraisingDataSelectionTypes.xfdtFundraisingActions, vNumber)
              dts.DisplayGrid(4).Populate(mvDataSet4)
              dts.DisplayGrid(4).ShowIfEmpty = SelectRowItem(pDataRow, "Logname").Length > 0
              vTabCount += 1
            Else
              dts.DisplayGrid(4).Clear()
              dts.DisplayGrid(4).ShowIfEmpty = False
            End If
            'BR19023
            mvDataSet3 = DataHelper.GetFundraisingData(CareNetServices.XMLFundraisingDataSelectionTypes.xfdtFundraisingDocuments, vNumber)
            dts.DisplayGrid(5).Populate(mvDataSet3)
            dts.DisplayGrid(5).ShowIfEmpty = False
            vTabCount += 1

            SetFinancialMenuToolbar(GetDataRow(mvDataSet, pDataRow))

          Case CareServices.XMLContactDataSelectionTypes.xcdtContactFundraisingEvents
            vNumber = SelectRowItemNumber(pDataRow, "ContactFundraisingNumber")
            dts.DisplayGrid(0).Populate(DataHelper.GetFundraisingEventData(CareNetServices.XMLFundraisingEventDataSelectionTypes.xfdtAnalysis, vNumber))
            vTabCount = 2

          Case CareServices.XMLContactDataSelectionTypes.xcdtContactGiftAidDeclarations
            vNumber = SelectRowItemNumber(pDataRow, "DeclarationNumber")
            dts.DisplayGrid(0).Populate(DataHelper.GetGiftAidData(CareServices.XMLGiftAidDataSelectionTypes.xgdtGiftAidUnClaimedPayments, vNumber))
            dts.DisplayGrid(1).Populate(DataHelper.GetGiftAidData(CareServices.XMLGiftAidDataSelectionTypes.xgdtGiftAidClaimedPayments, vNumber))
            vTabCount = 3
            SetFinancialMenuToolbar(GetDataRow(mvDataSet, pDataRow))

          Case CareServices.XMLContactDataSelectionTypes.xcdtContactH2HCollections
            vNumber = mvContactInfo.ContactNumber
            Dim vList As New ParameterList(True)
            vList.IntegerValue("CollectionNumber") = SelectRowItemNumber(pDataRow, "CollectionNumber")
            dts.DisplayGrid(0).Populate(DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactCollectionPayments, vNumber, vList))
            dts.DisplayGrid(0).ShowIfEmpty = True
            vTabCount = 2

          Case CareServices.XMLContactDataSelectionTypes.xcdtContactInformation
            dts.DisplayGrid(0).Populate(DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactHighProfileCategories, mvContactInfo.ContactNumber))
            dts.DisplayGrid(1).Populate(DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactDepartmentCategories, mvContactInfo.ContactNumber))
            dts.DisplayGrid(2).Populate(CareServices.XMLContactDataSelectionTypes.xcdtContactPreviousNames, mvContactInfo.ContactNumber)
            vTabCount = 4
            If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.cd_high_profile_links) Then
              dts.DisplayGrid(3).Populate(DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactHPLinks, mvContactInfo.ContactNumber))
              vTabCount = 5
            End If

          Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactAmendments
            Dim vList As New ParameterList(True, True)
            vList("OperationDate") = SelectRowItem(pDataRow, "OperationDate")
            vList("Operation") = SelectRowItem(pDataRow, "Operation")
            vList("TableName") = SelectRowItem(pDataRow, "TableName")
            Dim vJournalNumber As String = SelectRowItem(pDataRow, "JournalNumber")
            If IntegerValue(vJournalNumber) > 0 Then vList("JournalNumber") = vJournalNumber
            dts.DisplayGrid(0).Populate(DataHelper.GetContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactAmendmentDetails, mvContactInfo.ContactNumber, vList))
            If vList("Operation") = "update" Then
              For vIndex As Integer = 0 To dts.DisplayGrid(0).RowCount - 1
                If dts.DisplayGrid(0).GetValue(vIndex, "OldValues") <> dts.DisplayGrid(0).GetValue(vIndex, "NewValues") Then
                  dts.DisplayGrid(0).SetBoldRow(vIndex)
                End If
              Next
            End If
            vTabCount = 2

          Case CareServices.XMLContactDataSelectionTypes.xcdtContactJournals
            vNumber = SelectRowItemNumber(pDataRow, "ID")
            mvJournalLinkMenu.SetContext(mvContactInfo, GetDataRow(mvDataSet, pDataRow))
            dts.DisplayGrid(0).Populate(DataHelper.GetJournalData(CareServices.XMLJournalDataSelectionTypes.xadtJournalInformation, vNumber, SelectRowItem(pDataRow, "JournalType"), mvContactInfo))
            If dts.DisplayGrid(0).RequiredWidth > dts.DisplayGrid(0).Width Then
              Dim vOffset As Integer = dts.DisplayGrid(0).RequiredWidth - dts.DisplayGrid(0).Width
              If dts.DisplayGrid(0).GetColumnWidth(1) > vOffset Then
                dts.DisplayGrid(0).SetColumnWidth(1, CInt(dts.DisplayGrid(0).GetColumnWidth(1) - vOffset))
              End If
            End If
            vTabCount = 2

          Case CareServices.XMLContactDataSelectionTypes.xcdtContactLegacyBequests
            vNumber = SelectRowItemNumber(pDataRow, "LegacyNumber")
            Dim vBequestNumber As Integer = SelectRowItemNumber(pDataRow, "BequestNumber")
            mvContactInfo.SelectedBequestNumber = vBequestNumber
            Dim vList As New ParameterList(True)
            vList.IntegerValue("LegacyNumber") = vNumber
            vList.IntegerValue("BequestNumber") = vBequestNumber
            dts.DisplayGrid(0).ShowIfEmpty = True
            mvDataSet2 = DataHelper.GetLegacyBequestData(CareNetServices.XMLLegacyBequestDataSelectionTypes.xlbdstReceipts, vList)
            dts.DisplayGrid(0).Populate(mvDataSet2)
            If dts.DisplayGrid(0).DataRowCount > 0 Then ProcessSubRowSelection(dts.DisplayGrid(0), 0, 0)
            dts.DisplayGrid(1).ShowIfEmpty = True
            mvDataSet3 = DataHelper.GetLegacyBequestData(CareNetServices.XMLLegacyBequestDataSelectionTypes.xlbdstForecasts, vList)
            dts.DisplayGrid(1).Populate(mvDataSet3)
            vTabCount = 3

          Case CareServices.XMLContactDataSelectionTypes.xcdtContactLegacyLinks
            dgrMenuNewDocument.Visible = True

          Case CareServices.XMLContactDataSelectionTypes.xcdtContactMailings
            vNumber = SelectRowItemNumber(pDataRow, "AddressNumber")
            ShowContactMailingDetails(vNumber, pDataRow)
            vTabCount = 3
            'Need a change in VB 6 code to access xcdtContactCommunicationHistory directly 
            'till then access it using <TODO> 
          Case CareServices.XMLContactDataSelectionTypes.xcdtContactCommunicationHistory
            vNumber = SelectRowItemNumber(pDataRow, "AddressNumber")
            If vNumber > 0 Then
              mvViewMailingDocumentMenu = New MailingDocumentMenu(Me, CareServices.XMLContactDataSelectionTypes.xcdtContactMailings, mvContactInfo)
              dgrMenuNew.Visible = False
              dgrMenuEdit.Visible = False
              dgr.ContextMenuStrip = mvViewMailingDocumentMenu
              dgr.SetToolBarVisible()
              ShowContactMailingDetails(vNumber, pDataRow)
              vTabCount = 3
            Else
              If mvDocumentMenu Is Nothing Then mvDocumentMenu = New DocumentMenu(Me)
              mvDocumentMenu.DocumentType = BaseDocumentMenu.DocumentTypes.ContactDocuments
              dgr.ContextMenuStrip = mvDocumentMenu
              dgr.SetToolBarVisible()
              dgrMenuNew.Visible = True
              dgrMenuEdit.Visible = True
              dts.GridContextMenuStrip(0) = dgr0MenuStrip
              dts.GridContextMenuStrip(1) = dgr1MenuStrip
              mvDocumentMenu.DocumentNumber = 0
              vNumber = SelectRowItemNumber(pDataRow, "MailingNumber")
              ShowDocumentDetails(vNumber)
              vTabCount = 4
            End If

          Case CareServices.XMLContactDataSelectionTypes.xcdtContactMannedCollections
            vNumber = mvContactInfo.ContactNumber
            Dim vList As New ParameterList(True)
            vList.IntegerValue("CollectionNumber") = SelectRowItemNumber(pDataRow, "CollectionNumber")
            dts.DisplayGrid(2).Populate(DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactCollectionPayments, vNumber, vList))
            dts.DisplayGrid(2).ShowIfEmpty = True
            dts.TabPageText(2) = ControlText.TbpIncome
            vList.IntegerValue("ContactNumber") = vNumber
            dts.DisplayGrid(0).Populate(DataHelper.GetCampaignData(CareServices.XMLCampaignDataSelectionTypes.xcadtCollectorShifts, vList))
            dts.DisplayGrid(0).ShowIfEmpty = True
            dts.DisplayGrid(1).Populate(DataHelper.GetCampaignData(CareServices.XMLCampaignDataSelectionTypes.xcadtContactCollectionBoxes, vList))
            dts.DisplayGrid(1).ShowIfEmpty = True
            vTabCount = 4

          Case CareServices.XMLContactDataSelectionTypes.xcdtContactMembershipDetails
            vNumber = SelectRowItemNumber(pDataRow, "MembershipNumber")
            dts.DisplayGrid(0).Populate(DataHelper.GetMembershipData(CareServices.XMLMembershipDataSelectionTypes.xmdtMembershipPaymentPlanDetails, vNumber))
            Dim vPPNumber As Integer = SelectRowItemNumber(pDataRow, "PaymentPlanNumber")
            mvDataSet3 = DataHelper.GetPaymentPlanData(CareServices.XMLPaymentPlanDataSelectionTypes.xpdtPaymentPlanPayments, vPPNumber)
            dts.DisplayGrid(1).Populate(mvDataSet3)
            If dts.DisplayGrid(1).DataRowCount > 0 Then ProcessSubRowSelection(dts.DisplayGrid(1), 0, 0)
            dts.DisplayGrid(2).Populate(DataHelper.GetMembershipData(CareServices.XMLMembershipDataSelectionTypes.xmdtMembershipOtherMembers, vNumber))
            dts.DisplayGrid(3).Populate(DataHelper.GetMembershipData(CareServices.XMLMembershipDataSelectionTypes.xmdtMembershipChanges, vNumber))
            If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.fp_pay_plan_amendment_history) AndAlso AppValues.ConfigurationOption(AppValues.ConfigurationOptions.option_journal, True) Then
              dts.DisplayGrid(4).Populate(DataHelper.GetPaymentPlanData(CareServices.XMLPaymentPlanDataSelectionTypes.xpdtPaymentPlanAmendmentHistory, vPPNumber))
              dts.DisplayGrid(4).ShowIfEmpty = True
              vTabCount = 6
            Else
              vTabCount = 5
            End If
            If BooleanValue(SelectRowItem(pDataRow, "BranchMembership")) = True AndAlso AppValues.ControlValue(AppValues.ControlTables.membership_controls, AppValues.ControlValues.organisation_group).Length > 0 Then
              Dim vDGR As DisplayGrid = dts.DisplayGrid(5)
              If vTabCount = 5 Then vDGR = dts.DisplayGrid(4)
              'Populate MembershipHistory
              vDGR.Populate(DataHelper.GetMembershipData(CareServices.XMLMembershipDataSelectionTypes.xmdtMembershipGroups, vNumber))
              vDGR.ShowIfEmpty = True
              vTabCount += 1
              vDGR = dts.DisplayGrid(6)
              If vTabCount = 6 Then vDGR = dts.DisplayGrid(5)
              'Populate MembershipHistoryDetails
              vDGR.Populate(DataHelper.GetMembershipData(CareServices.XMLMembershipDataSelectionTypes.xmdtMembershipGroupHistory, vNumber))
              vDGR.ShowIfEmpty = True
              vTabCount += 1
            End If
            mvFinancialMenu.SetContext(mvDataType, GetDataRow(mvDataSet, pDataRow), mvContactInfo, mvReadOnlyPage)

          Case CareServices.XMLContactDataSelectionTypes.xcdtContactPaymentPlans
            vNumber = SelectRowItemNumber(pDataRow, "PaymentPlanNumber")
            dts.DisplayGrid(0).Populate(DataHelper.GetPaymentPlanData(CareServices.XMLPaymentPlanDataSelectionTypes.xpdtPaymentPlanDetails, vNumber))
            mvDataSet3 = DataHelper.GetPaymentPlanData(CareServices.XMLPaymentPlanDataSelectionTypes.xpdtPaymentPlanPayments, vNumber)
            mvDataSet4 = DataHelper.GetPaymentPlanData(CareServices.XMLPaymentPlanDataSelectionTypes.xpdtPaymentPlanMembers, vNumber)
            dts.DisplayGrid(1).Populate(mvDataSet3)
            dts.DisplayGrid(2).Populate(mvDataSet4)
            dts.DisplayGrid(3).Populate(DataHelper.GetPaymentPlanData(CareServices.XMLPaymentPlanDataSelectionTypes.xpdtPaymentPlanSubscriptions, vNumber))
            If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.fp_pay_plan_amendment_history) AndAlso AppValues.ConfigurationOption(AppValues.ConfigurationOptions.option_journal, True) Then
              dts.DisplayGrid(4).Populate(DataHelper.GetPaymentPlanData(CareServices.XMLPaymentPlanDataSelectionTypes.xpdtPaymentPlanAmendmentHistory, vNumber))
              dts.DisplayGrid(4).ShowIfEmpty = True
              vTabCount = 6
            Else
              vTabCount = 5
            End If
            mvFinancialMenu.SetContext(mvDataType, GetDataRow(mvDataSet, pDataRow), mvContactInfo, mvReadOnlyPage)
            If dts.DisplayGrid(1).DataRowCount > 0 Then ProcessSubRowSelection(dts.DisplayGrid(1), 0, 0)
            If dts.DisplayGrid(2).DataRowCount > 0 Then ProcessSubRowSelection(dts.DisplayGrid(2), 0, 0)

          Case CareServices.XMLContactDataSelectionTypes.xcdtContactPledges
            SetFinancialMenuToolbar(GetDataRow(mvDataSet, pDataRow))
            mvFinancialMenu.SetContext(mvDataType, GetDataRow(mvDataSet, pDataRow), mvContactInfo, mvReadOnlyPage)
            vNumber = SelectRowItemNumber(pDataRow, "PledgeNumber")
            mvDataSet2 = DataHelper.GetPledgeData(CareServices.XMLPledgeDataSelectionTypes.xgdtPledgePayments, vNumber)
            dts.DisplayGrid(0).Populate(mvDataSet2)
            If dts.DisplayGrid(0).DataRowCount > 0 Then ProcessSubRowSelection(dts.DisplayGrid(0), 0, 0)
            vTabCount = 2

          Case CareServices.XMLContactDataSelectionTypes.xcdtContactPositions, CareNetServices.XMLContactDataSelectionTypes.xcdtContactViewOrganisations
            Dim vList As New ParameterList(True)
            vNumber = SelectRowItemNumber(pDataRow, "ContactNumber")
            If mvContactInfo.ContactType = ContactInfo.ContactTypes.ctOrganisation Then
              vList.IntegerValue("ContactNumber2") = vNumber
            Else
              vList.IntegerValue("OrganisationNumber") = vNumber
              mvContactInfo.OrganisationNumber = vNumber
            End If
            mvContactInfo.SelectedContactNumber2 = vNumber
            'Roles
            mvDataSet2 = DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactRoles, mvContactInfo.ContactNumber, vList)
            dts.DisplayGrid(0).Populate(mvDataSet2)
            dts.DisplayGrid(0).ShowIfEmpty = True
            vNumber = SelectRowItemNumber(pDataRow, "ContactPositionNumber")
            mvContactInfo.SelectedContactPositionNumber = vNumber
            mvContactInfo.SelectedContactPositionValidFrom = SelectRowItem(pDataRow, "ValidFrom")
            mvContactInfo.SelectedContactPositionValidTo = SelectRowItem(pDataRow, "ValidTo")
            'Activities
            vList = New ParameterList(True)
            vList.IntegerValue("ContactPositionNumber") = vNumber
            mvDataSet3 = DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactPositionActivities, mvContactInfo.ContactNumber, vList)
            dts.DisplayGrid(1).Populate(mvDataSet3)
            dts.DisplayGrid(1).ShowIfEmpty = True
            'Links
            mvDataSet4 = DataHelper.GetContactData(CType(CareNetServices.XMLContactDataSelectionTypes.xcdtContactPositionLinks, CareServices.XMLContactDataSelectionTypes), mvContactInfo.ContactNumber, vList)
            dts.DisplayGrid(2).Populate(mvDataSet4)
            dts.DisplayGrid(2).ShowIfEmpty = True
            'Actions
            mvDataSet5 = DataHelper.GetContactData(CType(CareNetServices.XMLContactDataSelectionTypes.xcdtContactPositionActions, CareServices.XMLContactDataSelectionTypes), mvContactInfo.ContactNumber, vList)
            dts.DisplayGrid(3).Populate(mvDataSet5)
            dts.DisplayGrid(3).ShowIfEmpty = True
            mvActionMenu = New ActionMenu(Me)
            mvActionMenu.ContactPositionNumber = vNumber
            mvActionMenu.ActionType = BaseActionMenu.ActionTypes.PositionActions
            mvActionMenu.SetVisibleItems(New System.ComponentModel.CancelEventArgs)
            dts.DisplayGrid(3).ContextMenuStrip = mvActionMenu
            If dts.DisplayGrid(3).DataRowCount > 0 Then ProcessSubRowSelection(dts.DisplayGrid(3), 0, 0)
            'Documents
            mvDataSet6 = DataHelper.GetContactData(CType(CareNetServices.XMLContactDataSelectionTypes.xcdtContactPositionDocuments, CareServices.XMLContactDataSelectionTypes), mvContactInfo.ContactNumber, vList)
            dts.DisplayGrid(4).Populate(mvDataSet6)
            dts.DisplayGrid(4).ShowIfEmpty = True
            mvDocumentMenu = New DocumentMenu(Me)
            mvDocumentMenu.ContactPositionNumber = vNumber
            mvDocumentMenu.DocumentType = BaseDocumentMenu.DocumentTypes.PositionDocuments
            mvDocumentMenu.SetVisibleItems(New System.ComponentModel.CancelEventArgs)
            dts.DisplayGrid(4).ContextMenuStrip = mvDocumentMenu
            If dts.DisplayGrid(4).DataRowCount > 0 Then ProcessSubRowSelection(dts.DisplayGrid(4), 0, 0)
            'Timesheet
            mvDataSet7 = DataHelper.GetContactData(CType(CareNetServices.XMLContactDataSelectionTypes.xcdtContactPositionTimesheets, CareServices.XMLContactDataSelectionTypes), mvContactInfo.ContactNumber, vList)
            dts.DisplayGrid(5).Populate(mvDataSet7)
            dts.DisplayGrid(5).ShowIfEmpty = True

            vTabCount = 7


          Case CareServices.XMLContactDataSelectionTypes.xcdtContactPostPGPledges
            SetFinancialMenuToolbar(GetDataRow(mvDataSet, pDataRow))
            mvFinancialMenu.SetContext(mvDataType, GetDataRow(mvDataSet, pDataRow), mvContactInfo, mvReadOnlyPage)
            vNumber = SelectRowItemNumber(pDataRow, "PledgeNumber")
            mvDataSet2 = DataHelper.GetPledgeData(CareServices.XMLPledgeDataSelectionTypes.xgdtPostPGPledgePayments, vNumber)
            dts.DisplayGrid(0).Populate(mvDataSet2)
            If dts.DisplayGrid(0).DataRowCount > 0 Then ProcessSubRowSelection(dts.DisplayGrid(0), 0, 0)
            vTabCount = 2

          Case CareServices.XMLContactDataSelectionTypes.xcdtContactProcessedTransactions
            vNumber = SelectRowItemNumber(pDataRow, "BatchNumber")
            Dim vTNumber As Integer = SelectRowItemNumber(pDataRow, "TransactionNumber")
            mvDataSet2 = DataHelper.GetTransactionData(CareServices.XMLTransactionDataSelectionTypes.xtdtFinancialHistoryDetails, vNumber, vTNumber)
            dts.DisplayGrid(0).Populate(mvDataSet2)
            If dts.DisplayGrid(0).DataRowCount > 0 Then ProcessSubRowSelection(dts.DisplayGrid(0), 0, 0)
            mvFinancialMenu.SetContext(mvDataType, GetDataRow(mvDataSet, pDataRow), mvContactInfo, mvReadOnlyPage, dgr.MultipleRowsSelected)
            mvAnalysisFinancialMenu.SetContext(mvContactInfo, mvFinancialMenu.DataRow, GetDataRow(mvDataSet, pDataRow))
            If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.fp_pp_show_payment_details) Then
              dts.DisplayGrid(1).Populate(DataHelper.GetPaymentPlanData(CareNetServices.XMLPaymentPlanDataSelectionTypes.xpdtPaymentPlanPaymentHistDetails, 0, 0, "", 0))
              dts.DisplayGrid(1).ShowIfEmpty = True
              vTabCount = 3
            Else
              vTabCount = 2
            End If
            If dts.DisplayGrid(0).DataRowCount > 0 Then ProcessSubRowSelection(dts.DisplayGrid(0), 0, 0)

          Case CareServices.XMLContactDataSelectionTypes.xcdtContactPurchaseInvoices
            vNumber = SelectRowItemNumber(pDataRow, "PurchaseInvoiceNumber")
            dts.DisplayGrid(0).Populate(DataHelper.GetPurchaseInvoiceData(CareServices.XMLPurchaseInvoiceDataSelectionTypes.xodtPurchaseInvoiceDetails, vNumber))
            Dim vChequeReferenceNumber As Integer = SelectRowItemNumber(pDataRow, "ChequeReferenceNumber")
            If vChequeReferenceNumber > 0 Then
              mvDataSet3 = DataHelper.GetPurchaseInvoiceData(CareServices.XMLPurchaseInvoiceDataSelectionTypes.xodtPurchaseInvoiceChequeInformation, 0, vChequeReferenceNumber)
              dts.DisplayGrid(1).Populate(mvDataSet3)
              If dts.DisplayGrid(1).DataRowCount > 0 Then ProcessSubRowSelection(dts.DisplayGrid(1), 0, 0)
              vTabCount = 3
            Else
              vTabCount = 2
            End If
            'BR17340
            mvPurchaseInvoiceMenu.SetContext(mvContactInfo, GetDataRow(mvDataSet, pDataRow))

          Case CareServices.XMLContactDataSelectionTypes.xcdtContactPurchaseOrders
            If mvGridCurrentRow > 0 AndAlso mvGridCurrentRow < dgr.DataRowCount Then pDataRow = mvGridCurrentRow
            vNumber = SelectRowItemNumber(pDataRow, "PurchaseOrderNumber")
            dts.DisplayGrid(0).Populate(DataHelper.GetPurchaseOrderData(CareServices.XMLPurchaseOrderDataSelectionTypes.xodtPurchaseOrderDetails, vNumber))
            mvDataSet3 = DataHelper.GetPurchaseOrderData(CareServices.XMLPurchaseOrderDataSelectionTypes.xodtPurchaseOrderPayments, vNumber)
            dts.DisplayGrid(1).Populate(mvDataSet3)
            'Always Show payments grid if purchase order type allows ad-hoc payments
            dts.DisplayGrid(1).ShowIfEmpty = SelectRowItem(pDataRow, "AdHocPayments").StartsWith("Y")
            mvDataSet4 = DataHelper.GetPurchaseOrderData(CareServices.XMLPurchaseOrderDataSelectionTypes.xodtPurchaseOrderHistory, vNumber)
            dts.DisplayGrid(2).Populate(mvDataSet4)
            vTabCount = 4
            mvPurchaseOrderMenu.SetContext(mvContactInfo, GetDataRow(mvDataSet, pDataRow))
            mvPurchaseOrderPaymentMenu.SetParentContext(mvContactInfo, GetDataRow(mvDataSet, pDataRow))
            If dts.DisplayGrid(1).DataRowCount > 0 Then ProcessSubRowSelection(dts.DisplayGrid(1), 0, 0)
            Dim e As New System.ComponentModel.CancelEventArgs
            mvPurchaseOrderPaymentMenu.SetVisibleItems(e)
            dts.DisplayGrid(1).SetToolBarVisible()
            If (mvGridCurrentRow > 0 Or mvGridCurrentRowPayments > 0) AndAlso mvGridCurrentRow < dgr.DataRowCount Then
              pDataRow = mvGridCurrentRow
              dgr.SelectRow(pDataRow, True)
              dts.DisplayGrid(1).SelectRow(mvGridCurrentRowPayments, True)
            End If
            mvGridCurrentRow = 0
            mvGridCurrentRowPayments = 0
          Case CareServices.XMLContactDataSelectionTypes.xcdtContactScores
            AppValues.SetConfigurationValue(AppValues.ConfigurationValues.profile_user_sequence, dgr.GetSelectedRowIntegers("Sequence").CSList.Replace(",", "|"))
            mvMarketingChart.PopulateFromScores(mvDataSet, True)

          Case CareServices.XMLContactDataSelectionTypes.xcdtContactStandingOrders
            mvFinancialMenu.SetContext(mvDataType, GetDataRow(mvDataSet, pDataRow), mvContactInfo, mvReadOnlyPage)

          Case CareServices.XMLContactDataSelectionTypes.xcdtContactSubscriptions
            mvFinancialMenu.SetContext(mvDataType, GetDataRow(mvDataSet, pDataRow), mvContactInfo, mvReadOnlyPage)

          Case CareServices.XMLContactDataSelectionTypes.xcdtContactUnMannedCollections
            vNumber = mvContactInfo.ContactNumber
            Dim vList As New ParameterList(True)
            vList.IntegerValue("CollectionNumber") = SelectRowItemNumber(pDataRow, "CollectionNumber")
            dts.DisplayGrid(1).Populate(DataHelper.GetCampaignData(CareServices.XMLCampaignDataSelectionTypes.xcadtContactCollectionBoxes, vList))
            dts.DisplayGrid(1).ShowIfEmpty = True
            dts.DisplayGrid(2).Populate(DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactCollectionPayments, vNumber, vList))
            dts.DisplayGrid(2).ShowIfEmpty = True
            dts.DisplayGrid(0).Populate(DataHelper.GetCampaignData(CareServices.XMLCampaignDataSelectionTypes.xcadtCollectionResources, vList))
            dts.DisplayGrid(0).ShowIfEmpty = True
            vTabCount = 4

          Case CareServices.XMLContactDataSelectionTypes.xcdtContactUnProcessedTransactions,
            CareServices.XMLContactDataSelectionTypes.xcdtContactDeliveryTransactions,
            CareServices.XMLContactDataSelectionTypes.xcdtContactSalesTransactions
            vNumber = SelectRowItemNumber(pDataRow, "BatchNumber")
            Dim vTNumber As Integer = SelectRowItemNumber(pDataRow, "TransactionNumber")
            Dim vList As New ParameterList(True, True)
            vList("ContactNumber") = mvContactInfo.ContactNumber.ToString
            Dim vTransactionType As CareNetServices.XMLTransactionDataSelectionTypes
            Select Case mvDataType
              Case CareServices.XMLContactDataSelectionTypes.xcdtContactUnProcessedTransactions
                vTransactionType = CareNetServices.XMLTransactionDataSelectionTypes.xtdtTransactionAnalysis
              Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactSalesTransactions
                vTransactionType = CareNetServices.XMLTransactionDataSelectionTypes.xtdtSalesTransactionAnalysis
              Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactDeliveryTransactions
                vTransactionType = CareNetServices.XMLTransactionDataSelectionTypes.xtdtDeliveryTransactionAnalysis
            End Select
            dts.DisplayGrid(0).Populate(DataHelper.GetTransactionData(vTransactionType, vNumber, vTNumber, vList))
            vTabCount = 2
            mvFinancialMenu.SetContext(mvDataType, GetDataRow(mvDataSet, pDataRow), mvContactInfo, False)

          Case CareServices.XMLContactDataSelectionTypes.xcdtContactXactionInMemoriamDonated,
               CareServices.XMLContactDataSelectionTypes.xcdtContactXactionInMemoriamReceived,
               CareServices.XMLContactDataSelectionTypes.xcdtContactXactionPaidInBy,
               CareServices.XMLContactDataSelectionTypes.xcdtContactXactionSentOnBehalfOf,
               CareServices.XMLContactDataSelectionTypes.xcdtContactXactionContributedTo,
               CareServices.XMLContactDataSelectionTypes.xcdtContactXactionHandledBy
            mvTransactionLinkMenu.SetContext(mvContactInfo, GetDataRow(mvDataSet, pDataRow))

          Case CareServices.XMLContactDataSelectionTypes.xcdtContactCPD
            Dim vList As New ParameterList(True)
            vList.IntegerValue("ContactCPDCycleNumber") = SelectRowItemNumber(pDataRow, "CopyContactCPDCycleNumber")
            Dim vCPDPeriodNumber As Integer = SelectRowItemNumber(pDataRow, "ContactCPDPeriodNumber")
            Dim vCPDPointNumber As Integer = SelectRowItemNumber(pDataRow, "ContactCPDPointNumber")
            If vCPDPeriodNumber > 0 Then vList.IntegerValue("ContactCPDPeriodNumber") = vCPDPeriodNumber
            If SelectRowItem(pDataRow, "CategoryType") <> "" Then vList("CategoryType") = SelectRowItem(pDataRow, "CategoryType")
            If SelectRowItem(pDataRow, "CPDType") = "O" Then vList("CPDType") = "O"
            dts.DisplayGrid(0).Populate(DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactCPDDetails, mvContactInfo.ContactNumber, vList))
            dts.DisplayGrid(0).ShowIfEmpty = True
            vList = New ParameterList(True, True)
            If vCPDPointNumber > 0 Then
              vList.IntegerValue("ContactCPDPointNumber") = vCPDPointNumber
              mvDataSet3 = DataHelper.GetContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactCPDPointDocuments, mvContactInfo.ContactNumber, vList)
            Else
              vList.IntegerValue("ContactCPDCycleNumber") = SelectRowItemNumber(pDataRow, "CopyContactCPDCycleNumber")
              If vCPDPeriodNumber > 0 Then vList.IntegerValue("ContactCPDPeriodNumber") = vCPDPeriodNumber
              mvDataSet3 = DataHelper.GetContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactCPDCycleDocuments, mvContactInfo.ContactNumber, vList)
            End If
            dts.DisplayGrid(1).Populate(mvDataSet3)
            dts.DisplayGrid(1).ShowIfEmpty = True
            mvDocumentMenu = New DocumentMenu(Me)
            If vCPDPointNumber > 0 Then
              mvDocumentMenu.DocumentType = BaseDocumentMenu.DocumentTypes.CPDPointDocuments
              mvDocumentMenu.CPDPointNumber = vCPDPointNumber
            Else
              mvDocumentMenu.DocumentType = BaseDocumentMenu.DocumentTypes.CPDCycleDocuments
              mvDocumentMenu.CPDPeriodNumber = vCPDPeriodNumber
            End If
            mvDocumentMenu.SetVisibleItems(New System.ComponentModel.CancelEventArgs)
            dts.DisplayGrid(1).ContextMenuStrip = mvDocumentMenu
            If dts.DisplayGrid(1).DataRowCount > 0 Then ProcessSubRowSelection(dts.DisplayGrid(1), 0, 0)
            vTabCount = 3

          Case CareServices.XMLContactDataSelectionTypes.xcdtContactCancelledProvisionalTrans
            vNumber = SelectRowItemNumber(pDataRow, "BatchNumber")
            Dim vTNumber As Integer = SelectRowItemNumber(pDataRow, "TransactionNumber")
            dts.DisplayGrid(0).Populate(DataHelper.GetTransactionData(CareServices.XMLTransactionDataSelectionTypes.xtdtTransactionAnalysis, vNumber, vTNumber))
            vTabCount = 2
            mvFinancialMenu.SetContext(mvDataType, GetDataRow(mvDataSet, pDataRow), mvContactInfo, False)
          Case CareServices.XMLContactDataSelectionTypes.xcdtContactMeetings
            vNumber = SelectRowItemNumber(pDataRow, "MeetingNumber")
            mvMeetingMenu.MeetingNumber = vNumber
            Dim vCLNumber As Integer = SelectRowItemNumber(pDataRow, "CommunicationsLogNumber")
            Dim vAgenda As String = ""
            Dim vList As New ParameterList(True)
            vList.IntegerValue("MeetingNumber") = vNumber
            Dim vTable As DataTable = DataHelper.SelectContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactMeetings, mvContactInfo.ContactNumber, vList)
            If vTable IsNot Nothing AndAlso vTable.Rows.Count = 1 Then
              vAgenda = vTable.Rows(0).Item("Agenda").ToString
            End If
            mvDataSet3 = DataHelper.GetDocumentData(CareServices.XMLDocumentDataSelectionTypes.xddtDocumentSubjects, vCLNumber)
            dts.DisplayGrid(1).Populate(mvDataSet3)
            dts.DisplayGrid(1).ShowIfEmpty = True
            dts.DisplayGrid(2).Populate(GetAgenda(vAgenda))
            dts.DisplayGrid(2).ShowIfEmpty = True
            mvDataSet2 = DataHelper.GetMeetingData(CType(CareNetServices.XMLDocumentDataSelectionTypes.xddtMeetingLinks, CareServices.XMLDocumentDataSelectionTypes), vNumber)
            dts.DisplayGrid(0).Populate(mvDataSet2)
            dts.DisplayGrid(0).ShowIfEmpty = True
            If dts.DisplayGrid(0).DataRowCount > 0 Then ProcessSubRowSelection(dts.DisplayGrid(0), 0, 0)
            vTabCount = 4
          Case CareServices.XMLContactDataSelectionTypes.xcdtContactSurveys
            Dim vList As New ParameterList(True)
            Dim vRow As DataRow = GetDataRow(mvDataSet, pDataRow)
            vTabCount = 2
            vList("ContactSurveyNumber") = vRow.Item("ContactSurveyNumber").ToString
            mvDataSet2 = DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactSurveyResponses, mvContactInfo.ContactNumber, vList)
            dts.DisplayGrid(0).Populate(mvDataSet2)
            dts.DisplayGrid(0).DisplayTitle = "Responses"
            dts.GridContextMenuStrip(0) = dgr0MenuStrip
          Case CareServices.XMLContactDataSelectionTypes.xcdtContactRegisteredUsers
            If dgr.DataRowCount > 0 Then
              dgrMenuNew.Available = False
            End If
            vTabCount = 0
          Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactServiceBookings
            mvServiceBookingMenu.SetContext(mvDataType, GetDataRow(mvDataSet, pDataRow))
            vNumber = SelectRowItemNumber(pDataRow, "BatchNumber")
            Dim vTNumber As Integer = SelectRowItemNumber(pDataRow, "TransactionNumber")
            Dim vList As New ParameterList(True, True)
            Dim vRow As DataRow = GetDataRow(mvDataSet, pDataRow)
            vList("ServiceBookingNumber") = vRow("ServiceBookingNumber").ToString
            dts.DisplayGrid(0).ShowIfEmpty = True
            vRow = Nothing
            If vTNumber > 0 AndAlso vNumber > 0 Then
              mvDataSet2 = DataHelper.GetTransactionData(CareNetServices.XMLTransactionDataSelectionTypes.xtdtServiceBookingDetails, vNumber, vTNumber, vList)
              dts.DisplayGrid(0).Populate(mvDataSet2)
              vRow = GetDataRow(mvDataSet2, dts.DisplayGrid(0).CurrentRow)
            Else
              dts.DisplayGrid(0).Clear()
              dts.DisplayGrid(0).DisplayTitle = "Related Financial Data"
            End If
            mvTransactionLinkMenu.SetContext(mvContactInfo, vRow)
            vTabCount = 2

          Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactLoans
            vNumber = SelectRowItemNumber(pDataRow, "PaymentPlanNumber")
            dts.DisplayGrid(0).Populate(DataHelper.GetPaymentPlanData(CareServices.XMLPaymentPlanDataSelectionTypes.xpdtPaymentPlanDetails, vNumber))
            mvDataSet3 = DataHelper.GetPaymentPlanData(CareServices.XMLPaymentPlanDataSelectionTypes.xpdtPaymentPlanPayments, vNumber)
            dts.DisplayGrid(1).Populate(mvDataSet3)
            dts.DisplayGrid(2).Populate(DataHelper.GetLoanData(CareNetServices.XMLLoanDataSelectionTypes.xldstInterestRates, vNumber))
            vTabCount = 4
            mvFinancialMenu.SetContext(mvDataType, GetDataRow(mvDataSet, pDataRow), mvContactInfo, False)

          Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactExamCertificates
            Dim vTempList As New ParameterList(True, True)
            vTempList.Add("ContactExamCertId", SelectRowItemNumber(pDataRow, "ContactExamCertId"))
            dts.DisplayGrid(0).Populate(DataHelper.GetContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactExamCertificateItems, mvContactInfo.ContactNumber, vTempList))
            dts.DisplayGrid(0).ShowIfEmpty = True
            dts.DisplayGrid(0).DisplayTitle = "Attributes"
            vTabCount = 2
            vTempList = New ParameterList(True, True)
            vTempList.Add("ContactExamCertId", dgr.GetValue(dgr.CurrentDataRow, "ContactExamCertId"))
            Dim vQueuedReprints As DataSet = DataHelper.GetContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactExamCertificateReprints,
                                                                       mvContactInfo.ContactNumber,
                                                                       vTempList)
            If vQueuedReprints IsNot Nothing AndAlso
               (Not vQueuedReprints.Tables.Contains("DataRow") OrElse
                vQueuedReprints.Tables("DataRow").Rows.Count < 1) Then
              mvContactExamCertificateMenu.ItemsEnabled = True
            End If

          Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactCPDPointsWithoutCycle
            Dim vList As New ParameterList(True)
            Dim vCPDPointNumber As Integer = SelectRowItemNumber(pDataRow, "ContactCPDPointNumber")
            If vCPDPointNumber > 0 Then vList.IntegerValue("ContactCPDPointNumber") = SelectRowItemNumber(pDataRow, "ContactCPDPointNumber")
            mvDataSet2 = DataHelper.GetContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactCPDPointDocuments, mvContactInfo.ContactNumber, vList)
            dts.DisplayGrid(0).Populate(mvDataSet2)
            dts.DisplayGrid(0).ShowIfEmpty = True
            mvDocumentMenu = New DocumentMenu(Me)
            mvDocumentMenu.DocumentType = BaseDocumentMenu.DocumentTypes.CPDPointDocuments
            mvDocumentMenu.CPDPointNumber = vCPDPointNumber
            mvDocumentMenu.SetVisibleItems(New System.ComponentModel.CancelEventArgs)
            dts.DisplayGrid(0).ContextMenuStrip = mvDocumentMenu
            If dts.DisplayGrid(0).DataRowCount > 0 Then ProcessSubRowSelection(dts.DisplayGrid(0), 0, 0)
            vTabCount = 2

          Case Else
            vTabCount = 0
        End Select
      Else
        dpl.Clear()
        If mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactCustomFormData Then ProcessNew()
        SetFinancialMenuToolbar(Nothing)
        If mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactDocuments Then
          Dim e As New System.ComponentModel.CancelEventArgs
          mvDocumentMenu.SetVisibleItems(e)
          dgr.SetToolBarVisible()
        End If
      End If

      Select Case mvDataType
        Case CareServices.XMLContactDataSelectionTypes.xcdtContactCustomFormData,
         CareServices.XMLContactDataSelectionTypes.xcdtContactAppointments,
         CareServices.XMLContactDataSelectionTypes.xcdtContactScores,
         CareServices.XMLContactDataSelectionTypes.xcdtContactDashboard,
         CType(CareServices.XMLContactDataSelectionTypes.xcdtContactPositionLinks + 2, CareServices.XMLContactDataSelectionTypes) ' TODO: The new enum xcdtContactNetwork is not implemented in vb code
          'Do nothing
        Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactCreditCustomers
          dts.SetTabPages(vTabCount, (dts.DisplayGrid(0).DataRowCount = 0))
          If dgr.DataRowCount > pRow Then
            ProcessSubRowSelection(dts.DisplayGrid(0), 0, 0)
          End If
          dgr0MenuNew.Visible = dgrMenuEdit.Available       'There must be an item at the top level and we must be able to edit it in order to create a sub-item
          dgr0MenuEdit.Visible = dgrMenuEdit.Available AndAlso dts.DisplayGrid(0).DataRowCount > 0
          dgr1MenuNew.Visible = dgrMenuEdit.Available       'There must be an item at the top level and we must be able to edit it in order to create a sub-item
          dgr1MenuEdit.Visible = dgrMenuEdit.Available AndAlso dts.DisplayGrid(1).DataRowCount > 0
        Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactExamSummary, CareNetServices.XMLContactDataSelectionTypes.xcdtContactExamDetails
          dgr0MenuNew.Visible = False
          dgr0MenuEdit.Visible = False
          dgr1MenuNew.Visible = False
          dgr1MenuEdit.Visible = False
          dgr2MenuNew.Visible = False
          dgr2MenuEdit.Visible = False
          'dts.SetTabPages(vTabCount, True, False)
        Case CareServices.XMLContactDataSelectionTypes.xcdtContactFundraising
          dts.SetTabPages(vTabCount, dts.DisplayGrid(0).DataRowCount = 0)
          If dgr.DataRowCount > pRow Then
            mvActionMenu.AddParameters("FundraisingRequestNumber", dgr.GetValue(dgr.CurrentDataRow, "FundraisingRequestNumber"))
            mvActionMenu.AddParameters("Logname", dgr.GetValue(dgr.CurrentDataRow, "Logname"))
            If AppValues.ControlValue(AppValues.ControlValues.fundraising_status).Length > 0 Then ProcessSubRowSelection(dts.DisplayGrid(0), 0, 0)
            ProcessSubRowSelection(dts.DisplayGrid(4), 0, 0)
          End If
          dgr.SetToolBarVisible()
        Case CareServices.XMLContactDataSelectionTypes.xcdtContactLegacyBequests
          dgr0MenuNew.Visible = False
          dgr0MenuEdit.Visible = dgrMenuEdit.Available AndAlso dts.DisplayGrid(0).DataRowCount > 0
          dgr1MenuNew.Visible = dgrMenuEdit.Available       'There must be an item at the top level and we must be able to edit it in order to create a sub-item
          dgr1MenuEdit.Visible = dgrMenuEdit.Available AndAlso dts.DisplayGrid(1).DataRowCount > 0
          dts.SetTabPages(vTabCount)
        Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactSurveys
          dgr0MenuNew.Visible = False
          dgr0MenuEdit.Visible = True
          dts.SetTabPages(vTabCount)
        Case CareServices.XMLContactDataSelectionTypes.xcdtContactRegisteredUsers
          dgrMenuNew.Visible = dgrMenuNew.Available AndAlso dgr.DataRowCount = 0
          dts.SetTabPages(vTabCount)
        Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactPurchaseOrders
          dgr1MenuNew.Visible = True
          dgr1MenuEdit.Visible = dts.DisplayGrid(1).DataRowCount > 0
          dts.SetTabPages(vTabCount)
        Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactPurchaseOrders
          dgr1MenuNew.Visible = True
          dgr1MenuEdit.Visible = dts.DisplayGrid(1).DataRowCount > 0
          dts.SetTabPages(vTabCount)
        Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactDocuments
          Dim vVisible As Boolean = (DataHelper.UserInfo.AccessLevel > UserInfo.UserAccessLevel.ualReadOnly)    'Ensure menus not displayed for read-only user
          dgr0MenuNew.Visible = dgrMenuEdit.Available AndAlso vVisible       'There must be an item at the top level and we must be able to edit it in order to create a sub-item
          dgr0MenuEdit.Visible = dgrMenuEdit.Available AndAlso dts.DisplayGrid(0).DataRowCount > 0 AndAlso vVisible
          dgr1MenuNew.Visible = dgrMenuEdit.Available AndAlso vVisible       'There must be an item at the top level and we must be able to edit it in order to create a sub-item
          dgr1MenuEdit.Visible = dgrMenuEdit.Available AndAlso dts.DisplayGrid(1).DataRowCount > 0 AndAlso vVisible
          dgr2MenuNew.Visible = dgrMenuEdit.Available AndAlso vVisible  'There must be an item at the top level and we must be able to edit it in order to create a sub-item
          dgr2MenuEdit.Visible = dgrMenuEdit.Available AndAlso dts.DisplayGrid(2).DataRowCount > 0 AndAlso vVisible
          dts.SetTabPages(vTabCount)
        Case CareServices.XMLContactDataSelectionTypes.xcdtContactViewOrganisations, CareNetServices.XMLContactDataSelectionTypes.xcdtContactPositions
          dgr0MenuNew.Visible = dgrMenuEdit.Available       'There must be an item at the top level and we must be able to edit it in order to create a sub-item
          dgr0MenuEdit.Visible = dgrMenuEdit.Available AndAlso dts.DisplayGrid(0).DataRowCount > 0
          dgr1MenuNew.Visible = dgrMenuEdit.Available       'There must be an item at the top level and we must be able to edit it in order to create a sub-item
          dgr1MenuEdit.Visible = dgrMenuEdit.Available AndAlso dts.DisplayGrid(1).DataRowCount > 0
          dgr2MenuNew.Visible = dgrMenuEdit.Available       'There must be an item at the top level and we must be able to edit it in order to create a sub-item
          dgr2MenuEdit.Visible = dgrMenuEdit.Available AndAlso dts.DisplayGrid(2).DataRowCount > 0
          dgr3MenuNew.Visible = dgrMenuEdit.Available       'There must be an item at the top level and we must be able to edit it in order to create a sub-item
          dgr3MenuEdit.Visible = dgrMenuEdit.Available AndAlso dts.DisplayGrid(3).DataRowCount > 0
          dgr4MenuNew.Visible = dgrMenuEdit.Available       'There must be an item at the top level and we must be able to edit it in order to create a sub-item
          dgr4MenuEdit.Visible = dgrMenuEdit.Available AndAlso dts.DisplayGrid(4).DataRowCount > 0
          dgr5MenuNew.Visible = dgrMenuEdit.Available       'There must be an item at the top level and we must be able to edit it in order to create a sub-item
          dgr5MenuEdit.Visible = dgrMenuEdit.Available AndAlso dts.DisplayGrid(5).DataRowCount > 0
          dgr5MenuDelete.Visible = dgrMenuEdit.Available AndAlso dts.DisplayGrid(5).DataRowCount > 0  'If able to edit top level item, then allow delete sub-item
          dts.SetTabPages(vTabCount)
        Case Else
          dgr0MenuNew.Visible = dgrMenuEdit.Available       'There must be an item at the top level and we must be able to edit it in order to create a sub-item
          dgr0MenuEdit.Visible = dgrMenuEdit.Available AndAlso dts.DisplayGrid(0).DataRowCount > 0
          dgr1MenuNew.Visible = dgrMenuEdit.Available       'There must be an item at the top level and we must be able to edit it in order to create a sub-item
          dgr1MenuEdit.Visible = dgrMenuEdit.Available AndAlso dts.DisplayGrid(1).DataRowCount > 0
          dgr2MenuNew.Visible = dgrMenuEdit.Available       'There must be an item at the top level and we must be able to edit it in order to create a sub-item
          dgr2MenuEdit.Visible = dgrMenuEdit.Available AndAlso dts.DisplayGrid(2).DataRowCount > 0
          dts.SetTabPages(vTabCount)
      End Select
      If mvDataType <> CareNetServices.XMLContactDataSelectionTypes.xcdtContactMeetings AndAlso dgr0MenuStrip IsNot Nothing Then
        'Ensure that if we have previously been to the meetings links grid, we remove the additional items that may have been added
        If dts.DisplayGrid(0).DataRowCount > 0 AndAlso mvDgr0IndependentItemNames IsNot Nothing Then dgr0MenuStripBuilder()
      End If
      If mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactInformation AndAlso
        AppValues.ConfigurationOption(AppValues.ConfigurationOptions.cd_high_profile_links) AndAlso
        AppValues.ConfigurationOption(AppValues.ConfigurationOptions.cd_default_relationship_tab) AndAlso
        (dts.DisplayGrid(3).DataRowCount > 0 OrElse dts.DisplayGrid(3).ShowIfEmpty) Then dts.SelectTab(4)
      If vTabCount >= 2 Then dts.DisplayGrid(0).SetToolBarVisible()
      If vTabCount >= 3 Then dts.DisplayGrid(1).SetToolBarVisible()
      If vTabCount >= 4 Then dts.DisplayGrid(2).SetToolBarVisible()
      If vTabCount >= 5 Then dts.DisplayGrid(3).SetToolBarVisible()
      If vTabCount >= 6 Then dts.DisplayGrid(4).SetToolBarVisible()
      If vTabCount >= 7 Then dts.DisplayGrid(5).SetToolBarVisible()
      dts.TabDock = DockStyle.Fill
      splTab.Show()
      dpl.Show()
      splTab.BringToFront()
      dpl.AutoScroll = True
      If Not dpl.HasChildren AndAlso dpl.Height = 0 Then
        splTab.Hide()
        dpl.Hide()
      End If
    Finally
      dts.DisplayGrid(3).ResumeLayout()
      dts.DisplayGrid(2).ResumeLayout()
      dts.DisplayGrid(1).ResumeLayout()
      dts.DisplayGrid(0).ResumeLayout()
      dts.ResumeLayout()
      dpl.ResumeLayout()
      Me.ResumeLayout()
    End Try
  End Sub

  Private Sub SetFinancialMenuToolbar(ByVal pDataRow As DataRow)
    If mvDataType <> CareServices.XMLContactDataSelectionTypes.xcdtContactCustomFormData AndAlso mvFinancialMenu IsNot Nothing Then
      mvFinancialMenu.SetContext(mvDataType, pDataRow, mvContactInfo, mvReadOnlyPage)
      Dim e As New System.ComponentModel.CancelEventArgs
      mvFinancialMenu.SetVisibleItems(e)
      dgr.SetToolBarVisible()
    End If
  End Sub

  Private Sub ShowDocumentDetails(ByVal pNumber As Integer)
    mvDocumentMenu.DocumentNumber = pNumber
    mvContactInfo.SelectedDocumentNumber = pNumber
    dts.TabPageText(0) = ControlText.TbpPrecis
    dts.SetText(DataHelper.GetDocumentPrecis(pNumber))
    mvDataSet2 = DataHelper.GetDocumentData(CareServices.XMLDocumentDataSelectionTypes.xddtDocumentSubjects, pNumber)
    dts.DisplayGrid(0).Populate(mvDataSet2)
    mvDataSet3 = DataHelper.GetDocumentData(CareServices.XMLDocumentDataSelectionTypes.xddtDocumentLinks, pNumber)
    dts.DisplayGrid(1).Populate(mvDataSet3)
    mvDocumentMenu.SetNotifyProcessed(dts.DisplayGrid(1))
    dts.DisplayGrid(2).MaxGridRows = DisplayTheme.HistoryMaxGridRows
    dts.DisplayGrid(2).Populate(DataHelper.GetDocumentData(CareServices.XMLDocumentDataSelectionTypes.xddtDocumentHistory, pNumber))

  End Sub

  Private Sub ShowContactMailingDetails(ByVal pNumber As Integer, ByVal pDataRow As Integer)
    Dim vList As New ParameterList(True)
    vList.IntegerValue("AddressNumber") = pNumber
    dts.DisplayGrid(0).Populate(DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactAddressPositionAndOrganisation, mvContactInfo.ContactNumber, vList), "OrganisationNumber <> " & mvContactInfo.ContactNumber)
    If dts.DisplayGrid(0).DisplayTitle.Length = 0 Then dts.DisplayGrid(0).DisplayTitle = "Position and Organisation"
    ' Need to see if we can display the  menu. Best way to do this is by not setting the context if we cant show the menu item
    Dim vCommsNumber As Integer = SelectRowItemNumber(pDataRow, "CommunicationNumber")
    If vCommsNumber > 0 Then 'Communication number will always be set for a Contact Emailings Link record
      vList.IntegerValue("CommunicationNumber") = vCommsNumber
      vList.IntegerValue("MailingNumber") = SelectRowItemNumber(pDataRow, "MailingNumber")
      dts.DisplayGrid(1).Populate(CareServices.XMLContactDataSelectionTypes.xcdtContactEmailingsLinks, mvContactInfo.ContactNumber, vList)
    Else
      dts.DisplayGrid(1).Clear()
    End If
    If ProcessContactMailingsMenu(GetDataRow(mvDataSet, pDataRow)) Then
      mvViewMailingDocumentMenu.SetContext(mvContactInfo, GetDataRow(mvDataSet, pDataRow))
    Else
      mvViewMailingDocumentMenu.SetContext(mvContactInfo, Nothing)
    End If
  End Sub
  Private Function SelectRowItem(ByVal pRow As Integer, ByVal pItem As String) As String
    If mvDataSet.Tables.Contains("DataRow") Then
      Dim vRow As DataRow = mvDataSet.Tables("DataRow").Rows(pRow)
      Return vRow.Item(pItem).ToString
    End If
    Return ""
  End Function

  Private Function SelectRowItemNumber(ByVal pRow As Integer, ByVal pItem As String) As Integer
    If mvDataSet.Tables.Contains("DataRow") Then
      Dim vRow As DataRow = mvDataSet.Tables("DataRow").Rows(pRow)
      Return IntegerValue(vRow.Item(pItem).ToString)
    End If
  End Function

  Private Function SelectSubRowItem(ByVal pDataSet As DataSet, ByVal pRow As Integer, ByVal pItem As String) As String
    If pDataSet.Tables.Contains("DataRow") Then
      Dim vRow As DataRow = pDataSet.Tables("DataRow").Rows(pRow)
      Return vRow.Item(pItem).ToString
    End If
    Return ""
  End Function

  Private Function SelectSubRowItemNumber(ByVal pDataSet As DataSet, ByVal pRow As Integer, ByVal pItem As String) As Integer
    If pDataSet.Tables.Contains("DataRow") Then
      Dim vRow As DataRow = pDataSet.Tables("DataRow").Rows(pRow)
      Return IntegerValue(vRow.Item(pItem).ToString)
    End If
  End Function

  Private Sub sel_CanCustomise(ByVal Sender As Object, ByVal pResult As String) Handles sel.CanCustomise
    Dim vParams As New ParameterList(True)
    vParams.Add("SelectionPages", "Y")
    vParams("DataSelectionType") = sel.DataSelectionType.ToString
    vParams("ParameterName") = mvContactInfo.ContactGroupParameterName
    vParams("ParameterValue") = mvContactInfo.ContactGroup
    If Me.ContactInfo.ContactTypeCode <> "O" Then 'Therefore ViewInContactCard Organisation Groups needs to be listed
      vParams.Add("IncludeGroupsFromContactCard", "Y")
    End If
    Dim vDisplayList As New frmDisplayList(frmDisplayList.ListUsages.CustomiseDisplayList, vParams)
    If vDisplayList.ShowDialog() = DialogResult.OK Then
      Dim vEntityGroup As EntityGroup = DataHelper.ContactAndOrganisationGroups(mvContactInfo.ContactGroup)
      vEntityGroup.ResetSelectionPages()
      sel.Init(mvContactInfo, True)
    End If

  End Sub

  Private Sub sel_ContactTabSelected(ByVal pSender As Object, ByVal pType As CareServices.XMLContactDataSelectionTypes, ByVal pCustomForm As Integer, ByVal pGroupID As String, ByVal pReadOnlyPage As Boolean) Handles sel.ContactTabSelected
    Dim vCursor As New BusyCursor
    Try
      If mvAccessibilityRoleReset Then
        splRight.AccessibleRole = mvRightSplitterAccessibleRole
        mvAccessibilityRoleReset = False
      End If
      If Not mvInitialised Then
        'Get here when the card is first opened due to an after_select on the TabSelector TreeView
        'mvDataType should already be set but if the first form is a custom one it won't be set
        mvCustomFormNumber = pCustomForm
        mvGroupID = pGroupID
        mvReadOnlyPage = pReadOnlyPage
      Else
        If mvDataType <> pType OrElse
          ((mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactCustomFormData) And (pCustomForm <> mvCustomFormNumber)) OrElse
          ((mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactCategories) And (pGroupID <> mvGroupID)) OrElse
          ((mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactLinksTo) And (pGroupID <> mvGroupID)) OrElse
          ((mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactViewOrganisations) And (pGroupID <> mvGroupID)) Then
          mvDataType = pType
          mvCustomFormNumber = pCustomForm
          mvGroupID = pGroupID
          mvReadOnlyPage = pReadOnlyPage
          RefreshCard()
        End If
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Private Sub PopulateDisplayGrid(ByVal sender As Object, ByVal pDGR As DisplayGrid, ByVal pPanelItem As PanelItem)
    Dim vList As New ParameterList(True)
    vList.IntegerValue("CustomForm") = mvCustomFormNumber
    vList("Detail") = pPanelItem.AttributeName
    If pPanelItem.ControlType = PanelItem.ControlTypes.ctGrid OrElse pPanelItem.ControlType = PanelItem.ControlTypes.ctSpreadSheet Then
      vList("Values") = dgr.GetRowCSValues(mvSelectedRow)
    End If
    pDGR.Populate(CareServices.XMLContactDataSelectionTypes.xcdtContactCustomFormData, mvContactInfo.ContactNumber, vList)
  End Sub

  Private Sub dgr_PrintParameters(ByVal pSender As Object, ByRef pJobName As String) Handles dgr.GetPrintParameters, dts.GetPrintParameters
    Dim vPrintCaption As New StringBuilder
    If mvContactInfo IsNot Nothing Then
      vPrintCaption.Append(mvContactInfo.ContactName)
      vPrintCaption.Append(" - ")
      vPrintCaption.Append(sel.SelectedNodeText)
      If pSender IsNot dgr Then
        vPrintCaption.Append(" (")
        vPrintCaption.Append(dts.GridTabPageText(DirectCast(pSender, DisplayGrid)))
        vPrintCaption.Append(")")
      End If
      pJobName = vPrintCaption.ToString
    End If
  End Sub

  Private Sub dgr_DocumentSelected(ByVal pSender As Object, ByVal pRow As Integer, ByVal pDocumentNumber As Integer) Handles dgr.DocumentSelected, dts.DocumentSelected
    Try
      If AllowEditDocument() Then
        FormHelper.EditDocument(pDocumentNumber, Me, Nothing)
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
  Private Sub dgr_RowSelected(ByVal pSender As Object, ByVal pRow As Integer, ByVal pDataRow As Integer) Handles dgr.RowSelected, dts.RowSelected
    Try
      If pSender Is dgr Then
        ProcessRowSelection(pRow, pDataRow)
      Else
        ProcessSubRowSelection(pSender, pRow, pDataRow)
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
  Private Sub dgr_ContactSelected(ByVal pSender As Object, ByVal pRow As Integer, ByVal pContactNumber As Integer) Handles dgr.ContactSelected, dts.ContactSelected
    If mvDataType = CareNetServices.XMLContactDataSelectionTypes.xcdtContactLegacyBequests Then
      'Don't select contact as the contact is not visible on the grid
    Else
      ContactSelectedHandler(pSender, pContactNumber)
    End If
  End Sub

  Private Sub dgr_ActionSelected(ByVal pSender As Object, ByVal pRow As Integer, ByVal pNumber As Integer) Handles dgr.ActionSelected, dts.ActionSelected
    If pSender Is dgr Then
      ActionSelectedHandler(pSender, pNumber)
    Else
      Try
        FormHelper.EditAction(pNumber, Me, Nothing)
      Catch vException As Exception
        DataHelper.HandleException(vException)
      End Try
    End If
  End Sub

  Private Sub dgr_EventSelected(ByVal pSender As Object, ByVal pRow As Integer, ByVal pEventNumber As Integer) Handles dgr.EventSelected, dts.EventSelected
    Try
      FormHelper.ShowEventIndex(pEventNumber)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub dgr_ExamUnitSelected(ByVal pSender As Object, ByVal pRow As Integer, ByVal pUnitNumber As Integer) Handles dgr.ExamUnitSelected, dts.ExamUnitSelected
    Try
      FormHelper.ShowExamIndex(pUnitNumber, "U")
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub dgr_ExamCentreSelected(ByVal pSender As Object, ByVal pRow As Integer, ByVal pUnitNumber As Integer) Handles dgr.ExamCentreSelected, dts.ExamCentreSelected
    Try
      FormHelper.ShowExamIndex(pUnitNumber, "N")
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub dgr_FundraisingRequestSelected(ByVal pSender As Object, ByVal pRow As Integer, ByVal pFundraisingRequestNumber As Integer) Handles dgr.DocumentFundraisingRequestSelected, dts.DocumentFundraisingRequestSelected
    Try
      'BR19023 BR19359
      Dim vList As New ParameterList(True)
      vList.Add("FundraisingRequestNumber", pFundraisingRequestNumber)

      Dim vContactNumber As Integer
      Dim vDataSet As New DataSet
      vDataSet = DataHelper.FindData(CareNetServices.XMLDataFinderTypes.xdftFundraisingRequestsFinder, vList)
      If vDataSet IsNot Nothing Then
        If vDataSet.Tables.Contains("DataRow") AndAlso vDataSet.Tables("DataRow").Columns.Contains("RequestNumber") Then
          Dim vDataTable As DataTable = mvDataSet.Tables("DataRow")
          Dim vRow As DataRow = vDataSet.Tables("DataRow").Rows(0)
          vContactNumber = IntegerValue(vRow.Item("ContactNumber").ToString)
        End If
      End If
      FormHelper.ShowCardIndex(CareNetServices.XMLContactDataSelectionTypes.xcdtContactFundraising, vContactNumber, False)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub dgr_ExamCentreUnitSelected(ByVal pSender As Object, ByVal pRow As Integer, ByVal pUnitNumber As Integer) Handles dgr.ExamCentreUnitSelected, dts.ExamCentreUnitSelected
    Try
      FormHelper.ShowExamIndex(pUnitNumber, "X")
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub dgr_WorkstreamSelected(ByVal sender As Object, ByVal pRow As Integer, ByVal pWorkstreamId As Integer) Handles dgr.WorkstreamSelected, dts.WorkstreamSelected
    Try
      Dim vWorkstreamGroup As String = String.Empty
      Dim vList As New ParameterList(True, True)
      vList.IntegerValue("WorkstreamId") = pWorkstreamId
      Dim vDT As DataTable = WorkstreamDataHelper.SelectWorkstreamData("", WorkstreamService.XMLDataSelectionTypes.WorkstreamDetails, vList)
      If vDT IsNot Nothing AndAlso vDT.Rows.Count > 0 Then
        vWorkstreamGroup = vDT.Rows(0).Item("WorkstreamGroup").ToString
      End If
      FormHelper.ShowWorkstreamIndex(vWorkstreamGroup, pWorkstreamId)
    Catch vEX As Exception
      DataHelper.HandleException(vEX)
    End Try
  End Sub

  Private Sub dgr_CPDCyclePeriodSelected(ByVal sender As Object, ByVal pRow As Integer, ByVal pCPDPeriodNumber As Integer) Handles dts.CPDCyclePeriodSelected
    Dim vContactNumber As Integer = mvContactInfo.ContactNumber
    Dim vList As New ParameterList(True, True)
    vList.IntegerValue("ContactCpdPeriodNumber") = pCPDPeriodNumber
    Dim vTable As DataTable = DataHelper.GetTableFromDataSet(DataHelper.FindData(CareNetServices.XMLDataFinderTypes.xdftCPDCyclePeriodFinder, vList))
    If vTable IsNot Nothing Then
      If vTable.Columns.Contains("ContactNumber") Then vContactNumber = IntegerValue(vTable.Rows(0).Item("ContactNumber").ToString)
    End If
    If vContactNumber <> mvContactInfo.ContactNumber Then
      'TODO: Ask to switch contact?
      FormHelper.ShowCardIndex(CareNetServices.XMLContactDataSelectionTypes.xcdtContactCPD, vContactNumber, False)
    Else
      sel.SetSelectionType(CDBNETCL.CareNetServices.XMLContactDataSelectionTypes.xcdtContactCPD)
    End If
    SelectRowItem("ContactCPDPeriodNumber", pCPDPeriodNumber)
  End Sub

  Private Sub dgr_CPDPointSelected(ByVal sender As Object, ByVal pRow As Integer, ByVal pCPDPointNumber As Integer) Handles dts.CPDPointSelected
    Dim vContactNumber As Integer = mvContactInfo.ContactNumber
    Dim vGotCyclePeriod As Boolean = False

    Dim vList As New ParameterList(True, True)
    vList.IntegerValue("ContactCpdPointNumber") = pCPDPointNumber
    Dim vTable As DataTable = DataHelper.GetTableFromDataSet(DataHelper.FindData(CareNetServices.XMLDataFinderTypes.xdftCPDPointFinder, vList))
    If vTable IsNot Nothing Then
      If vTable.Columns.Contains("ContactNumber") Then vContactNumber = IntegerValue(vTable.Rows(0).Item("ContactNumber").ToString)
      If vTable.Columns.Contains("ContactCpdPeriodNumber") AndAlso IntegerValue(vTable.Rows(0).Item("ContactCpdPeriodNumber").ToString) > 0 Then vGotCyclePeriod = True
    End If

    Dim vType As CareNetServices.XMLContactDataSelectionTypes = CareNetServices.XMLContactDataSelectionTypes.xcdtContactCPDPointsWithoutCycle
    If vGotCyclePeriod Then vType = CareNetServices.XMLContactDataSelectionTypes.xcdtContactCPD
    If vContactNumber <> mvContactInfo.ContactNumber Then
      'TODO: Ask to switch contact?
      FormHelper.ShowCardIndex(vType, vContactNumber, False)
    Else
      sel.SetSelectionType(vType)
    End If
    SelectRowItem("ContactCPDPointNumber", pCPDPointNumber)
  End Sub

  Private Sub dgr_CustomiseCardSet(ByVal pSender As Object, ByVal pDataSelectionType As Integer, ByVal pRevert As Boolean) Handles dgr.CustomiseCardSet, dts.CustomiseCardSet
    Dim vParams As New ParameterList(True)
    If pRevert Then
      DataHelper.RevertChanges(pDataSelectionType, mvContactInfo.ContactGroupParameterName, mvContactInfo.ContactGroup)
      RefreshCard()
    Else
      vParams("DataSelectionType") = pDataSelectionType.ToString
      If TypeOf (pSender) Is DisplayGrid Then
        Dim vGrid As DisplayGrid = DirectCast(pSender, DisplayGrid)
        Dim vCustomiseByGroup As Boolean = False
        If vGrid Is dgr Then
          vCustomiseByGroup = True
        Else
          Select Case mvDataType
            Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactAddresses
              vCustomiseByGroup = (vGrid Is dts.DisplayGrid(1))
            Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactCommsNumbers
              vCustomiseByGroup = (vGrid Is dts.DisplayGrid(0))
            Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactPositions, CareNetServices.XMLContactDataSelectionTypes.xcdtContactViewOrganisations
              If (vGrid Is dts.DisplayGrid(3)) OrElse (vGrid Is dts.DisplayGrid(4)) OrElse (vGrid Is dts.DisplayGrid(5)) Then
                'Position Actions OR Position Documents OR Position Timesheet
                vCustomiseByGroup = True
              End If
          End Select
        End If
        If vCustomiseByGroup Then
          vParams("ParameterName") = mvContactInfo.ContactGroupParameterName
          vParams("ParameterValue") = mvContactInfo.ContactGroup
        End If
      End If
      Dim vDisplayList As New frmDisplayList(frmDisplayList.ListUsages.CustomiseDisplayList, vParams)
      If vDisplayList.ShowDialog() = DialogResult.OK Then
        RefreshCard()
      End If
    End If
  End Sub

  Private Sub ContactSelectedHandler(ByVal pSender As Object, ByVal pContactNumber As Integer)
    Try
      'BR 10298 Only navigate if a different contact is selected
      If pContactNumber <> mvContactInfo.ContactNumber Then
        FormHelper.ShowContactCardIndex(pContactNumber)
      End If
    Catch vException As Exception
      If mvDataType = CareNetServices.XMLContactDataSelectionTypes.xcdtContactLinksTo Or mvDataType = CareNetServices.XMLContactDataSelectionTypes.xcdtContactLinksFrom Then  'BR19564
        ShowWarningMessage(mvContactInfo.ViewAccessMessage)
      Else
        DataHelper.HandleException(vException)
      End If
    End Try
  End Sub

  Private Sub ActionSelectedHandler(ByVal pSender As Object, ByVal pContactNumber As Integer)
    Try
      ShowMaintenanceForm(New frmCardMaintenance(Me,
                                                 mvContactInfo,
                                                 mvDataType,
                                                 mvDataSet,
                                                 True,
                                                 dgr.CurrentDataRow),
                          CareServices.XMLMaintenanceControlTypes.xmctActionTopic)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub CalendarDoubleClickedHandler(ByVal pType As CalendarView.CalendarItemTypes, ByVal pDescription As String, ByVal pStart As Date, ByVal pEnd As Date, ByVal pUniqueID As Integer)
    If pType = CalendarView.CalendarItemTypes.catOther Then
      'Need to handle separately as MainHelper.CalendarDoubleClicked will display for User Contact
      Dim vList As New ParameterList
      vList("AppointmentDesc") = pDescription
      vList("StartDate") = pStart.ToString
      vList("EndDate") = pEnd.ToString
      Dim vForm As New frmCardMaintenance(CareServices.XMLMaintenanceControlTypes.xmctAppointments, mvContactInfo.ContactNumber, Me, vList)
      vForm.Show()
    Else
      MainHelper.CalendarDoubleClicked(Me, pType, pDescription, pStart, pEnd, pUniqueID)
    End If
  End Sub

  Private Sub dgrMenuNewDocument_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgrMenuNewDocument.Click
    If dgr.CurrentRow >= 0 Then
      Dim vContactInfo As New ContactInfo(IntegerValue(dgr.GetValue(dgr.CurrentRow, "ContactNumber")))
      vContactInfo.RelatedContact = mvContactInfo
      FormHelper.NewDocument(Me, vContactInfo)
    End If
  End Sub

  Private Sub dgrMenuEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgrMenuEdit.Click
    HandleMenuClick(True, MenuSource.dgr)
  End Sub
  Private Sub dgr0MenuEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgr0MenuEdit.Click
    If ContactDataType = CareNetServices.XMLContactDataSelectionTypes.xcdtContactSurveys Then
      Dim vForm As New frmSurveyResponses(mvContactInfo.ContactNumber, IntegerValue(mvDataSet2.Tables("DataRow").Rows(dts.DisplayGrid(0).CurrentRow).Item("ContactSurveyNumber").ToString))
      vForm.ShowDialog()
      If vForm.DialogResult = System.Windows.Forms.DialogResult.OK Then
        ProcessRowSelection(dgr.CurrentRow, dgr.CurrentDataRow)
      End If
    Else
      HandleMenuClick(True, MenuSource.dgr0)
    End If
  End Sub
  Private Sub dgr1MenuEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgr1MenuEdit.Click
    HandleMenuClick(True, MenuSource.dgr1)
  End Sub
  Private Sub dgr2MenuEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgr2MenuEdit.Click
    HandleMenuClick(True, MenuSource.dgr2)
  End Sub
  Private Sub dgr3MenuEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgr3MenuEdit.Click
    HandleMenuClick(True, MenuSource.dgr3)
  End Sub
  Private Sub dgr4MenuEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgr4MenuEdit.Click
    HandleMenuClick(True, MenuSource.dgr4)
  End Sub
  Private Sub dgr5MenuEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgr5MenuEdit.Click
    HandleMenuClick(True, MenuSource.dgr5)
  End Sub
  Private Sub dplMenuEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dplMenuEdit.Click
    HandleMenuClick(True, MenuSource.dpl)
  End Sub
  Private Sub dgrMenuNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgrMenuNew.Click
    HandleMenuClick(False, MenuSource.dgr)
  End Sub
  Private Sub dgr0MenuNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgr0MenuNew.Click
    HandleMenuClick(False, MenuSource.dgr0)
  End Sub
  Private Sub dgr1MenuNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgr1MenuNew.Click
    HandleMenuClick(False, MenuSource.dgr1)
  End Sub
  Private Sub dgr2MenuNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgr2MenuNew.Click
    HandleMenuClick(False, MenuSource.dgr2)
  End Sub
  Private Sub dgr3MenuNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgr3MenuNew.Click
    HandleMenuClick(False, MenuSource.dgr3)
  End Sub
  Private Sub dgr4MenuNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgr4MenuNew.Click
    HandleMenuClick(False, MenuSource.dgr4)
  End Sub
  Private Sub dgr5MenuNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgr5MenuNew.Click
    HandleMenuClick(False, MenuSource.dgr5)
  End Sub
  Private Sub dplMenuNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dplMenuNew.Click
    HandleMenuClick(False, MenuSource.dpl)
  End Sub
  Private Sub actionMenuEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mvActionMenu.EditAction
    HandleActionMenuClick(sender, True)
  End Sub
  Private Sub actionMenuNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mvActionMenu.NewAction
    HandleActionMenuClick(sender, False)

  End Sub
  Private Sub dgr5MenuDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgr5MenuDelete.Click
    HandleMenuDeleteClick(MenuSource.dgr5)
  End Sub
  Private Sub HandleActionMenuClick(ByVal sender As System.Object, ByVal pEdit As Boolean)
    Dim vMenuSource As MenuSource = MenuSource.dgr
    Dim vItem As ToolStripMenuItem = TryCast(sender, ToolStripMenuItem)
    If vItem IsNot Nothing Then
      Dim vContextMenuStrip As ContextMenuStrip = TryCast(vItem.GetCurrentParent, ContextMenuStrip)
      If vContextMenuStrip IsNot Nothing Then
        If vContextMenuStrip.SourceControl Is dts.DisplayGrid(0) Then
          vMenuSource = MenuSource.actiondgr0
        ElseIf vContextMenuStrip.SourceControl Is dts.DisplayGrid(3) Then
          vMenuSource = MenuSource.dgr3
        ElseIf vContextMenuStrip.SourceControl Is dts.DisplayGrid(4) Then
          vMenuSource = MenuSource.actiondgr4
        Else
          Dim vActionMenu As BaseActionMenu = TryCast(vItem.GetCurrentParent, BaseActionMenu)
          If Not vActionMenu Is Nothing Then
            If vActionMenu.ActionType = BaseActionMenu.ActionTypes.PositionActions Then
              vMenuSource = MenuSource.dgr3
            End If
          End If
        End If
      End If
    End If
    HandleMenuClick(pEdit, vMenuSource)
  End Sub

  Private Enum MenuSource
    dgr
    dpl
    dgr0
    dgr1
    dgr2
    actiondgr0
    actiondgr4
    dplCustomise
    dplRevert
    dgr3
    dgr4
    dgr5
  End Enum

  Private Sub HandleMenuClick(ByVal pEdit As Boolean, ByVal pSource As MenuSource)
    Dim vForm As frmCardMaintenance = Nothing
    Dim vCursor As New BusyCursor
    Try
      Dim vEdit As Boolean = pEdit
      'If (vItem.Tag IsNot Nothing) AndAlso (mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactActions OrElse _
      '    mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactLegacyActions) Then
      '  vEdit = DirectCast(vItem.Tag, MenuToolbarCommand).CommandID = ActionMenu.ActionMenuItems.amiEdit
      'End If
      Dim vMaintenanceType As CareServices.XMLMaintenanceControlTypes
      If pSource = MenuSource.dgr Then
        If mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactCategories And mvGroupID.Length > 0 Then
          ShowDataSheet(Me, frmDataSheet.DataSheetTypes.dstActivities, mvContactInfo, "B", "", mvGroupID, sel.SelectedNodeText, True)
        ElseIf mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactLinksTo And mvGroupID.Length > 0 Then
          ShowDataSheet(Me, frmDataSheet.DataSheetTypes.dstRelationships, mvContactInfo, "B", "", mvGroupID, sel.SelectedNodeText, True)
        ElseIf mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactCPD Then
          ' Here we need to actually populate the card maintenance form with different data to the original grid on Card Set form
          Dim vDataSet As DataSet = DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactCPDCycles, mvContactInfo.ContactNumber)
          Dim vTable As DataTable = DataHelper.GetTableFromDataSet(vDataSet)
          If Not vTable Is Nothing AndAlso vTable.Rows.Count > 0 Then mvContactInfo.SetCPDValues(IntegerValue(dgr.GetValue(dgr.CurrentRow, "CopyContactCPDCycleNumber")), dgr.GetValue(dgr.CurrentRow, "CPDCycleType"))
          vForm = New frmCardMaintenance(Me, mvContactInfo, mvDataType, vDataSet, vEdit, 0)
        ElseIf mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactFundraising Then
          vForm = FundraisingActionsMaintenance(pSource, vEdit)
        Else
          If vEdit AndAlso (mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactCategories OrElse
             mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactLegacyAssets) Then
            If Not dgr.GetValue(dgr.CurrentRow, "Access").StartsWith("Y") Then
              ShowInformationMessage(InformationMessages.ImActivityOwnershipViolation)
              Exit Sub
            End If
          End If
          If mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactViewOrganisations Then
            Dim vParams As New ParameterList(True)
            vParams("OrganisationGroupCode") = mvGroupID
            vForm = New frmCardMaintenance(Me, mvContactInfo, mvDataType, mvDataSet, vEdit, dgr.CurrentDataRow, CareNetServices.XMLMaintenanceControlTypes.xmctNone, vParams)
          Else
            vForm = New frmCardMaintenance(Me, mvContactInfo, mvDataType, mvDataSet, vEdit, dgr.CurrentDataRow)
          End If
          If mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactPositions OrElse mvDataType = CareNetServices.XMLContactDataSelectionTypes.xcdtContactViewOrganisations Then
            AddHandler vForm.ShowApplicationParameters, AddressOf ShowApplicationParametersHandler
          End If
        End If
      ElseIf pSource = MenuSource.actiondgr0 OrElse pSource = MenuSource.actiondgr4 Then
        If mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactFundraising Then
          vForm = FundraisingActionsMaintenance(pSource, vEdit)
        End If
      ElseIf pSource = MenuSource.dpl Then
        vForm = New frmCardMaintenance(Me, mvContactInfo, mvDataType, mvDataSet, vEdit, dgr.CurrentRow)
      ElseIf pSource = MenuSource.dplCustomise Then
        Dim vParams As New ParameterList(True)
        Dim vTable As DataTable = mvDataSet.Tables("Column")
        For Each vRow As DataRow In vTable.Rows
          If vRow.Item("Name").ToString = "DataSelection" Then
            vParams("DataSelectionType") = vRow.Item("Value").ToString
            vParams("ParameterName") = mvContactInfo.ContactGroupParameterName
            vParams("ParameterValue") = mvContactInfo.ContactGroup
            vParams("ContactDetails") = "Y"
            Dim vDisplayList As New frmDisplayList(frmDisplayList.ListUsages.CustomiseDisplayList, vParams)
            vDisplayList.ShowDialog()
            Exit For
          End If
        Next
      ElseIf pSource = MenuSource.dplRevert Then
        If ShowQuestion(QuestionMessages.QmRevertModule, MessageBoxButtons.OKCancel) = DialogResult.OK Then
          Dim vParams As New ParameterList(True)
          Dim vTable As DataTable = mvDataSet.Tables("Column")
          For Each vRow As DataRow In vTable.Rows
            If vRow.Item("Name").ToString = "DataSelection" Then
              vParams("DataSelectionType") = vRow.Item("Value").ToString
              vParams.Add("AccessMethod", "S")
              vParams.Add(mvContactInfo.ContactGroupParameterName, mvContactInfo.ContactGroup)
              vParams.Add("Logname", DataHelper.UserInfo.Logname.ToString)
              vParams.Add("Department", DataHelper.UserInfo.Department.ToString)
              vParams.Add("Client", DataHelper.GetClientCode())
              vParams.Add("WebPageItemNumber", "")
              DataHelper.DeleteItem(CareNetServices.XMLMaintenanceControlTypes.xmctDisplayListItem, vParams)
              Exit For
            End If
          Next
        End If
      ElseIf pSource = MenuSource.dgr1 AndAlso mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactPositions AndAlso AppValues.ControlValue(AppValues.ControlTables.contact_controls, AppValues.ControlValues.position_activity_group).Length > 0 Then
        ShowPositionAdditionalData(PositionAdditionalData.Activities, Me, mvContactInfo, Nothing, False, mvGroupID, sel.SelectedNodeText)
      ElseIf pSource = MenuSource.dgr2 AndAlso mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactPositions AndAlso AppValues.ControlValue(AppValues.ControlTables.contact_controls, AppValues.ControlValues.position_relationship_group).Length > 0 Then
        ShowPositionAdditionalData(PositionAdditionalData.Relationships, Me, mvContactInfo, Nothing, False, mvGroupID, sel.SelectedNodeText)
      Else
        Dim vSubDataType As CareServices.XMLContactDataSelectionTypes
        Dim vDataSet As DataSet
        Dim vRow As Integer
        Dim vList As ParameterList = Nothing
        Select Case mvDataType
          Case CareServices.XMLContactDataSelectionTypes.xcdtContactPositions, CareNetServices.XMLContactDataSelectionTypes.xcdtContactViewOrganisations
            If pSource = MenuSource.dgr0 Then
              vSubDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactRoles
              vMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctRole
              vDataSet = mvDataSet2
              vRow = dts.DisplayGrid(0).CurrentRow
            ElseIf pSource = MenuSource.dgr1 Then
              vSubDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactPositionActivities
              vMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctPositionActivity
              vDataSet = mvDataSet3
              vRow = dts.DisplayGrid(1).CurrentRow
            ElseIf pSource = MenuSource.dgr2 Then
              vSubDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactPositionLinks
              vMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctPositionLinks
              vDataSet = mvDataSet4
              vRow = dts.DisplayGrid(2).CurrentRow
            ElseIf pSource = MenuSource.dgr3 Then 'Action
              vSubDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactPositionActions
              vMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctAction
              vDataSet = mvDataSet5
              vRow = dts.DisplayGrid(3).CurrentRow
              If dgr.CurrentDataRow >= 0 Then
                vList = New ParameterList()
                vList("ContactPositionNumber") = dgr.GetValue(dgr.CurrentDataRow, "ContactPositionNumber")
              End If
            ElseIf pSource = MenuSource.dgr4 Then 'Document
              'Handled in BaseDocumentMenu class 
            Else 'Timesheet
              vSubDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactPositionTimesheets
              vMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctTimesheet
              vDataSet = mvDataSet7
              vRow = dts.DisplayGrid(5).CurrentRow
              If dgr.CurrentDataRow >= 0 Then
                vList = New ParameterList()
                vList("RoleContactPositionNumber") = dgr.GetValue(dgr.CurrentDataRow, "ContactPositionNumber")
              End If
            End If
          Case CareServices.XMLContactDataSelectionTypes.xcdtContactAddresses
            vSubDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactAddressUsages
            vMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctAddressUsage
            vDataSet = mvDataSet3
            vRow = dts.DisplayGrid(1).CurrentRow
          Case CareServices.XMLContactDataSelectionTypes.xcdtContactCommsNumbers
            vSubDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactCommunicationUsages
            vMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctCommunicationUsage
            vDataSet = mvDataSet2
            vRow = dts.DisplayGrid(0).CurrentRow
            'Need a change in VB 6 code to access xcdtContactCommunicationHistory directly 
            'till then access it using <TODO> 
          Case CareServices.XMLContactDataSelectionTypes.xcdtContactDocuments, CType(CareServices.XMLContactDataSelectionTypes.xcdtContactPositionLinks + 3, CareServices.XMLContactDataSelectionTypes)
            vSubDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactDocuments
            If pSource = MenuSource.dgr0 Then
              vDataSet = mvDataSet2
              vRow = dts.DisplayGrid(0).CurrentRow
              vMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctDocumentTopic
            Else
              vDataSet = mvDataSet3
              vRow = dts.DisplayGrid(1).CurrentRow
              vMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctDocumentLink
            End If
          Case CType(CareNetServices.XMLContactDataSelectionTypes.xcdtContactMeetings, CareServices.XMLContactDataSelectionTypes)
            If dgr.CurrentDataRow >= 0 Then
              vList = New ParameterList()
              vList("MeetingNumber") = dgr.GetValue(dgr.CurrentDataRow, "MeetingNumber")
              vList.IntegerValue("CommunicationsLogNumber") = IntegerValue(dgr.GetValue(dgr.CurrentDataRow, "CommunicationsLogNumber"))
            End If
            If pSource = MenuSource.dgr0 Then
              vDataSet = mvDataSet2
              vRow = dts.DisplayGrid(0).CurrentRow
              vMaintenanceType = CType(CareNetServices.XMLMaintenanceControlTypes.xmctMeetingLinks, CareServices.XMLMaintenanceControlTypes)
            Else
              vDataSet = mvDataSet3
              vRow = dts.DisplayGrid(1).CurrentRow
              vMaintenanceType = CType(CareNetServices.XMLMaintenanceControlTypes.xmctMeetingTopic, CareServices.XMLMaintenanceControlTypes)
            End If
          Case CareServices.XMLContactDataSelectionTypes.xcdtContactActions, CareServices.XMLContactDataSelectionTypes.xcdtContactLegacyActions
            vSubDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactActions
            If pSource = MenuSource.dgr0 Then
              vDataSet = mvDataSet2
              vRow = dts.DisplayGrid(0).CurrentRow
              vMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctActionTopic
            Else
              vDataSet = mvDataSet3
              vRow = dts.DisplayGrid(1).CurrentRow
              vMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctActionLink
            End If
          Case CareServices.XMLContactDataSelectionTypes.xcdtContactLegacyBequests
            If pSource = MenuSource.dgr0 Then
              vSubDataType = CareServices.XMLContactDataSelectionTypes.xcdtNone
              vMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctLegacyBequestReceipt
              vDataSet = mvDataSet2
              vRow = dts.DisplayGrid(0).CurrentRow
            Else
              vSubDataType = CareServices.XMLContactDataSelectionTypes.xcdtNone
              vMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctLegacyBequestForecast
              vDataSet = mvDataSet3
              vRow = dts.DisplayGrid(1).CurrentRow
            End If
          Case Else
            vDataSet = Nothing
        End Select
        vForm = New frmCardMaintenance(Me, mvContactInfo, vSubDataType, vDataSet, vEdit, vRow, vMaintenanceType, vList)
      End If
      ShowMaintenanceForm(vForm, vMaintenanceType)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub
  Private Sub HandleMenuDeleteClick(ByVal pSource As MenuSource)
    Try

      If ConfirmDelete() Then
        Select Case mvDataType
          Case CareServices.XMLContactDataSelectionTypes.xcdtContactViewOrganisations, CareServices.XMLContactDataSelectionTypes.xcdtContactPositions
            If pSource = MenuSource.dgr5 Then  'Timesheet
              Dim vRow As Integer = dts.DisplayGrid(5).CurrentRow
              Dim vList As New ParameterList(True, True)
              vList("ContactPositionNumber") = dts.DisplayGrid(5).DataRows(vRow)("ContactPositionNumber").ToString()
              vList("TimesheetNumber") = dts.DisplayGrid(5).DataRows(vRow)("ContactTimesheetNumber").ToString()
              DataHelper.DeleteItem(CareNetServices.XMLMaintenanceControlTypes.xmctTimesheet, vList)
              RefreshData()
            End If
        End Select
      End If

    Catch vEX As Exception
      DataHelper.HandleException(vEX)
    End Try

  End Sub
  Private Sub ShowMaintenanceForm(ByVal pForm As frmCardMaintenance)
    ShowMaintenanceForm(pForm, CareNetServices.XMLMaintenanceControlTypes.xmctNone)
  End Sub

  Private Sub ShowMaintenanceForm(ByVal pForm As frmCardMaintenance, pMaintenanceType As CareNetServices.XMLMaintenanceControlTypes)
    If pForm IsNot Nothing Then
      mvCustomiseMenu = New CustomiseMenu
      Dim vMaintenanceType As CareServices.XMLMaintenanceControlTypes = If(pMaintenanceType <> CareNetServices.XMLMaintenanceControlTypes.xmctNone, pMaintenanceType, GetMaintenanceType(mvDataType, mvContactInfo.ContactType))
      If vMaintenanceType = CareNetServices.XMLMaintenanceControlTypes.xmctNone Then
        vMaintenanceType = pForm.MaintenanceType
      End If
      If vMaintenanceType <> CareNetServices.XMLMaintenanceControlTypes.xmctNone Then
        mvCustomiseMenu.SetContext(pForm, vMaintenanceType, mvContactInfo.ContactGroup)
        pForm.SetCustomiseMenu(mvCustomiseMenu)
        Dim vLocation As Point = splRight.PointToScreen(splRight.Location)
        Dim vSize As Size = splRight.Size
        pForm.SetInitialBounds(vLocation, vSize)
        If MDIForm IsNot Nothing Then
          Try
            If FormView = FormViews.Modern Then
              pForm.TopLevel = False
              pForm.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
              Dim vVisibleControls As New List(Of Control)
              For Each vControl As Control In Me.splBottom.Panel2.Controls
                vVisibleControls.Add(vControl)
                vControl.Visible = False
              Next vControl
              Me.splBottom.Panel2.Controls.Add(pForm)
              AddHandler pForm.FormClosed, Sub(sender As Object, e As FormClosedEventArgs)
                                             Me.splBottom.Panel2.Controls.Remove(pForm)
                                             For Each vControl As Control In vVisibleControls
                                               vControl.Visible = True
                                             Next vControl
                                             Me.splBottom.Panel1.Enabled = True
                                             Me.splTop.Panel1.Enabled = True
                                             Me.splBottom.Panel2.Focus()
                                           End Sub
              pForm.Dock = DockStyle.Fill
              Me.splBottom.Panel1.Enabled = False
              Me.splTop.Panel1.Enabled = False
            End If
            pForm.Show()
          Catch vEx As Exception
            DataHelper.HandleException(vEx)
          End Try
        Else
          pForm.Show(Me)
        End If
      End If
    End If
  End Sub

  Private Sub frmCardSet_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
    If mvContactInfo IsNot Nothing Then MainHelper.SetStatusContact(mvContactInfo, False)
  End Sub

  Private Sub frmCardSet_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
    If mvCareWebBrowser IsNot Nothing AndAlso mvCareWebBrowser.ConfirmNavigateAway Then
      If Not ConfirmCancel() Then e.Cancel = True
    End If
  End Sub
  Private Sub frmCardSet_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
    If mvContactInfo IsNot Nothing Then MainHelper.SetStatusContact(mvContactInfo, True)
    If mvActiveChildControl IsNot Nothing Then mvActiveChildControl.Focus()
  End Sub
  Private Sub frmCardSet_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Deactivate
    SetActiveChildControl()
  End Sub
  Private Sub frmCardSet_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
    If mvDocumentMenu IsNot Nothing Then mvDocumentMenu.Dispose()
    If mvActionMenu IsNot Nothing Then mvActionMenu.Dispose()
    If mvCustomPanel IsNot Nothing Then mvCustomPanel.Dispose()
    If mvFinancialMenu IsNot Nothing Then mvFinancialMenu.Dispose()
    If mvAnalysisFinancialMenu IsNot Nothing Then mvAnalysisFinancialMenu.Dispose()
    If mvTransactionLinkMenu IsNot Nothing Then mvTransactionLinkMenu.Dispose()
    If mvJournalLinkMenu IsNot Nothing Then mvJournalLinkMenu.Dispose()
    If mvPurchaseOrderMenu IsNot Nothing Then mvPurchaseOrderMenu.Dispose()
    If mvPurchaseOrderPaymentMenu IsNot Nothing Then mvPurchaseOrderPaymentMenu.Dispose()
    If mvContactEventDelegateMenu IsNot Nothing Then mvContactEventDelegateMenu.Dispose()
    If mvViewMailingDocumentMenu IsNot Nothing Then mvViewMailingDocumentMenu.Dispose()
    If mvCustomiseMenu IsNot Nothing Then mvCustomiseMenu.Dispose()
    If mvExamsCustomiseMenu IsNot Nothing Then mvExamsCustomiseMenu.Dispose()
    If mvPurchaseInvoiceChequeMenu IsNot Nothing Then mvPurchaseInvoiceChequeMenu.Dispose()
    If mvCareWebBrowser IsNot Nothing Then mvCareWebBrowser.Dispose()
    If mvPurchaseInvoiceMenu IsNot Nothing Then mvPurchaseInvoiceMenu.Dispose()
  End Sub

  Private Sub cmdDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
    'Only supported for custom data on this form
    Try
      'TODO Confirm cancel changes
      If Not ConfirmDelete() Then Exit Sub
      Dim vList As New ParameterList(True)
      vList.IntegerValue("ContactNumber") = mvContactInfo.ContactNumber
      vList.IntegerValue("CustomForm") = mvCustomFormNumber
      vList("IgnoreUnknownParameters") = "Y"
      mvCustomPanel.AddValuesToList(vList)
      DataHelper.DeleteItem(CareServices.XMLMaintenanceControlTypes.xmctCustomForm, vList)
    Catch vCareException As CareException
      If vCareException.ErrorNumber = CareException.ErrorNumbers.enOracleUserError OrElse
         vCareException.ErrorNumber = CareException.ErrorNumbers.enODBCUserDefinedError Then
        ShowInformationMessage(vCareException.Message)
      Else
        DataHelper.HandleException(vCareException)
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      RefreshData()
    End Try
  End Sub

  Private Sub cmdSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdSave.Click
    'Only supported for custom data on this form
    Try
      Dim vList As New ParameterList(True)
      vList.IntegerValue("ContactNumber") = mvContactInfo.ContactNumber
      vList.IntegerValue("CustomForm") = mvCustomFormNumber
      vList("IgnoreUnknownParameters") = "Y"
      If mvCustomPanel.AddValuesToList(vList, True, EditPanel.AddNullValueTypes.anvtAll) Then
        If mvSelectedRow < 0 Then
          DataHelper.AddItem(CareServices.XMLMaintenanceControlTypes.xmctCustomForm, vList)
        Else
          DataHelper.UpdateItem(CareServices.XMLMaintenanceControlTypes.xmctCustomForm, vList)
        End If
        mvCustomPanel.DataChanged = False
        RefreshData()
      End If
    Catch vCareException As CareException
      If vCareException.ErrorNumber = CareException.ErrorNumbers.enOracleUserError OrElse
         vCareException.ErrorNumber = CareException.ErrorNumbers.enODBCUserDefinedError Then
        ShowInformationMessage(vCareException.Message)
      Else
        DataHelper.HandleException(vCareException)
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdNew.Click
    dgr.SelectRow(-1)
    ProcessNew(True)
  End Sub

  Private Sub mvCustomiseMenu_UpdatePanel(ByVal pRevert As Boolean) Handles mvCustomiseMenu.UpdatePanel, mvExamsCustomiseMenu.UpdatePanel
    'Here we should be updating the edit panel to show the new layout but this is really difficult
    'For now we are just going to close the form
    Select Case mvDataType
      Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactExamSummary,
        CareNetServices.XMLContactDataSelectionTypes.xcdtContactExamDetails
        mvExamUnitSelector.sel.ReSelectNode()
    End Select

  End Sub

  Private Sub mvCustomiseMenu_MenuSelected(ByVal pSender As Object, ByVal pItem As CustomiseMenu.CustomiseMenuItems) Handles mvCustomiseMenu.MenuSelected, mvExamsCustomiseMenu.MenuSelected
    Try
      Select Case pItem
        Case CustomiseMenu.CustomiseMenuItems.Edit
          Select Case mvDataType
            Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactExamDetails, CareNetServices.XMLContactDataSelectionTypes.xcdtContactExamSummary
              If mvDataType = CareNetServices.XMLContactDataSelectionTypes.xcdtContactExamDetails Then
                Dim vDataRow As DataRow = DataHelper.GetRowFromDataSet(mvExamUnitSelector.ItemDataSet)
                Dim vContextBooking As Integer = IntegerValue(vDataRow("ExamBookingId").ToString)
                If vContextBooking <> mvExamUnitSelector.ExamBookingId Then
                  ShowInformationMessage(InformationMessages.ImExamBookingEditFromOtherBooking, vContextBooking.ToString)
                  Return
                End If
              End If
              Dim vForm As New frmCardMaintenance(Me, mvContactInfo, mvDataType, mvExamUnitSelector.ItemDataSet, True, 0)
              mvCustomiseMenu = New CustomiseMenu
              Dim vMaintenanceType As CareServices.XMLMaintenanceControlTypes
              If mvDataType = CareNetServices.XMLContactDataSelectionTypes.xcdtContactExamSummary Then
                vMaintenanceType = CareNetServices.XMLMaintenanceControlTypes.xmctExamStudentUnitSummary
              Else
                vMaintenanceType = CareNetServices.XMLMaintenanceControlTypes.xmctExamStudentUnitDetail
              End If
              mvCustomiseMenu.SetContext(vForm, vMaintenanceType, mvContactInfo.ContactGroup)
              vForm.SetCustomiseMenu(mvCustomiseMenu)
              Dim vLocation As Point = mvExamUnitSelector.tab.PointToScreen(mvExamUnitSelector.tab.Location)
              Dim vSize As Size = mvExamUnitSelector.tab.Size
              vForm.SetInitialBounds(vLocation, vSize)
              If MDIForm IsNot Nothing Then
                vForm.Show()
              Else
                vForm.Show(Me)
              End If
          End Select
        Case CustomiseMenu.CustomiseMenuItems.Cancel
          Select Case mvDataType
            Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactExamDetails
              Dim vCanCancel As Boolean = True
              Dim vCancelBookingUnit As Boolean = True 'flag to record whether to cancel just the unit or the whole booking

              Dim vDataRow As DataRow = DataHelper.GetRowFromDataSet(mvExamUnitSelector.ItemDataSet)
              Dim vContextBooking As Integer = IntegerValue(vDataRow("ExamBookingId").ToString)
              If vContextBooking <> mvExamUnitSelector.ExamBookingId Then
                ShowInformationMessage(InformationMessages.ImExamBookingEditFromOtherBooking, vContextBooking.ToString)
                vCanCancel = False
              End If

              If vCanCancel Then
                Dim vParentLinkId As Integer = IntegerValue(vDataRow("ParentUnitLinkId").ToString)
                If vParentLinkId = 0 Then 'User wants to cancel top level unit.  Cancel booking instead?
                  Dim vDlgResult As DialogResult = Utilities.ShowQuestion(QuestionMessages.QmExamBookingUnitCancelBookingInstead, MessageBoxButtons.YesNoCancel)
                  Select Case vDlgResult
                    Case DialogResult.Yes 'User wants to cancel whole booking instead of top level unit
                      vCancelBookingUnit = False
                    Case DialogResult.Cancel
                      vCanCancel = False
                    Case Else
                      vCancelBookingUnit = True 'User selected No to cancelling Booking.  Proceed with default functionality of cancelling EBU only
                  End Select
                End If
              End If

              If vCanCancel Then
                Dim vParams As ParameterList = FormHelper.ShowApplicationParameters(CareServices.FunctionParameterTypes.fptCancellationReasonAndSource)
                If vParams IsNot Nothing AndAlso vParams.Count > 0 Then
                  vParams.IntegerValue("ExamBookingId") = mvExamUnitSelector.ExamBookingId
                  If vCancelBookingUnit Then
                    vParams.IntegerValue("ExamBookingUnitId") = mvExamUnitSelector.ExamBookingUnitId
                  End If
                  CancelExamBooking(vParams)
                End If
              End If
          End Select
      End Select
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub CancelExamBooking(ByVal pParams As ParameterList)
    Dim vRetry As Boolean
    Dim vReturnList As New ParameterList
    Do
      vRetry = False
      Try
        vReturnList = ExamsDataHelper.CancelExamBooking(pParams)
      Catch vEx As CareException
        Select Case vEx.ErrorNumber
          Case CareException.ErrorNumbers.enCancellationFeeMissing
            'Needs to get the amount from the transaction
            Dim vRow As DataRow = DataHelper.GetRowFromDataSet(DataHelper.GetTransactionData(CareServices.XMLTransactionDataSelectionTypes.xtdtTransactionAnalysis, SelectRowItemNumber(dgr.CurrentRow, "BatchNumber"), SelectRowItemNumber(dgr.CurrentRow, "TransactionNumber"), 0, SelectRowItemNumber(dgr.CurrentRow, "LineNumber")))
            Dim vDefaults As New ParameterList
            If vRow IsNot Nothing Then vDefaults("CancellationFeeAmount") = vRow("Amount").ToString
            Dim vCancellationFeeParams As ParameterList = FormHelper.ShowApplicationParameters(CareServices.FunctionParameterTypes.fptEnterCancellationFee, vDefaults)
            If vCancellationFeeParams IsNot Nothing AndAlso vCancellationFeeParams.Count > 0 Then
              pParams("CancellationFeeAmount") = vCancellationFeeParams("CancellationFeeAmount")
              vRetry = True
            End If
          Case CareException.ErrorNumbers.enCCAuthorisationFailed, CareException.ErrorNumbers.enCardAuthorisationUnexpectedTimeout
            ShowInformationMessage(vEx.Message)
          Case CareException.ErrorNumbers.enInvoiceAllocationError, CareException.ErrorNumbers.enUnallocateCreditNote
            If ShowQuestion(vEx.Message, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
              pParams("AllocationsChecked") = "Y"
              vRetry = True
              If vEx.ErrorNumber = CareException.ErrorNumbers.enUnallocateCreditNote Then pParams("UnallocateCreditNote") = "Y"
            Else
              Exit Sub
            End If
          Case CareException.ErrorNumbers.enAllocateOrUnallocateCreditNote
            Dim vDialogueResult As System.Windows.Forms.DialogResult = ShowQuestion(vEx.Message, MessageBoxButtons.YesNoCancel)
            Select Case vDialogueResult
              Case System.Windows.Forms.DialogResult.Yes, System.Windows.Forms.DialogResult.No
                pParams("AllocationsChecked") = "Y"
                vRetry = True
                If vDialogueResult = System.Windows.Forms.DialogResult.No Then pParams("UnallocateCreditNote") = "Y"
              Case Else 'Cancel
                Exit Sub
            End Select
          Case Else
            Throw vEx
        End Select
      End Try
    Loop While vRetry
    If vReturnList.Contains("Message") Then ShowInformationMessage(vReturnList("Message"))
    Dim vMsg As String = If(vReturnList.Contains("ExamBookingUnitId"), InformationMessages.ImExamBookingUnitCancelled, InformationMessages.ImExamBookingCancelled)
    ShowInformationMessage(vMsg)
      RefreshCard()
  End Sub
  Private Sub ProcessNew()
    ProcessNew(False)
  End Sub
  Private Sub ProcessNew(ByVal pFromNew As Boolean)
    Try
      mvSelectedRow = -1
      mvCustomPanel.Clear()
      If mvAllowUpdate = False Then
        mvCustomPanel.EnableControls(mvCustomPanel, True)
        cmdSave.Enabled = pFromNew
      End If
      cmdDelete.Enabled = False
      If IsEditableCustomForm(mvCustomFormNumber) Then
        Dim vList As New ParameterList(True)
        vList.IntegerValue("CustomForm") = mvCustomFormNumber
        vList("Default") = "Y"
        vList("Values") = ""
        Dim vTable As DataTable = DataHelper.GetTableFromDataSet(DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactCustomFormData, mvContactInfo.ContactNumber, vList))
        If vTable IsNot Nothing AndAlso vTable.Rows.Count > 0 AndAlso vTable.Columns.Contains("ParameterName") Then
          For Each vRow As DataRow In vTable.Rows
            If vRow.Item("ParameterName").ToString = "Status" Then
              Dim vTextLookup As TextLookupBox = TryCast(FindControl(mvCustomPanel, "Status"), TextLookupBox)
              If vTextLookup IsNot Nothing Then vTextLookup.FillComboWithRestriction(vRow.Item("DefaultValue").ToString, mvContactInfo.ContactGroup)
              mvCustomPanel.SetValue(vRow.Item("ParameterName").ToString, vRow.Item("DefaultValue").ToString)
              If vTextLookup IsNot Nothing Then vTextLookup.SetDependancies()
            Else
              mvCustomPanel.SetValue(vRow.Item("ParameterName").ToString, vRow.Item("DefaultValue").ToString)
            End If
          Next
        End If
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
  Private Sub ProcessSubRowSelection(ByVal pSender As Object, ByVal pRow As Integer, ByVal pDataRow As Integer)
    If pSender Is dts.DisplayGrid(0) Then
      Select Case mvDataType
        Case CareServices.XMLContactDataSelectionTypes.xcdtContactProcessedTransactions
          mvAnalysisFinancialMenu.SetContext(mvContactInfo, mvFinancialMenu.DataRow, GetDataRow(mvDataSet2, pDataRow))
          If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.fp_pp_show_payment_details) Then
            Dim vRow As DataRow = GetDataRow(mvDataSet2, pDataRow)
            Dim vPayPlanNumber As Integer = IntegerValue(vRow.Item("PaymentPlanNumber").ToString)
            Dim vPaymentNumber As Integer = IntegerValue(vRow.Item("PaymentPlanPayNumber").ToString)
            dts.DisplayGrid(1).Populate(DataHelper.GetPaymentPlanData(CareNetServices.XMLPaymentPlanDataSelectionTypes.xpdtPaymentPlanPaymentHistDetails, vPayPlanNumber, 0, "", vPaymentNumber))
            ProcessSubRowSelection(dts.DisplayGrid(1), 0, 0)
          End If
        Case CareServices.XMLContactDataSelectionTypes.xcdtContactCreditCustomers
          mvFinancialMenu.SetContext(mvDataType, GetDataRow(mvDataSet2, pDataRow), mvContactInfo, False)
          If dts.DisplayGrid(0).DataRowCount > pRow Then
            Dim vTransType As String = SelectSubRowItem(mvDataSet2, pDataRow, "TransactionType")
            Dim vBatchNumber As Integer = SelectSubRowItemNumber(mvDataSet2, pDataRow, "BatchNumber")
            Dim vTransNumber As Integer = SelectSubRowItemNumber(mvDataSet2, pDataRow, "TransactionNumber")
            Dim vInvoiceNumber As String = SelectSubRowItem(mvDataSet2, pDataRow, "InvoiceNumber")
            Dim vList As New ParameterList(True, True)
            vList.IntegerValue("BatchNumber") = vBatchNumber
            vList.IntegerValue("TransactionNumber") = vTransNumber
            If IntegerValue(vInvoiceNumber) > 0 Then vList("InvoiceNumber") = vInvoiceNumber
            mvDataSet3 = DataHelper.GetContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactSalesLedgerReceipts, mvContactInfo.ContactNumber, vList)
            dts.SubDisplayGrid(0).Populate(mvDataSet3)
            dts.SubDisplayGrid(0).SetToolBarVisible()
            mvFinancialSubMenu = New FinancialSubMenu(Me, CareNetServices.XMLContactDataSelectionTypes.xcdtContactSalesLedgerReceipts, mvContactInfo)
            mvFinancialSubMenu.SetContext(mvDataType, GetDataRow(mvDataSet2, pDataRow), mvContactInfo, False)
            If dts.SubDisplayGrid(0).ContextMenuStrip IsNot mvFinancialSubMenu Then dts.SubDisplayGrid(0).ContextMenuStrip = mvFinancialSubMenu
            dts.AdjustSubDisplayGrid(0)
            ProcessSubRowSelection(dts.SubDisplayGrid(0), 0, 0)
          End If
        Case CareServices.XMLContactDataSelectionTypes.xcdtContactPledges,
             CareServices.XMLContactDataSelectionTypes.xcdtContactPostPGPledges,
             CareServices.XMLContactDataSelectionTypes.xcdtContactLegacyBequests
          mvTransactionLinkMenu.SetContext(mvContactInfo, GetDataRow(mvDataSet2, pDataRow))
        Case CareServices.XMLContactDataSelectionTypes.xcdtContactFundraising
          mvFundraisingPaymentMenu.SetContext(mvContactInfo, GetDataRow(mvDataSet, dgr.CurrentDataRow), GetDataRow(mvDataSet2, pDataRow))
          Dim e As New System.ComponentModel.CancelEventArgs
          mvFundraisingPaymentMenu.SetVisibleItems(e)

          dts.DisplayGrid(0).SetToolBarVisible()
          If dts.DisplayGrid(0).DataRowCount > pRow Then
            Dim FRNumber As Integer = SelectSubRowItemNumber(mvDataSet2, pDataRow, "FundraisingRequestNumber")
            Dim FPSNumber As Integer = SelectSubRowItemNumber(mvDataSet2, pDataRow, "ScheduledPaymentNumber")
            Dim vList As New ParameterList(True)
            vList.AddSystemColumns()
            vList.IntegerValue("ScheduledPaymentNumber") = FPSNumber
            mvDataSet3 = DataHelper.GetFundraisingData(CareNetServices.XMLFundraisingDataSelectionTypes.xfdtPaymentHistory, FRNumber, vList)
            dts.SubDisplayGrid(0).Populate(mvDataSet3)
            dts.AdjustSubDisplayGrid(0)
            If dts.SubDisplayGrid(0).ContextMenuStrip IsNot mvTransactionLinkMenu Then dts.SubGridContextMenuStrip(0) = mvTransactionLinkMenu
            ProcessSubRowSelection(dts.SubDisplayGrid(0), 0, 0)
          End If
        Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactServiceBookings
          mvTransactionLinkMenu.SetContext(mvContactInfo, GetDataRow(mvDataSet2, dts.DisplayGrid(0).CurrentRow))
        Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactCPDPointsWithoutCycle
          If mvDocumentMenu IsNot Nothing Then
            mvDocumentMenu.DocumentNumber = SelectSubRowItemNumber(mvDataSet2, pDataRow, "DocumentNumber")
            mvDocumentMenu.SetVisibleItems(New System.ComponentModel.CancelEventArgs)
          End If
        Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactMeetings
          'Meeting Links - we need to change the available menu items and set the grid toolbar based upon the menu items
          dgr0MenuStripBuilder()
          dts.DisplayGrid(0).SetToolBarVisible()
      End Select
    ElseIf pSender Is dts.DisplayGrid(1) Then
      Select Case mvDataType
        Case CareServices.XMLContactDataSelectionTypes.xcdtContactPaymentPlans, CareServices.XMLContactDataSelectionTypes.xcdtContactMembershipDetails, CareNetServices.XMLContactDataSelectionTypes.xcdtContactLoans
          mvTransactionLinkMenu.SetContext(mvContactInfo, GetDataRow(mvDataSet3, pDataRow))
        Case CareServices.XMLContactDataSelectionTypes.xcdtContactPurchaseOrders
          mvPurchaseOrderPaymentMenu.SetContext(mvContactInfo, GetDataRow(mvDataSet3, pDataRow))
        Case CareServices.XMLContactDataSelectionTypes.xcdtContactPurchaseInvoices
          mvPurchaseInvoiceChequeMenu.SetContext(mvContactInfo, GetDataRow(mvDataSet3, pDataRow))
        Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactCPD
          If mvDocumentMenu IsNot Nothing Then
            mvDocumentMenu.DocumentNumber = SelectSubRowItemNumber(mvDataSet3, pDataRow, "DocumentNumber")
            mvDocumentMenu.SetVisibleItems(New System.ComponentModel.CancelEventArgs)
          End If
      End Select
    ElseIf pSender Is dts.DisplayGrid(2) Then
      Select Case mvDataType
        Case CareServices.XMLContactDataSelectionTypes.xcdtContactPaymentPlans
          Dim vRow As DataRow = Nothing
          If dts.DisplayGrid(2).DataRowCount > 0 Then
            vRow = GetDataRow(mvDataSet4, pDataRow)
            mvFinancialSubMenu.SetContext(mvDataType, vRow, mvContactInfo, mvReadOnlyPage)
          End If
      End Select
    ElseIf pSender Is dts.SubDisplayGrid(0) Then
      Select Case mvDataType
        Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactCreditCustomers
          Dim vRow As DataRow = Nothing
          If dts.SubDisplayGrid(0) IsNot Nothing AndAlso dts.SubDisplayGrid(0).DataRowCount > 0 Then vRow = GetDataRow(mvDataSet3, pDataRow)
          mvFinancialSubMenu.SetContext(mvDataType, vRow, mvContactInfo, False)
        Case CareServices.XMLContactDataSelectionTypes.xcdtContactFundraising
          Dim vRow As DataRow = Nothing
          If dts.SubDisplayGrid(0).DataRowCount > 0 Then vRow = GetDataRow(mvDataSet3, pDataRow)
          mvTransactionLinkMenu.SetContext(mvContactInfo, vRow)
      End Select
    ElseIf pSender Is dts.DisplayGrid(3) Then
      Select Case mvDataType
        Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactViewOrganisations, CareNetServices.XMLContactDataSelectionTypes.xcdtContactPositions
          If mvActionMenu IsNot Nothing Then
            mvActionMenu.ActionNumber = SelectSubRowItemNumber(mvDataSet5, pDataRow, "ActionNumber")
            mvActionMenu.SetVisibleItems(New System.ComponentModel.CancelEventArgs)
            mvContactInfo.SelectedActionNumber = mvActionMenu.ActionNumber
          End If
      End Select
    ElseIf pSender Is dts.DisplayGrid(4) Then
      Select Case mvDataType
        Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactFundraising
          If dts.DisplayGrid(4).DataRowCount > pRow Then
            Dim vNumber As Integer = SelectSubRowItemNumber(mvDataSet4, pDataRow, "ActionNumber")
            mvActionMenu.ActionNumber = vNumber
            mvActionMenu.ActionStatus = SelectSubRowItem(mvDataSet4, pDataRow, "ActionStatus")
            mvContactInfo.SelectedActionNumber = vNumber
            Dim e As New System.ComponentModel.CancelEventArgs
            mvActionMenu.SetVisibleItems(e)
            dts.DisplayGrid(4).SetToolBarVisible()
            Dim vDataSet As DataSet = DataHelper.GetActionData(CareNetServices.XMLActionDataSelectionTypes.xadtActionLinks, vNumber)
            Dim vDGR As New DisplayGrid
            vDGR.Populate(vDataSet)
            mvActionMenu.SetNotify(vDGR)
          End If
        Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactViewOrganisations, CareNetServices.XMLContactDataSelectionTypes.xcdtContactPositions
          If mvDocumentMenu IsNot Nothing Then
            mvDocumentMenu.DocumentNumber = SelectSubRowItemNumber(mvDataSet6, pDataRow, "DocumentNumber")
            mvDocumentMenu.SetVisibleItems(New System.ComponentModel.CancelEventArgs)
          End If
      End Select
    End If
  End Sub

  Private Function GetDataRow(ByVal pDataSet As DataSet, ByVal pRow As Integer) As DataRow
    Dim vTable As DataTable = DataHelper.GetTableFromDataSet(pDataSet)
    If vTable IsNot Nothing Then
      Return vTable.Rows(pRow)
    Else
      Return Nothing
    End If
  End Function
  Private Function CancelPaymentPlan(ByVal pItem As CareServices.XMLPaymentPlanMenuTypes, ByRef pCompleted As Boolean) As ParameterList
    Dim vTryAgain As Boolean = False
    Dim vReturnList As ParameterList
    Dim vList As New ParameterList(True)
    vList.IntegerValue("PaymentPlanNumber") = SelectRowItemNumber(dgr.CurrentDataRow, "PaymentPlanNumber")
    vList.IntegerValue("ContactNumber") = ContactInfo.ContactNumber
    Select Case pItem
      Case CareServices.XMLPaymentPlanMenuTypes.xpmtCancelCovenant
        vList.IntegerValue("CovenantNumber") = SelectRowItemNumber(dgr.CurrentDataRow, "CovenantNumber")
      Case CareServices.XMLPaymentPlanMenuTypes.xpmtCancelCreditCardAuthority, CareServices.XMLPaymentPlanMenuTypes.xpmtFutureCancelCreditCardAuthority
        vList.IntegerValue("CreditCardAuthorityNumber") = SelectRowItemNumber(dgr.CurrentDataRow, "CreditCardAuthorityNumber")
      Case CareServices.XMLPaymentPlanMenuTypes.xpmtCancelDirectDebit, CareServices.XMLPaymentPlanMenuTypes.xpmtFutureCancelDirectDebit
        vList.IntegerValue("DirectDebitNumber") = SelectRowItemNumber(dgr.CurrentDataRow, "DirectDebitNumber")
      Case CareNetServices.XMLPaymentPlanMenuTypes.xpmtCancelLoan
        vList.IntegerValue("LoanNumber") = SelectRowItemNumber(dgr.CurrentDataRow, "LoanNumber")
      Case CareServices.XMLPaymentPlanMenuTypes.xpmtCancelMember
        vList.IntegerValue("MembershipNumber") = SelectRowItemNumber(dgr.CurrentDataRow, "MembershipNumber")

      Case CareServices.XMLPaymentPlanMenuTypes.xpmtCancelStandingOrder, CareServices.XMLPaymentPlanMenuTypes.xpmtFutureCancelStandingOrder
        vList.IntegerValue("StandingOrderNumber") = SelectRowItemNumber(dgr.CurrentDataRow, "StandingOrderNumber")

    End Select
    Do
      vTryAgain = False
      vReturnList = DataHelper.ProcessPaymentPlanMenu(pItem, vList)
      If Not vReturnList.Contains("PaymentPlanNumber") Then
        If pItem = CareServices.XMLPaymentPlanMenuTypes.xpmtFutureCancelCreditCardAuthority _
        OrElse pItem = CareServices.XMLPaymentPlanMenuTypes.xpmtFutureCancelDirectDebit _
        OrElse pItem = CareServices.XMLPaymentPlanMenuTypes.xpmtFutureCancelPaymentPlan _
        OrElse pItem = CareServices.XMLPaymentPlanMenuTypes.xpmtFutureCancelStandingOrder Then
          vReturnList("FutureCancellation") = "Y"
          Dim vParams As ParameterList = FormHelper.ShowApplicationParameters(CareServices.FunctionParameterTypes.fptCancellationReasonSourceAndDate, vReturnList)
          If vParams IsNot Nothing AndAlso vParams.Count > 0 Then
            For Each vValue As DictionaryEntry In vParams
              If Not vList.Contains(vValue.Key.ToString) Then vList.Add(vValue.Key.ToString, vValue.Value.ToString)
            Next
            vTryAgain = True
          End If
        Else
          Select Case pItem
            Case CareServices.XMLPaymentPlanMenuTypes.xpmtCancelCovenant
              vReturnList("CancellationType") = "Covenant"
            Case CareServices.XMLPaymentPlanMenuTypes.xpmtCancelCreditCardAuthority
              vReturnList("CancellationType") = "Credit Card Authority"
            Case CareServices.XMLPaymentPlanMenuTypes.xpmtCancelDirectDebit
              vReturnList("CancellationType") = "Direct Debit"
            Case CareServices.XMLPaymentPlanMenuTypes.xpmtCancelMember
              vReturnList("CancellationType") = "Membership"
            Case CareServices.XMLPaymentPlanMenuTypes.xpmtCancelPaymentPlan
              vReturnList("CancellationType") = "Payment Plan"
            Case CareServices.XMLPaymentPlanMenuTypes.xpmtCancelStandingOrder
              vReturnList("CancellationType") = "Standing Order"
          End Select
          If SelectRowItem(dgr.CurrentDataRow, "CancellationReason").Length > 0 Then
            vReturnList.IntegerValue("PaymentPlanNumber") = SelectRowItemNumber(dgr.CurrentDataRow, "PaymentPlanNumber")
            vReturnList.IntegerValue("ContactNumber") = ContactInfo.ContactNumber
            vReturnList("CancellationReason") = SelectRowItem(dgr.CurrentDataRow, "CancellationReason")
            vReturnList("CancellationSource") = SelectRowItem(dgr.CurrentDataRow, "CancellationSource")
            vReturnList("StatusDate") = Today.ToString
          End If
          Dim vParams As ParameterList = FormHelper.ShowApplicationParameters(CareServices.FunctionParameterTypes.fptCancelPaymentPlan, vReturnList)
          If vParams IsNot Nothing AndAlso vParams.Count > 0 Then
            'add the necessary items in vParams to vList
            For Each vValue As DictionaryEntry In vParams
              If Not vList.Contains(vValue.Key.ToString) Then vList.Add(vValue.Key.ToString, vValue.Value.ToString)
            Next
            If vReturnList.Contains("MembersPerOrder") Then vList("MembersPerOrder") = vReturnList("MembersPerOrder")
            If vReturnList.Contains("AssociateMembership") Then vList("AssociateMembership") = vReturnList("AssociateMembership")
            If vReturnList.Contains("UpdateDetailsSource") And Not vParams.Contains("UpdateDetailsSource") Then
              vList("UpdateDetailsSource") = vReturnList("UpdateDetailsSource")
              If vList("UpdateDetailsSource") = "Ask" Then vList("UpdateDetailsSource") = "N"
            End If
            vTryAgain = True
            vList("ChangeDDPayer") = "N"   'Set this to 'N' so that the server knows we have already checked this.
            If pItem = CareNetServices.XMLPaymentPlanMenuTypes.xpmtCancelMember _
            AndAlso BooleanValue(vList.ValueIfSet("CancelMember")) = True AndAlso BooleanValue(vList.ValueIfSet("CancelMembership")) = False AndAlso BooleanValue(vList.ValueIfSet("CancelDirectDebit")) = False Then
              'Cancel just the individual Member
              If SelectRowItemNumber(dgr.CurrentDataRow, "MembersPerOrder") = 2 AndAlso BooleanValue(SelectRowItem(dgr.CurrentDataRow, "DirectDebitStatus")) = True Then
                'When MembersPerOrder = 2 and the Membership is paid by DirectDebit, may need to move the DD
                Dim vDDCheckList As New ParameterList(True, True)
                vDDCheckList.IntegerValue("PaymentPlanNumber") = SelectRowItemNumber(dgr.CurrentDataRow, "PaymentPlanNumber")
                vDDCheckList.IntegerValue("MembershipNumber") = SelectRowItemNumber(dgr.CurrentDataRow, "MembershipNumber")
                vDDCheckList("CancellationReason") = vList("CancellationReason")
                Dim vDDReturnList As ParameterList = DataHelper.CanChangeDDPayer(vDDCheckList)
                If BooleanValue(vDDReturnList.ValueIfSet("CanChangeDDPayer")) Then
                  Dim vDDResult As DialogResult = ShowQuestion(QuestionMessages.QmCancelMemberMoveDD, MessageBoxButtons.YesNoCancel, vDDReturnList("DirectDebitPayerName"), vDDReturnList("DirectDebitNewPayerName"))
                  If vDDResult = System.Windows.Forms.DialogResult.Cancel Then
                    vTryAgain = False
                  Else
                    vList("ChangeDDPayer") = If(vDDResult = System.Windows.Forms.DialogResult.Yes, "Y", "N")
                    If vDDResult = System.Windows.Forms.DialogResult.Yes Then vList("DirectDebitNewPayerContactNumber") = vDDReturnList("DirectDebitNewPayerContactNumber")
                  End If
                End If
              End If
            End If
          End If
          If vReturnList.Contains("CancellationReason") Or vReturnList.Contains("CancellationSource") Then
            If vParams.Contains("CancellationReason") Or vParams.Contains("Source") Then
              If vReturnList("CancellationReason") <> vParams("CancellationReason") OrElse vReturnList("CancellationSource") <> vParams("Source") Then
                vTryAgain = True
                vList("CancellationReason") = vParams("CancellationReason")
                vList("OriginalCancellationReason") = vReturnList("CancellationReason")
                'to do - Simon to add new enum in CareService
                pItem = CType(CareNetServices.XMLPaymentPlanMenuTypes.xpmtChangeCancellationReason, CareServices.XMLPaymentPlanMenuTypes)
                If vParams.Contains("Source") Then vList("CancellationSource") = vParams("Source")
                Select Case pItem
                  Case CareServices.XMLPaymentPlanMenuTypes.xpmtCancelCovenant
                    vList("CancelCovenant") = "Y"
                  Case CareServices.XMLPaymentPlanMenuTypes.xpmtCancelCreditCardAuthority
                    vList("CancelCreditCardAuthority") = "Y"
                  Case CareServices.XMLPaymentPlanMenuTypes.xpmtCancelDirectDebit
                    vList("CancelDirectDebit") = "Y"
                  Case CareServices.XMLPaymentPlanMenuTypes.xpmtCancelMember
                    vList("CancelMember") = "Y"
                  Case CareServices.XMLPaymentPlanMenuTypes.xpmtCancelPaymentPlan
                    vList("CancelPaymentPlan") = "Y"
                  Case CareServices.XMLPaymentPlanMenuTypes.xpmtCancelStandingOrder
                    vList("CancelStandingOrder") = "Y"
                End Select
              Else
                vTryAgain = False
              End If
            End If
          End If
        End If
      Else
        pCompleted = True
      End If
    Loop While vTryAgain
    Return vReturnList
  End Function

  Private Function AllowEditDocument() As Boolean
    Select Case mvDataType
      Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactCommunicationHistory
        Return False
      Case Else
        Return True
    End Select
  End Function

  Private Sub mvFundraisingPaymentMenu_MenuSelected(ByVal pItem As BaseFinancialMenu.FinancialMenuItems, ByVal pDataRow As System.Data.DataRow, ByVal pChangeDetails As Boolean, ByVal pFinancialMenu As BaseFinancialMenu) Handles mvFundraisingPaymentMenu.MenuSelected
    Try
      Select Case pItem
        Case BaseFinancialMenu.FinancialMenuItems.fmiNew, BaseFinancialMenu.FinancialMenuItems.fmiEdit
          Dim vList As New ParameterList
          vList("FundraisingRequestNumber") = dgr.GetValue(dgr.CurrentRow, "FundraisingRequestNumber")
          vList("HideDefaultPaymentType") = CBoolYN(DoubleValue(dgr.GetValue(dgr.CurrentRow, "PledgedAmount")) = 0 AndAlso DoubleValue(dgr.GetValue(dgr.CurrentRow, "ExpectedAmount")) = 0)
          Dim vForm As New frmCardMaintenance(Me, mvContactInfo, CareServices.XMLContactDataSelectionTypes.xcdtNone, mvDataSet2, pItem = BaseFinancialMenu.FinancialMenuItems.fmiEdit, dts.DisplayGrid(0).CurrentRow, CareServices.XMLMaintenanceControlTypes.xmctFundraisingPaymentSchedule, vList)
          mvCustomiseMenu = New CustomiseMenu
          mvCustomiseMenu.SetContext(vForm, CareServices.XMLMaintenanceControlTypes.xmctFundraisingPaymentSchedule, mvContactInfo.ContactGroup)
          vForm.SetCustomiseMenu(mvCustomiseMenu)
          vForm.Show()
        Case BaseFinancialMenu.FinancialMenuItems.fmiGoToActions
          If mvDataSet4 IsNot Nothing AndAlso mvDataSet4.Tables.Contains("DataRow") Then
            Dim vDataSet As DataSet = mvDataSet4.Copy
            vDataSet.Tables("DataRow").DefaultView.RowFilter = "ScheduledPaymentNumber = '" & dts.DisplayGrid(0).GetValue(dts.DisplayGrid(0).CurrentDataRow, "ScheduledPaymentNumber") _
            & "' AND ISNULL(CompletedOn,'') = ''"
            Dim vTable As DataTable = vDataSet.Tables("DataRow").DefaultView.ToTable
            If vTable.Rows.Count > 0 Then
              vDataSet.Tables.Remove("DataRow")
              vDataSet.Tables.Add(vTable)
              Dim vForm As New frmSelectListItem(vDataSet, frmSelectListItem.ListItemTypes.litFundraisingActions)
              If vForm.ShowDialog = System.Windows.Forms.DialogResult.OK Then
                dts.DisplayGrid(4).SelectRow("ActionNumber", vForm.SelectedRow.ToString)
                dts.SelectTab(5)
              End If
            End If
          End If
        Case BaseFinancialMenu.FinancialMenuItems.fmiNewAdHocAction
          HandleMenuClick(False, MenuSource.actiondgr0)
        Case BaseFinancialMenu.FinancialMenuItems.fmiNewActionFromTemplate
          Dim vList As New ParameterList
          vList("FundraisingRequestNumber") = dgr.GetValue(dgr.CurrentDataRow, "FundraisingRequestNumber")
          vList("Logname") = dgr.GetValue(dgr.CurrentDataRow, "Logname")
          vList("ScheduledPaymentNumber") = dts.DisplayGrid(0).GetValue(dts.DisplayGrid(0).CurrentDataRow, "ScheduledPaymentNumber")
          FormHelper.NewActionFromTemplate(Me, mvContactInfo.ContactNumber, vList)
      End Select
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
  Private Sub mvFinancialMenu_MenuSelected(ByVal pItem As FinancialMenu.FinancialMenuItems, ByVal pDataRow As DataRow, ByVal pChangeDetails As Boolean, ByVal pFinancialMenu As BaseFinancialMenu) Handles mvFinancialMenu.MenuSelected
    Dim vInvoiceParams As New ParameterList
    Try
      Dim vRefresh As Boolean = False
      Select Case pItem
        Case FinancialMenu.FinancialMenuItems.fmiNew, FinancialMenu.FinancialMenuItems.fmiEdit
          Dim vEdit As Boolean = pItem = FinancialMenu.FinancialMenuItems.fmiEdit
          Dim vForm As New frmCardMaintenance(Me, mvContactInfo, mvDataType, mvDataSet, vEdit, dgr.CurrentRow)
          ShowMaintenanceForm(vForm)

        Case FinancialMenu.FinancialMenuItems.fmiAdvanceRenewalDate
          Dim vDefaults As New ParameterList
          vDefaults("RenewalDate") = SelectRowItem(dgr.CurrentRow, "RenewalDate")
          vDefaults("NewRenewalDate") = SelectRowItem(dgr.CurrentRow, "RenewalDate")
          Dim vAppParamsList As ParameterList = FormHelper.ShowApplicationParameters(CareServices.FunctionParameterTypes.fptAdvanceRenewalDate, vDefaults)
          If vAppParamsList IsNot Nothing AndAlso vAppParamsList.Count > 0 Then
            Dim vList As New ParameterList(True)
            vList.IntegerValue("PaymentPlanNumber") = SelectRowItemNumber(dgr.CurrentRow, "PaymentPlanNumber")
            If vAppParamsList.Contains("RenewalChangeReason") Then vList("RenewalChangeReason") = vAppParamsList("RenewalChangeReason")
            vList("RenewalChangeValue") = vAppParamsList("AdvanceMonths")
            vList("RenewalDate") = vAppParamsList("NewRenewalDate")
            Dim vReturnList As ParameterList = DataHelper.ProcessPaymentPlanMenu(CareServices.XMLPaymentPlanMenuTypes.xpmtAdvanceRenewalDate, vList)
            If vReturnList.IntegerValue("PaymentPlanNumber") > 0 Then ShowInformationMessage(InformationMessages.ImRenewalDateAdvanced, vDefaults("RenewalDate"), vList("RenewalDate"))
            vRefresh = True
          End If

        Case FinancialMenu.FinancialMenuItems.fmiAmendBooking
          Dim vCareEventInfo As New CareEventInfo(IntegerValue(dgr.GetValue(dgr.CurrentRow, "EventNumber")))
          vRefresh = FormHelper.ShowAmendEventBookingForm(Me, dgr, vCareEventInfo)

        Case FinancialMenu.FinancialMenuItems.fmiAmendDueDate
          Dim vList As New ParameterList
          vList("DueDate") = GetDataRow(mvDataSet2, dts.DisplayGrid(0).CurrentRow).Item("DueDate").ToString
          Dim vReturnList As ParameterList = FormHelper.ShowApplicationParameters(CareServices.FunctionParameterTypes.fptInvoicePaymentDue, vList)
          If vReturnList.Count > 0 AndAlso vReturnList("DueDate") <> vList("DueDate") Then
            DataHelper.UpdateInvoice(CInt(GetDataRow(mvDataSet2, dts.DisplayGrid(0).CurrentRow).Item("StoredInvoiceNumber")), vReturnList)
            vRefresh = True
          End If

        Case BaseFinancialMenu.FinancialMenuItems.fmiChangeInvoiceAddress
          Dim vAddressNumber As Integer = IntegerValue(GetDataRow(mvDataSet2, dts.DisplayGrid(0).CurrentRow).Item("AddressNumber").ToString)
          Dim vDataSet As DataSet = DataHelper.GetContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactAddresses, mvContactInfo.ContactNumber)
          Dim vForm As New frmSelectAddress(vDataSet, frmSelectAddress.SelectAddressTypes.satTrader, vAddressNumber)
          If vForm.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            Dim vList As New ParameterList(True)
            vList.IntegerValue("AddressNumber") = vForm.AddressNumber
            vList.IntegerValue("BatchNumber") = mvFinancialMenu.TargetBatchNumber
            vList.IntegerValue("TransactionNumber") = mvFinancialMenu.TargetTransactionNumber
            DataHelper.UpdateInvoice(0, vList)
            vRefresh = True
          End If

        Case FinancialMenu.FinancialMenuItems.fmiAmendMembership
          Dim vForm As New frmCardMaintenance(Me, mvContactInfo, mvDataType, mvDataSet, True, dgr.CurrentRow)
          vForm.Show()
        Case FinancialMenu.FinancialMenuItems.fmiChangeCancel
          Select Case mvDataType
            Case CareServices.XMLContactDataSelectionTypes.xcdtContactMembershipDetails
              CancelPaymentPlan(CareServices.XMLPaymentPlanMenuTypes.xpmtCancelMember, vRefresh)
            Case CareServices.XMLContactDataSelectionTypes.xcdtContactDirectDebits
              CancelPaymentPlan(CareServices.XMLPaymentPlanMenuTypes.xpmtCancelDirectDebit, vRefresh)
            Case CareServices.XMLContactDataSelectionTypes.xcdtContactCovenants
              CancelPaymentPlan(CareServices.XMLPaymentPlanMenuTypes.xpmtCancelCovenant, vRefresh)
            Case CareServices.XMLContactDataSelectionTypes.xcdtContactStandingOrders
              CancelPaymentPlan(CareServices.XMLPaymentPlanMenuTypes.xpmtCancelStandingOrder, vRefresh)
            Case CareServices.XMLContactDataSelectionTypes.xcdtContactCreditCardAuthorities
              CancelPaymentPlan(CareServices.XMLPaymentPlanMenuTypes.xpmtCancelCreditCardAuthority, vRefresh)
            Case CareServices.XMLContactDataSelectionTypes.xcdtContactPaymentPlans
              CancelPaymentPlan(CareServices.XMLPaymentPlanMenuTypes.xpmtCancelPaymentPlan, vRefresh)
          End Select
        Case FinancialMenu.FinancialMenuItems.fmiCancelEventBooking, FinancialMenu.FinancialMenuItems.fmiCancel, BaseFinancialMenu.FinancialMenuItems.fmiCancelExamBooking
          Select Case mvDataType
            Case CareServices.XMLContactDataSelectionTypes.xcdtContactCovenants
              CancelPaymentPlan(CareServices.XMLPaymentPlanMenuTypes.xpmtCancelCovenant, vRefresh)

            Case CareServices.XMLContactDataSelectionTypes.xcdtContactCreditCardAuthorities
              CancelPaymentPlan(CareServices.XMLPaymentPlanMenuTypes.xpmtCancelCreditCardAuthority, vRefresh)

            Case CareServices.XMLContactDataSelectionTypes.xcdtContactDirectDebits
              CancelPaymentPlan(CareServices.XMLPaymentPlanMenuTypes.xpmtCancelDirectDebit, vRefresh)

            Case CareServices.XMLContactDataSelectionTypes.xcdtContactEventBookings, CareNetServices.XMLContactDataSelectionTypes.xcdtContactEventRoomBookings, CareNetServices.XMLContactDataSelectionTypes.xcdtContactExamDetails
              Dim vBatch As New BatchInfo(IntegerValue(pDataRow("BatchNumber")))
              Dim vList = New ParameterList(True, False)
              vList.IntegerValue("BatchNumber") = vBatch.BatchNumber
              Dim vParams As ParameterList = FormHelper.ShowApplicationParameters(CareServices.FunctionParameterTypes.fptCancellationReasonAndSource, vList)
              If vParams IsNot Nothing AndAlso vParams.Count > 0 Then
                vRefresh = True
                Dim vEventBooking As Boolean = mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactEventBookings
                Dim vExamBooking As Boolean = mvDataType = CareNetServices.XMLContactDataSelectionTypes.xcdtContactExamDetails
                If vEventBooking Then
                  vParams("BookingNumber") = dgr.GetValue(dgr.CurrentRow, "BookingNumber").ToString
                ElseIf vExamBooking Then
                  vParams("ExamBookingId") = dgr.GetValue(dgr.CurrentRow, "ExamBookingId").ToString
                Else
                  vParams("RoomBookingNumber") = dgr.GetValue(dgr.CurrentRow, "RoomBookingNumber").ToString
                  vParams("BookingStatus") = frmEventSet.ebsCancelled
                End If
                vInvoiceParams = vParams
                If vExamBooking Then
                  CancelExamBooking(vParams)
                Else
                  Dim vReturnList As New ParameterList
                  Try
                    If vEventBooking Then
                      vReturnList = DataHelper.CancelEventBooking(vParams)
                    Else
                      vReturnList = DataHelper.UpdateItem(CareNetServices.XMLMaintenanceControlTypes.xmctEventRoomBookings, vParams)
                    End If
                  Catch vEx As CareException
                    If vEx.ErrorNumber = CareException.ErrorNumbers.enCancellationFeeMissing Then
                      'Needs to get the amount from the transaction
                      Dim vRow As DataRow = DataHelper.GetRowFromDataSet(DataHelper.GetTransactionData(CareServices.XMLTransactionDataSelectionTypes.xtdtTransactionAnalysis, SelectRowItemNumber(dgr.CurrentRow, "BatchNumber"), SelectRowItemNumber(dgr.CurrentRow, "TransactionNumber"), 0, SelectRowItemNumber(dgr.CurrentRow, "LineNumber")))
                      Dim vDefaults As New ParameterList
                      If vRow IsNot Nothing Then vDefaults("CancellationFeeAmount") = vRow("Amount").ToString
                      Dim vCancellationFeeParams As ParameterList = FormHelper.ShowApplicationParameters(CareServices.FunctionParameterTypes.fptEnterCancellationFee, vDefaults)
                      If vCancellationFeeParams IsNot Nothing AndAlso vCancellationFeeParams.Count > 0 Then
                        vInvoiceParams("CancellationFeeAmount") = vCancellationFeeParams("CancellationFeeAmount")
                        Try
                          If vEventBooking Then
                            vReturnList = DataHelper.CancelEventBooking(vInvoiceParams)
                          Else
                            vReturnList = DataHelper.UpdateItem(CareNetServices.XMLMaintenanceControlTypes.xmctEventRoomBookings, vInvoiceParams)
                          End If
                        Catch vCareEx As CareException
                          If vCareEx.ErrorNumber = CareException.ErrorNumbers.enCCAuthorisationFailed OrElse vCareEx.ErrorNumber = CareException.ErrorNumbers.enCardAuthorisationUnexpectedTimeout Then
                            ShowInformationMessage(vCareEx.Message)
                          Else
                            Throw vCareEx
                          End If
                        End Try
                      Else
                        Exit Sub
                      End If
                    ElseIf vEx.ErrorNumber = CareException.ErrorNumbers.enCCAuthorisationFailed OrElse vEx.ErrorNumber = CareException.ErrorNumbers.enCardAuthorisationUnexpectedTimeout Then
                      ShowInformationMessage(vEx.Message)
                    Else
                      Throw vEx
                    End If
                  End Try
                  If vReturnList.Contains("Message") Then
                    ShowInformationMessage(vReturnList("Message"))
                  ElseIf vReturnList.Contains("ProcessWaitingList") Then
                    Dim vEventInfo As New CareEventInfo(DataHelper.GetEventData(CareServices.XMLEventDataSelectionTypes.xedtEventInformation, SelectRowItemNumber(dgr.CurrentRow, "EventNumber")).Tables("DataRow").Rows(0))
                    Dim vWaiting As New frmWaitingList(vEventInfo)
                    vWaiting.ShowDialog(Me)
                  End If
                  If vEventBooking Then
                    ShowInformationMessage(InformationMessages.ImEventBookingCancelled)
                  Else
                    ShowInformationMessage(InformationMessages.ImEventRoomBookingCancelled)
                  End If
                End If
              End If

            Case CareServices.XMLContactDataSelectionTypes.xcdtContactGiftAidDeclarations
              Dim vDefaults As New ParameterList
              vDefaults("ExcludeCancellationReason") = AppValues.ControlValue(AppValues.ControlValues.merge_cancellation_reason)
              'Set up the new End Date - current End Date, Todays Date or current Start Date, whichever is the earlier
              Dim vEndDate As Date
              Dim vStartDate As Date
              If Date.TryParse(SelectRowItem(dgr.CurrentRow, "EndDate"), vEndDate) Then
                If vEndDate > Now.Date Then vEndDate = Now.Date
              Else
                vEndDate = Now.Date
              End If
              vStartDate = Date.Parse(SelectRowItem(dgr.CurrentRow, "StartDate"))
              If vStartDate > vEndDate Then vEndDate = vStartDate
              vDefaults("CancelledOn") = vEndDate.ToString
              vDefaults("CancelledOnMaxDate") = vEndDate.ToString
              vDefaults("CancelledOnMinDate") = vStartDate.ToString
              Dim vRenameList As New ParameterList
              vRenameList("CancelledOn") = "End Date:"
              Dim vParams As ParameterList = FormHelper.ShowApplicationParameters(CareServices.FunctionParameterTypes.fptCancellationReasonSourceAndDate, vDefaults, vRenameList)
              If vParams IsNot Nothing AndAlso vParams.Count > 0 Then
                vParams.IntegerValue("DeclarationNumber") = SelectRowItemNumber(dgr.CurrentRow, "DeclarationNumber")
                Dim vTryAgain As Boolean = False
                Do
                  vTryAgain = False
                  Dim vReturnList As ParameterList = DataHelper.CancelGiftAidDeclaration(vParams)
                  If vReturnList.Contains("AdjustClaimedPayments") AndAlso vReturnList("AdjustClaimedPayments") = "Y" Then
                    If ShowQuestion(QuestionMessages.QmAdjustClaimedPayments, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
                      vTryAgain = True
                      vParams("DoCancel") = "Y" 'Indicates that all the previously done checks should be bypassed and the GAD can be cancelled
                    ElseIf vReturnList.Contains("Message") AndAlso vReturnList("Message").Length > 0 Then
                      ShowInformationMessage(vReturnList("Message"))
                    End If
                  End If
                Loop While vTryAgain
                vRefresh = True
              End If

            Case CareServices.XMLContactDataSelectionTypes.xcdtContactAppropriateCertificates
              Dim vDefaults As New ParameterList
              vDefaults("ExcludeCancellationReason") = AppValues.ControlValue(AppValues.ControlValues.merge_cancellation_reason)
              vDefaults("CancelledOn") = AppValues.TodaysDate
              Dim vParams As ParameterList = FormHelper.ShowApplicationParameters(CareServices.FunctionParameterTypes.fptCancellationReasonAndSource, vDefaults)
              If vParams IsNot Nothing AndAlso vParams.Count > 0 Then
                vParams.IntegerValue("CertificateNumber") = SelectRowItemNumber(dgr.CurrentRow, "CertificateNumber")
                Dim vReturnList As ParameterList = DataHelper.CancelAppropriateCertificate(vParams)
                If vReturnList.Contains("Message") AndAlso vReturnList("Message").Length > 0 Then ShowInformationMessage(vReturnList("Message"))
                vRefresh = True
              End If

            Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactLoans
              CancelPaymentPlan(CareNetServices.XMLPaymentPlanMenuTypes.xpmtCancelLoan, vRefresh)

            Case CareServices.XMLContactDataSelectionTypes.xcdtContactMembershipDetails
              Dim vReturnList As ParameterList = CancelPaymentPlan(CareServices.XMLPaymentPlanMenuTypes.xpmtCancelMember, vRefresh)
              If vReturnList.Contains("CancelledMemberIsPayer") Then ShowInformationMessage(InformationMessages.ImMemberIsPayer)

            Case CareServices.XMLContactDataSelectionTypes.xcdtContactPaymentPlans
              CancelPaymentPlan(CareServices.XMLPaymentPlanMenuTypes.xpmtCancelPaymentPlan, vRefresh)

            Case CareServices.XMLContactDataSelectionTypes.xcdtContactPledges
              Dim vDefaults As New ParameterList
              vDefaults("ExcludeCancellationReason") = AppValues.ControlValue(AppValues.ControlValues.merge_cancellation_reason)
              vDefaults("CancelledOn") = AppValues.TodaysDate
              Dim vParams As ParameterList = FormHelper.ShowApplicationParameters(CareServices.FunctionParameterTypes.fptCancellationReasonAndSource, vDefaults)
              If vParams IsNot Nothing AndAlso vParams.Count > 0 Then
                vParams.IntegerValue("GayePledgeNumber") = SelectRowItemNumber(dgr.CurrentRow, "PledgeNumber")
                Dim vReturnList As ParameterList = DataHelper.UpdateItem(CareServices.XMLMaintenanceControlTypes.xmctPreTaxPledges, vParams)
                vRefresh = True
              End If

            Case CareServices.XMLContactDataSelectionTypes.xcdtContactPostPGPledges
              Dim vDefaults As New ParameterList
              vDefaults("ExcludeCancellationReason") = AppValues.ControlValue(AppValues.ControlValues.merge_cancellation_reason)
              vDefaults("CancelledOn") = AppValues.TodaysDate
              Dim vParams As ParameterList = FormHelper.ShowApplicationParameters(CareServices.FunctionParameterTypes.fptCancellationReasonAndSource, vDefaults)
              If vParams IsNot Nothing AndAlso vParams.Count > 0 Then
                vParams.IntegerValue("PledgeNumber") = SelectRowItemNumber(dgr.CurrentRow, "PledgeNumber")
                Dim vReturnList As ParameterList = DataHelper.UpdateItem(CareServices.XMLMaintenanceControlTypes.xmctPostTaxPledges, vParams)
                vRefresh = True
              End If

            Case CareServices.XMLContactDataSelectionTypes.xcdtContactStandingOrders
              CancelPaymentPlan(CareServices.XMLPaymentPlanMenuTypes.xpmtCancelStandingOrder, vRefresh)

            Case CareServices.XMLContactDataSelectionTypes.xcdtContactUnProcessedTransactions,
             CareServices.XMLContactDataSelectionTypes.xcdtContactDeliveryTransactions,
             CareServices.XMLContactDataSelectionTypes.xcdtContactSalesTransactions
              Dim vReturnList As ParameterList = DataHelper.CancelReinstateProvisionalTransaction(IntegerValue(pDataRow("BatchNumber")), IntegerValue(pDataRow("TransactionNumber")), True)
              vRefresh = True
          End Select
        Case BaseFinancialMenu.FinancialMenuItems.fmiChangeExamCentre, FinancialMenu.FinancialMenuItems.fmiChangeExamCentre
          Select Case mvDataType
            Case CareServices.XMLContactDataSelectionTypes.xcdtContactEventBookings, CareNetServices.XMLContactDataSelectionTypes.xcdtContactEventRoomBookings, CareNetServices.XMLContactDataSelectionTypes.xcdtContactExamDetails
              Dim vCentreParams As New ParameterList(True)
              vCentreParams("ExamSessionId") = dgr.GetValue(dgr.CurrentRow, "ExamSessionId").ToString
              Dim vParams As ParameterList = FormHelper.ShowApplicationParameters(CareServices.FunctionParameterTypes.fptExamChangeCentre, vCentreParams)
              If vParams.Count > 0 Then
                vParams("ExamBookingId") = dgr.GetValue(dgr.CurrentRow, "ExamBookingId").ToString
                vParams("ContactNumber") = mvContactInfo.ContactNumber.ToString()
                ' Call web service 
                Dim vReturnList As New ParameterList
                vReturnList = ExamsDataHelper.ChangeExamBookingCentre(vParams)
                vRefresh = True
              End If
          End Select

        Case FinancialMenu.FinancialMenuItems.fmiChangeMembershipType
          Dim vCaption As String = ""
          If FormHelper.IsTraderLoaded(vCaption) Then
            ShowInformationMessage(InformationMessages.ImFormAlreadyOpen, vCaption)
          Else
            If BooleanValue(SelectRowItem(dgr.CurrentRow, "ApprovalMembership")) = True AndAlso DoubleValue(SelectRowItem(dgr.CurrentRow, "Balance")) <> 0 Then
              ShowInformationMessage(InformationMessages.ImCMTApprovalMemberBalance)
            Else
              Dim vMemberNumber As String = SelectRowItem(dgr.CurrentRow, "MemberNumber")
              If vMemberNumber.Length > 0 Then
                Dim vList As New ParameterList
                vList("MemberNumber") = vMemberNumber
                vList("MembershipNumber") = SelectRowItem(dgr.CurrentRow, "MembershipNumber")
                Dim vTraderApplication As New TraderApplication(IntegerValue(AppValues.ConfigurationValue(AppValues.ConfigurationValues.trader_application_cmt)))
                FormHelper.RunTraderApplication(vTraderApplication, vList)
              End If
            End If
          End If

        Case FinancialMenu.FinancialMenuItems.fmiChangePayer
          Dim vResult As DialogResult = System.Windows.Forms.DialogResult.Yes
          If pChangeDetails Then
            vResult = ShowQuestion(QuestionMessages.QmChangePPDsPayer, MessageBoxButtons.YesNoCancel)
            If vResult <> System.Windows.Forms.DialogResult.Cancel Then pChangeDetails = vResult = System.Windows.Forms.DialogResult.Yes
          End If
          If vResult <> System.Windows.Forms.DialogResult.Cancel Then
            Dim vParamList As New ParameterList(True)
            vParamList.IntegerValue("PaymentPlanNumber") = SelectRowItemNumber(dgr.CurrentRow, "PaymentPlanNumber")
            vParamList("ChangeDetails") = CBoolYN(pChangeDetails)
            'Find new payer contact
            vParamList.IntegerValue("ContactNumber") = FormHelper.ShowFinder(CareServices.XMLDataFinderTypes.xdftContacts, Me, True)
            If vParamList.IntegerValue("ContactNumber") > 0 Then
              Dim vContactInfo As New ContactInfo(vParamList.IntegerValue("ContactNumber"))
              vParamList.IntegerValue("AddressNumber") = vContactInfo.AddressNumber
              Dim vType As String = ""
              Dim vGifted As Boolean
              Dim vOneYear As Boolean
              'Pass the new payer's contact number to web service to get Change Payer Membership options
              Dim vDT As DataTable = DataHelper.GetPaymentPlanData(CareServices.XMLPaymentPlanDataSelectionTypes.xpdtPaymentPlanChangePayer, vParamList.IntegerValue("PaymentPlanNumber"), vContactInfo.ContactNumber).Tables("DataRow")
              For Each vDR As DataRow In vDT.Rows
                Select Case vDR.Item("Option").ToString
                  Case "GiftMembershipType"
                    vType = vDR.Item("Setting").ToString
                  Case "GiftedMembership"
                    vGifted = BooleanValue(vDR.Item("Setting").ToString)
                  Case "OneYearGift"
                    vOneYear = BooleanValue(vDR.Item("Setting").ToString)
                End Select
              Next
              'Ask the user some questions
              If vType = "Gift" Then
                'Payment plan is a membership that will become gifted when the payer is changed.
                vParamList("GiftMembership") = "Y"
                If vGifted Then
                  'Payment plan is already a gifted membership
                  If vOneYear Then
                    'Payment plan is already a one-year, gifted membership
                    vResult = ShowQuestion(QuestionMessages.QmRetainOneYearGift, MessageBoxButtons.YesNoCancel)
                  Else
                    vResult = ShowQuestion(QuestionMessages.QmMakeAlreadyGiftedOneYearGift, MessageBoxButtons.YesNoCancel)
                  End If
                Else
                  'Payment plan is not currently a gifted membership
                  vResult = ShowQuestion(QuestionMessages.QmContinueChangePayer, MessageBoxButtons.OKCancel)
                  If vResult = System.Windows.Forms.DialogResult.OK Then vResult = ShowQuestion(QuestionMessages.QmMakeOneYearGift, MessageBoxButtons.YesNoCancel)
                End If
                If vResult = System.Windows.Forms.DialogResult.Yes Then vParamList("OneYearGift") = "Y"
              ElseIf vType = "NonGift" Then
                'Payment plan is a membership that will not become gifted when the payer is changed.
                vParamList("GiftMembership") = "N"
              Else
                'Payment plan is not a membership
              End If
              'Change the Payer
              If vResult <> System.Windows.Forms.DialogResult.Cancel Then
                Dim vReturnList As ParameterList = DataHelper.ProcessPaymentPlanMenu(CareServices.XMLPaymentPlanMenuTypes.xpmtChangePayer, vParamList)
                If vReturnList.IntegerValue("PaymentPlanNumber") > 0 Then ShowInformationMessage(InformationMessages.ImPaymentPlanPayerChanged)
                vRefresh = True
              End If
            End If
          End If

        Case FinancialMenu.FinancialMenuItems.fmiEditNotes, FinancialMenu.FinancialMenuItems.fmiEditUnprocessedTransactionNotes, BaseFinancialMenu.FinancialMenuItems.fmiEditReference
          Dim vform As frmCardMaintenance

          If pItem = BaseFinancialMenu.FinancialMenuItems.fmiEditReference Then
            vform = New frmCardMaintenance(Me, mvContactInfo, mvDataType, mvDataSet, True, dgr.CurrentRow, CareNetServices.XMLMaintenanceControlTypes.xmctPostedBatchReference)
            vform.Show()
          Else
            vform = New frmCardMaintenance(Me, mvContactInfo, mvDataType, mvDataSet, True, dgr.CurrentRow)
            vform.Show()
          End If
        Case FinancialMenu.FinancialMenuItems.fmiFutureCancel
          Select Case mvDataType
            Case CareServices.XMLContactDataSelectionTypes.xcdtContactCreditCardAuthorities
              CancelPaymentPlan(CareServices.XMLPaymentPlanMenuTypes.xpmtFutureCancelCreditCardAuthority, vRefresh)
            Case CareServices.XMLContactDataSelectionTypes.xcdtContactDirectDebits
              CancelPaymentPlan(CareServices.XMLPaymentPlanMenuTypes.xpmtFutureCancelDirectDebit, vRefresh)
            Case CareServices.XMLContactDataSelectionTypes.xcdtContactPaymentPlans
              CancelPaymentPlan(CareServices.XMLPaymentPlanMenuTypes.xpmtFutureCancelPaymentPlan, vRefresh)
            Case CareServices.XMLContactDataSelectionTypes.xcdtContactStandingOrders
              CancelPaymentPlan(CareServices.XMLPaymentPlanMenuTypes.xpmtFutureCancelStandingOrder, vRefresh)
          End Select

        Case FinancialMenu.FinancialMenuItems.fmiFutureMembershipType
          Dim vList As New ParameterList(True)
          vList("SystemColumns") = "N"
          vList.IntegerValue("MembershipNumber") = SelectRowItemNumber(dgr.CurrentRow, "MembershipNumber")
          Dim vDataRow As DataRow = DataHelper.GetContactItem(CareServices.XMLContactDataSelectionTypes.xcdtContactMembershipDetails, mvContactInfo.ContactNumber, vList)
          If vDataRow IsNot Nothing Then
            vList.IntegerValue("PaymentPlanNumber") = IntegerValue(vDataRow.Item("PaymentPlanNumber").ToString)
            'Execute the following line of code in order to perform some validation before the FMT is allowed to be created.  This call will either return "OK" or will raise an error, which is handled in the Catch below.
            Dim vReturnList As ParameterList = DataHelper.ProcessPaymentPlanMenu(CareServices.XMLPaymentPlanMenuTypes.xpmtInitialFMTValidation, vList)
            vList("MembershipType") = vDataRow.Item("MembershipType").ToString
            vList("PaymentPlanMembershipType") = vDataRow.Item("PaymentPlanMembershipType").ToString
            vList("Existing") = "N"
            If vDataRow.Item("FutureMembershipType").ToString.Length > 0 Then
              'Existing FMT
              vList("FutureMembershipType") = vDataRow.Item("FutureMembershipType").ToString
              vList("Existing") = "Y"
            ElseIf vDataRow.Item("SubsequentMembershipType").ToString.Length > 0 Then
              'FMT doesn't exist, set default FMT
              vList("FutureMembershipType") = vDataRow.Item("SubsequentMembershipType").ToString
            End If
            If vDataRow.Item("FutureMembershipType").ToString.Length > 0 Then
              vList("FutureChangeDate") = vDataRow.Item("FutureChangeDate").ToString
            End If
            Dim vJoined As Date = Date.Parse(vDataRow.Item("Joined").ToString)
            Dim vRenewalDate As Date = Date.Parse(vDataRow.Item("RenewalDate").ToString)
            If vJoined = vRenewalDate Then
              vList("CalculatedFutureChangeDate") = vRenewalDate.AddYears(1).ToString
            Else
              vList("CalculatedFutureChangeDate") = vRenewalDate.ToString
            End If
            If Not vList.Contains("FutureChangeDate") Then vList("FutureChangeDate") = vList("CalculatedFutureChangeDate")
            vList("Product") = vDataRow.Item("FutureProduct").ToString
            vList("Rate") = vDataRow.Item("FutureRate").ToString
            vList("Amount") = vDataRow.Item("FutureAmount").ToString
            vList("SubsequentMembershipType") = vDataRow.Item("SubsequentMembershipType").ToString
            vList("SubsequentMembershipTypeChangeDate") = vDataRow.Item("SubsequentMembershipTypeChangeDate").ToString
            vList("PayerRequired") = vDataRow.Item("PayerRequired").ToString
            vList("Annual") = vDataRow.Item("Annual").ToString
            vList("MembershipTerm") = vDataRow.Item("MembershipTerm").ToString
            Dim vForm As New frmCardMaintenance(CareServices.XMLMaintenanceControlTypes.xmctFutureMembershipType, SelectRowItemNumber(dgr.CurrentRow, "MembershipNumber"), vList, mvContactInfo)
            AddHandler vForm.ShowApplicationParameters, AddressOf ShowApplicationParametersHandler
            Dim vDialogResult As DialogResult = vForm.ShowDialog(Me)
            vRefresh = vDialogResult <> DialogResult.Cancel   'Delete FutureMembershipType will now return the DialogResult as OK so that we can refresh
          End If

        Case BaseFinancialMenu.FinancialMenuItems.fmiFutureRenewalAmount
          Dim vList As New ParameterList(True, True)
          vList.IntegerValue("MembershipNumber") = SelectRowItemNumber(dgr.CurrentRow, "MembershipNumber")
          vList("FutureMembershipType") = SelectRowItem(dgr.CurrentRow, "FutureMembershipType")
          Dim vDR As DataRow = DataHelper.GetMembershipItem(CareNetServices.XMLMembershipDataSelectionTypes.xmdtMembershipFutureRenewalAmount, vList)
          If vDR IsNot Nothing AndAlso vDR.Table.Columns.Contains("FutureRenewalAmount") Then
            Dim vFutureRenewalAmount As Double = DoubleValue(vDR.Item("FutureRenewalAmount").ToString)
            Dim vMessage As New StringBuilder()
            vMessage.AppendLine(GetInformationMessage(InformationMessages.ImFutureRenewalAmountMsg, vFutureRenewalAmount.ToString("0.00")))
            vMessage.AppendLine("")
            vMessage.AppendLine(InformationMessages.ImFutureRenewalAmountDisclaimer)
            ShowInformationMessage(vMessage.ToString)
          End If

        Case FinancialMenu.FinancialMenuItems.fmiGoToBankAccount
          GoToBankAccount(IntegerValue(pDataRow("BankDetailsNumber")), vRefresh)
        Case FinancialMenu.FinancialMenuItems.fmiGoToChangedBy
          SelectTransaction(mvFinancialMenu.AdjustmentBatchNumber, mvFinancialMenu.AdjustmentTransactionNumber)
        Case FinancialMenu.FinancialMenuItems.fmiGoToChanges
          SelectTransaction(mvFinancialMenu.AdjustmentWasBatchNumber, mvFinancialMenu.AdjustmentWasTransactionNumber)
        Case FinancialMenu.FinancialMenuItems.fmiGoToCC
          GoToPaymentPlanType(CareServices.XMLContactDataSelectionTypes.xcdtContactCreditCardAuthorities, pDataRow, vRefresh)
        Case FinancialMenu.FinancialMenuItems.fmiGoToCovenant
          GoToPaymentPlanType(CareServices.XMLContactDataSelectionTypes.xcdtContactCovenants, pDataRow, vRefresh)
        Case FinancialMenu.FinancialMenuItems.fmiGoToDD
          GoToPaymentPlanType(CareServices.XMLContactDataSelectionTypes.xcdtContactDirectDebits, pDataRow, vRefresh)
        Case FinancialMenu.FinancialMenuItems.fmiGoToEvent
          FormHelper.ShowEventIndex(mvFinancialMenu.EventNumber(0))
        Case FinancialMenu.FinancialMenuItems.fmiGoToMembership
          GoToPaymentPlanType(CareServices.XMLContactDataSelectionTypes.xcdtContactMembershipDetails, pDataRow, vRefresh)
        Case FinancialMenu.FinancialMenuItems.fmiGoToPayPlan
          GoToPaymentPlanType(CareServices.XMLContactDataSelectionTypes.xcdtContactPaymentPlans, pDataRow, vRefresh)
        Case FinancialMenu.FinancialMenuItems.fmiGoToSO
          GoToPaymentPlanType(CareServices.XMLContactDataSelectionTypes.xcdtContactStandingOrders, pDataRow, vRefresh)
        Case BaseFinancialMenu.FinancialMenuItems.fmiGoToLoan
          GoToPaymentPlanType(CareServices.XMLContactDataSelectionTypes.xcdtContactLoans, pDataRow, vRefresh)
        Case FinancialMenu.FinancialMenuItems.fmiPaymentPlanConversion, FinancialMenu.FinancialMenuItems.fmiPaymentPlanMaintenance
          Dim vPPNumber As String = SelectRowItem(dgr.CurrentDataRow, "PaymentPlanNumber")
          If vPPNumber.Length > 0 Then
            Dim vList As New ParameterList
            vList("PaymentPlanNumber") = vPPNumber
            Dim vTraderApplication As TraderApplication
            If pItem = FinancialMenu.FinancialMenuItems.fmiPaymentPlanConversion Then
              vTraderApplication = New TraderApplication(IntegerValue(AppValues.ConfigurationValue(AppValues.ConfigurationValues.trader_application_conversion)))
            Else
              vTraderApplication = New TraderApplication(IntegerValue(AppValues.ConfigurationValue(AppValues.ConfigurationValues.trader_application_maintenance)))
            End If
            FormHelper.RunTraderApplication(vTraderApplication, vList)
          End If
        Case FinancialMenu.FinancialMenuItems.fmiPaymentPlanPrint
          Dim vPPNumber As String = SelectRowItem(dgr.CurrentRow, "PaymentPlanNumber")
          If vPPNumber.Length > 0 Then
            Dim vParamList As ParameterList = FormHelper.ShowApplicationParameters(CareNetServices.FunctionParameterTypes.fptPaymentPlanDocument)
            If vParamList.ContainsKey("StandardDocument") Then
              Dim vList As New ParameterList(True)
              vList("PaymentPlanNumber") = vPPNumber
              Dim vFileName As String = DataHelper.GetTempFile(".csv")
              If DataHelper.GetTraderMailingFile(CareNetServices.TraderMailmergeType.tmtPaymentPlan, vList, vFileName) Then
                'We have the mailmerge file, now need to perform the mailmerge
                vList = New ParameterList(True)
                vList("StandardDocument") = vParamList("StandardDocument")
                Dim vRow As DataRow = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtStandardDocuments, vList).Rows(0)
                Dim vApplication As ExternalApplication = GetDocumentApplication(vRow.Item("DocFileExtension").ToString)
                vApplication.MergeStandardDocument(vRow.Item("StandardDocument").ToString, vRow.Item("DocFileExtension").ToString, vFileName, True)
              End If
            End If
          End If

        Case FinancialMenu.FinancialMenuItems.fmiProduceMembershipCard
          Dim vReprintCardSet As Boolean = BooleanValue(GetDataRow(mvDataSet, dgr.CurrentRow).Item("ReprintCard").ToString)
          Dim vMembershipNumber As Integer = IntegerValue(GetDataRow(mvDataSet, dgr.CurrentRow).Item("MembershipNumber").ToString)
          'if reprint membership flag is not set, then set it so that the member will get picked up
          If Not vReprintCardSet Then
            Dim vMemberList As New ParameterList(True, True)
            vMemberList.IntegerValue("MembershipNumber") = vMembershipNumber
            vMemberList("ReprintMembershipCard") = "Y"
            DataHelper.UpdateItem(CareServices.XMLMaintenanceControlTypes.xmctMembership, vMemberList)
            vRefresh = True
          End If
          'Set up and run the membership card production for the selected member
          Dim vMailing As New GeneralMailing(CareNetServices.MailingTypes.mtyMembershipCards, CareNetServices.TaskJobTypes.tjtMembCardMailing)
          Dim vList As New ParameterList(True)  'create selection set
          vList("SelectionSetDesc") = "Single Membership Card Production"
          vList("NumberInMailing") = "1"
          vList("SelectionSetNumber") = vMailing.MailingInfo.SelectionSet.ToString
          Dim vReturn As ParameterList = DataHelper.AddItem(CareNetServices.XMLMaintenanceControlTypes.xmctSelectionSet, vList)
          vMailing.MailingInfo.SelectionSet = IntegerValue(vReturn("SelectionSetNumber"))
          DataHelper.AddSelectionSetContact(vMailing.MailingInfo.SelectionSet, ContactInfo)  'add selected contact to the selection set
          vList = New ParameterList(True)  'create criteria set
          vList("CriteriaSet") = "0"
          vList("SearchArea") = "membnumber"
          vList("SequenceNumber") = "1"
          vList("MainValue") = vMembershipNumber.ToString
          vList("IE") = "I"
          vList("CO") = ContactInfo.ContactTypeCode
          Dim vReturnList As ParameterList = DataHelper.AddCriteriaSetDetails(vList)
          vMailing.MailingInfo.CriteriaSet = IntegerValue(vReturnList("CriteriaSet").ToString)
          Dim vFrm As New frmGenMGen("MC", vMailing.MailingInfo, vMailing.MailingInfo.SelectionSet, True) 'show the form where standard doc can be selected
          vFrm.ShowDialog()

        Case FinancialMenu.FinancialMenuItems.fmiReinstateAutoPayMethod
          Dim vTryAgain As Boolean = False
          Dim vList As New ParameterList(True)
          Dim vReturnList As ParameterList
          Do
            vTryAgain = False
            If Not vList.Contains("PaymentPlanNumber") Then
              vList.IntegerValue("PaymentPlanNumber") = SelectRowItemNumber(dgr.CurrentRow, "PaymentPlanNumber")
              Dim vType As String = ""
              Dim vColumnName As String = ""
              Select Case mvDataType
                Case CareServices.XMLContactDataSelectionTypes.xcdtContactCreditCardAuthorities
                  vType = "A"
                  vColumnName = "CreditCardAuthorityNumber"
                Case CareServices.XMLContactDataSelectionTypes.xcdtContactDirectDebits
                  vType = "D"
                  vColumnName = "DirectDebitNumber"
                Case CareServices.XMLContactDataSelectionTypes.xcdtContactStandingOrders
                  vType = "B"
                  vColumnName = "StandingOrderNumber"
              End Select
              vList("PaymentPlanType") = vType
              vList.IntegerValue(vColumnName) = SelectRowItemNumber(dgr.CurrentRow, vColumnName)
            End If
            vReturnList = DataHelper.ProcessPaymentPlanMenu(CareServices.XMLPaymentPlanMenuTypes.xpmtReinstateAutoPayMethod, vList)
            If vReturnList.Contains("SkipPaymentCount") AndAlso vReturnList.IntegerValue("SkipPaymentCount") > 0 Then
              'Ask user to confirm the skipping of these payments
              Dim vDefaults As New ParameterList
              vDefaults.IntegerValue("PaymentNumber") = vReturnList.IntegerValue("SkipPaymentCount")
              vDefaults("Amount") = vReturnList("SkipPaymentValue")
              Dim vAppParamsList As ParameterList = FormHelper.ShowApplicationParameters(CareServices.FunctionParameterTypes.fptPayPlanMissedPayments, vDefaults)
              If vAppParamsList IsNot Nothing AndAlso vAppParamsList.Count > 0 Then
                'Continue by adding the CheckedHistoricAddresses and SkipPaymentCount parameters to vList and looping around to call the web service again
                vTryAgain = True
                vList("CheckedHistoricAddresses") = vReturnList("CheckedHistoricAddresses")
                vList("CheckedForCancelledPaymentPlan") = vReturnList("CheckedForCancelledPaymentPlan")
                vList("CheckedForAutoPaymentMethod") = vReturnList("CheckedForAutoPaymentMethod")
                If vAppParamsList("Checkbox") = "Y" Then
                  vList.IntegerValue("SkipPaymentCount") = vReturnList.IntegerValue("SkipPaymentCount")
                Else
                  vList.IntegerValue("SkipPaymentCount") = 0
                End If
              Else
                'cancel the process
                ShowInformationMessage(InformationMessages.ImPaymentMethodReinstatementCancelled, vReturnList("AutoPaymentMethod"))
              End If
            ElseIf vReturnList.IntegerValue("PaymentPlanNumber") > 0 Then
              ShowInformationMessage(InformationMessages.ImPaymentMethodReinstated, vReturnList("AutoPaymentMethod"))
              vRefresh = True
            End If
          Loop While vTryAgain

        Case FinancialMenu.FinancialMenuItems.fmiReinstateMembership
          Dim vTryAgain As Boolean = False
          Dim vList As New ParameterList(True)
          Dim vReturnList As ParameterList
          Do
            vTryAgain = False
            If Not vList.Contains("PaymentPlanNumber") Then
              vList.IntegerValue("PaymentPlanNumber") = SelectRowItemNumber(dgr.CurrentRow, "PaymentPlanNumber")
              vList.IntegerValue("MembershipNumber") = SelectRowItemNumber(dgr.CurrentRow, "MembershipNumber")
              vList.IntegerValue("ContactNumber") = mvContactInfo.ContactNumber
            End If
            vReturnList = DataHelper.ProcessPaymentPlanMenu(CareServices.XMLPaymentPlanMenuTypes.xpmtReinstateMembership, vList)
            If vReturnList.Contains("Balance") AndAlso DoubleValue(vReturnList("Balance")) > 0 Then
              'Ask user to confirm the outstanding balance of the membership
              Dim vDefaults As New ParameterList
              vDefaults("Balance") = vReturnList("Balance")
              vDefaults("FrequencyAmount") = vReturnList("FrequencyAmount")
              Dim vAppParamsList As ParameterList = FormHelper.ShowApplicationParameters(CareServices.FunctionParameterTypes.fptMembershipReinstatement, vDefaults)
              If vAppParamsList IsNot Nothing AndAlso vAppParamsList.Count > 0 Then
                'Continue by adding the CheckedHistoricAddresses and SkipPaymentCount parameters to vList and looping around to call the web service again
                vTryAgain = True
                vList("CheckedHistoricAddresses") = vReturnList("CheckedHistoricAddresses")
                vList("CheckedCurrentMembers") = vReturnList("CheckedCurrentMembers")
                vList("Balance") = vAppParamsList("Balance")
              Else
                'cancel the process
                ShowInformationMessage(InformationMessages.ImMembershipReinstatementCancelled)
              End If
            ElseIf vReturnList.IntegerValue("PaymentPlanNumber") > 0 Then
              ShowInformationMessage(InformationMessages.ImMembershipReinstated)
              vRefresh = True
            End If
          Loop While vTryAgain

        Case FinancialMenu.FinancialMenuItems.fmiReinstatePaymentPlan
          Dim vTryAgain As Boolean = False
          Dim vList As New ParameterList(True)
          Dim vReturnList As ParameterList
          Do
            vTryAgain = False
            If Not vList.Contains("PaymentPlanNumber") Then vList.IntegerValue("PaymentPlanNumber") = SelectRowItemNumber(dgr.CurrentRow, "PaymentPlanNumber")
            vReturnList = DataHelper.ProcessPaymentPlanMenu(CareServices.XMLPaymentPlanMenuTypes.xpmtReinstatePaymenPlan, vList)
            If vReturnList.Contains("SkipPaymentCount") AndAlso vReturnList.IntegerValue("SkipPaymentCount") > 0 Then
              'Ask user to confirm the skipping of these payments
              Dim vDefaults As New ParameterList
              vDefaults.IntegerValue("PaymentNumber") = vReturnList.IntegerValue("SkipPaymentCount")
              vDefaults("Amount") = vReturnList("SkipPaymentValue")
              Dim vAppParamsList As ParameterList = FormHelper.ShowApplicationParameters(CareServices.FunctionParameterTypes.fptPayPlanMissedPayments, vDefaults)
              If vAppParamsList IsNot Nothing AndAlso vAppParamsList.Count > 0 Then
                'Continue by adding the CheckedHistoricAddresses and SkipPaymentCount parameters to vList and looping around to call the web service again
                vTryAgain = True
                vList("CheckedHistoricAddresses") = vReturnList("CheckedHistoricAddresses")
                If vAppParamsList("Checkbox") = "Y" Then
                  vList.IntegerValue("SkipPaymentCount") = vReturnList.IntegerValue("SkipPaymentCount")
                Else
                  vList.IntegerValue("SkipPaymentCount") = 0
                End If
              Else
                'cancel the process
                ShowInformationMessage(InformationMessages.ImPaymentPlanReinstatementCancelled)
              End If
            ElseIf vReturnList.IntegerValue("PaymentPlanNumber") > 0 Then
              Dim vSB As New StringBuilder
              Dim vSB2 As New StringBuilder
              If vReturnList.Contains("Member") Then vSB.Append(", Member")
              If vReturnList.Contains("Loan") Then vSB.Append(", Loan")
              If vReturnList.Contains("Covenant") Then vSB.Append(", Covenant")
              If vReturnList.Contains("AutoPaymentMethod") Then
                vSB.Append(", ")
                vSB.Append(vReturnList("AutoPaymentMethod"))
              End If
              If vReturnList.Contains("FMTMsg") Then
                vSB2.AppendLine()
                vSB2.AppendLine()
                vSB2.Append(vReturnList("FMTMsg"))
              End If
              ShowInformationMessage(InformationMessages.ImPaymentPlanReinstated, vSB.ToString, vSB2.ToString)
              vRefresh = True
            End If
          Loop While vTryAgain

        Case FinancialMenu.FinancialMenuItems.fmiRemoveAllocations
          RemoveInvoiceAllocations(mvDataSet2, dts.DisplayGrid(0).CurrentRow, vRefresh)

        Case FinancialMenu.FinancialMenuItems.fmiReprintMembershipCard
          Dim vReprintCardSet As Boolean = BooleanValue(GetDataRow(mvDataSet, dgr.CurrentRow).Item("ReprintCard").ToString)
          Dim vMembershipNumber As Integer = IntegerValue(GetDataRow(mvDataSet, dgr.CurrentRow).Item("MembershipNumber").ToString)
          Dim vMemberNumber As Integer = IntegerValue(GetDataRow(mvDataSet, dgr.CurrentRow).Item("MemberNumber").ToString)
          Dim vMsg As String = QuestionMessages.QmSetReprintMshipCard
          If vReprintCardSet Then vMsg = QuestionMessages.QmClearReprintMshipCard
          'BR16573 Reprint Membership card now displayes the Member Number
          If ShowQuestion(vMsg, MessageBoxButtons.YesNo, vMemberNumber.ToString) = System.Windows.Forms.DialogResult.Yes Then
            Dim vMemberList As New ParameterList(True, True)
            vMemberList.IntegerValue("MembershipNumber") = vMembershipNumber
            vMemberList("ReprintMembershipCard") = CBoolYN(Not vReprintCardSet)
            DataHelper.UpdateItem(CareServices.XMLMaintenanceControlTypes.xmctMembership, vMemberList)
            vRefresh = True
          End If

        Case FinancialMenu.FinancialMenuItems.fmiReverse, BaseFinancialMenu.FinancialMenuItems.fmiMove, BaseFinancialMenu.FinancialMenuItems.fmiAnalysis, BaseFinancialMenu.FinancialMenuItems.fmiRefund
          Dim vTransDate As String = ""
          Dim vTransSign As String = ""
          Dim vStock As Boolean
          Dim vEventNumber As Integer
          Dim vEventAnalysis As Boolean = False
          Dim vList As New ParameterList(True, True)
          Dim vBatchNumber As Integer = IntegerValue(dgr.GetValue(dgr.CurrentRow, "BatchNumber"))
          Dim vTransactionNumber As Integer = IntegerValue(dgr.GetValue(dgr.CurrentRow, "TransactionNumber"))
          If mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactEventBookings Then
            vList = New ParameterList(True, False)
            vList("SmartClient") = "Y"
            vList.IntegerValue("BatchNumber") = vBatchNumber
            vList.IntegerValue("TransactionNumber") = vTransactionNumber
            vEventNumber = IntegerValue(dgr.GetValue(dgr.CurrentRow, "EventNumber"))
            Dim vRow As DataRow = DataHelper.GetRowFromDataSet(DataHelper.GetTransactionData(CareServices.XMLTransactionDataSelectionTypes.xtdtTransactionDetails, vBatchNumber, vTransactionNumber))
            If vRow IsNot Nothing Then
              Dim vContactNumber As Integer = IntegerValue(vRow.Item("ContactNumber").ToString)   'This is the payer Contact & could be different to the booker Contact
              Dim vPTRow As DataRow = DataHelper.GetRowFromDataSet(DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactProcessedTransactions, vContactNumber, vList))
              If vPTRow IsNot Nothing Then
                vTransDate = vPTRow.Item("TransactionDate").ToString
                vTransSign = vPTRow.Item("TransactionSign").ToString
                vStock = BooleanValue(vPTRow.Item("ContainsStock").ToString)
                If vStock = False Then vStock = BooleanValue(vPTRow.Item("ContainsPostage").ToString)
                If (BooleanValue(dgr.GetValue(dgr.CurrentRow, "CreditSale")) = True AndAlso BooleanValue(dgr.GetValue(dgr.CurrentRow, "InvoicePrinted")) = False) Then vEventAnalysis = True
                vList.AddSystemColumns()
              Else
                Throw New CareException(GetInformationMessage(InformationMessages.ImCannotFindFinancialHistoryDetails, vBatchNumber.ToString, vTransactionNumber.ToString), CareException.ErrorNumbers.enCannotFindFinancialHistoryDetails)
              End If
            Else
              Throw New CareException(GetInformationMessage(InformationMessages.ImCannotFindFinancialHistoryDetails, vBatchNumber.ToString, vTransactionNumber.ToString), CareException.ErrorNumbers.enCannotFindFinancialHistoryDetails)
            End If
          Else
            vTransDate = dgr.GetValue(dgr.CurrentRow, "TransactionDate")
            vTransSign = dgr.GetValue(dgr.CurrentRow, "TransactionSign")
            vStock = BooleanValue(dgr.GetValue(dgr.CurrentRow, "ContainsStock"))
            If vStock = False Then vStock = BooleanValue(dgr.GetValue(dgr.CurrentRow, "ContainsPostage"))
          End If
          vList.IntegerValue("BatchNumber") = vBatchNumber
          vList.IntegerValue("TransactionNumber") = vTransactionNumber
          Dim vContainsSLItems As Boolean = BooleanValue(dgr.GetValue(dgr.CurrentRow, "ContainsSalesLedgerItems"))
          vList("PaymentMethodCode") = dgr.GetValue(dgr.CurrentRow, "PaymentMethodCode")
          If dgr.MultipleRowsSelected Then
            Dim vSelectedTrans As ArrayListEx = dgr.GetSelectedRowIntegers("BatchNumber", "TransactionNumber")
            FormHelper.RunFinancialAdjustments(CareServices.AdjustmentTypes.atAdjustment, vList, vTransDate, vTransSign, vStock, vSelectedTrans)
          Else
            vList("ContainsSalesLedgerItems") = CBoolYN(vContainsSLItems)
            Select Case pItem
              Case BaseFinancialMenu.FinancialMenuItems.fmiReverse

                FormHelper.RunFinancialAdjustments(CareServices.AdjustmentTypes.atReverse, vList, vTransDate, vTransSign, vStock)
              Case BaseFinancialMenu.FinancialMenuItems.fmiRefund
                FormHelper.RunFinancialAdjustments(CareServices.AdjustmentTypes.atRefund, vList, vTransDate, vTransSign, vStock)
              Case BaseFinancialMenu.FinancialMenuItems.fmiMove
                FormHelper.RunFinancialAdjustments(CareServices.AdjustmentTypes.atMove, vList, vTransDate, vTransSign, vStock)
              Case BaseFinancialMenu.FinancialMenuItems.fmiAnalysis
                If vEventAnalysis Then
                  FormHelper.RunFinancialAdjustments(CareServices.AdjustmentTypes.atEventAdjustment, vList, vTransDate, vTransSign, vStock)
                Else
                  FormHelper.RunFinancialAdjustments(CareServices.AdjustmentTypes.atAdjustment, vList, vTransDate, vTransSign, vStock)
                End If
            End Select
            If vEventNumber > 0 Then
              MainHelper.RefreshEventData(CareServices.XMLEventDataSelectionTypes.xedtEventBookings, vEventNumber)
              MainHelper.RefreshData(CareServices.XMLContactDataSelectionTypes.xcdtContactEventBookings, mvContactInfo.ContactNumber)
            End If
          End If
          vRefresh = True

        Case FinancialMenu.FinancialMenuItems.fmiRefundInAdvance, FinancialMenu.FinancialMenuItems.fmiReverseInAdvance
          Dim vPPNumber As Integer = IntegerValue(dgr.GetValue(dgr.CurrentRow, "PaymentPlanNumber"))
          Dim vTable As DataTable = DataHelper.GetTableFromDataSet(DataHelper.GetPaymentPlanData(CareServices.XMLPaymentPlanDataSelectionTypes.xpdtPaymentPlanPayments, vPPNumber))
          If vTable IsNot Nothing Then
            vTable.DefaultView.RowFilter = "Status='I'"
            vTable = vTable.DefaultView.ToTable
            If vTable.Rows.Count > 0 Then
              Dim vRow As DataRow = vTable.Rows(0)
              Dim vList As New ParameterList(True, True)
              vList("BatchNumber") = vRow("BatchNumber").ToString
              vList("TransactionNumber") = vRow("TransactionNumber").ToString
              vList.IntegerValue("PaymentPlanNumber") = vPPNumber
              Dim vTransTable As DataTable = DataHelper.GetTableFromDataSet(DataHelper.GetTransactionData(CareServices.XMLTransactionDataSelectionTypes.xtdtTransactionDetails, vList.IntegerValue("BatchNumber"), vList.IntegerValue("TransactionNumber")))
              Dim vTransSign As String = ""
              If vTransTable IsNot Nothing AndAlso vTransTable.Rows.Count > 0 Then
                vTransSign = vTransTable.Rows(0)("TransactionSign").ToString
              End If
              FormHelper.RunFinancialAdjustments(CareServices.AdjustmentTypes.atNone, vList, "", vTransSign, False)
              vRefresh = True
            Else
              ShowInformationMessage(InformationMessages.ImNoPaymentHistoryForInAdvancePayments)
            End If
          End If

        Case FinancialMenu.FinancialMenuItems.fmiSkipPayment
          Dim vParamList As New ParameterList(True)
          vParamList.IntegerValue("PaymentPlanNumber") = SelectRowItemNumber(dgr.CurrentRow, "PaymentPlanNumber")
          Dim vReturnList As ParameterList = DataHelper.ProcessPaymentPlanMenu(CareServices.XMLPaymentPlanMenuTypes.xpmtSkipPayment, vParamList)
          If vReturnList.IntegerValue("PaymentPlanNumber") > 0 Then
            ShowInformationMessage(InformationMessages.ImPaymentPlanPaymentSkipped)
            vRefresh = True
          End If

        Case BaseFinancialMenu.FinancialMenuItems.fmiSubChangeCommunication
          Dim vParamList As New ParameterList()
          vParamList.IntegerValue("SubscriptionNumber") = SelectRowItemNumber(dgr.CurrentRow, "SubscriptionNumber")
          vParamList.IntegerValue("AddressNumber") = SelectRowItemNumber(dgr.CurrentRow, "AddressNumber")
          vParamList.IntegerValue("ContactNumber") = mvContactInfo.ContactNumber
          Dim vCommunicationNumber As Integer = SelectRowItemNumber(dgr.CurrentRow, "CommunicationNumber")
          If vCommunicationNumber > 0 Then vParamList.IntegerValue("CommunicationNumber") = vCommunicationNumber
          Dim vFrmAP As New frmApplicationParameters(CareServices.FunctionParameterTypes.fptChangeSubscriptionCommunication, vParamList, Nothing)
          If vFrmAP.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
            Dim vList As New ParameterList(True)
            If vFrmAP.ReturnList.ContainsKey("CommunicationNumber") Then
              vList("CommunicationNumber") = vFrmAP.ReturnList("CommunicationNumber")
            Else
              vList("CommunicationNumber") = ""
            End If
            vList("SubscriptionNumber") = vParamList("SubscriptionNumber")
            DataHelper.UpdateSubscription(vList)
            vRefresh = True
          End If

        Case BaseFinancialMenu.FinancialMenuItems.fmiConfirmTransaction
          Dim vBatch As New BatchInfo(IntegerValue(pDataRow("BatchNumber")))
          Dim vList As New ParameterList(True)
          vList("BatchNumber") = pDataRow("BatchNumber").ToString
          vList("TransactionNumber") = pDataRow("TransactionNumber").ToString
          Select Case vBatch.BatchType
            Case CareServices.BatchTypes.GiftInKind, CareServices.BatchTypes.SaleOrReturn, CareServices.BatchTypes.Cash
              If DataHelper.GetCount(CareNetServices.XMLGetCountTypes.xgctConfirmedTrans, vList) > 0 Then
                ShowErrorMessage(InformationMessages.ImTransactionAlreadyConfirmed)
              Else
                Dim vAppNumber As String = AppValues.ConfigurationValue(AppValues.ConfigurationValues.trader_application_conf)
                If vAppNumber.Length = 0 Then vAppNumber = AppValues.ConfigurationValue(AppValues.ConfigurationValues.trader_application_fa)
                Dim vTA As New TraderApplication(IntegerValue(vAppNumber), vList.IntegerValue("BatchNumber"))
                vTA.BatchInfo = vBatch
                vTA.BatchNumber = vList.IntegerValue("BatchNumber")
                vTA.TransactionNumber = vList.IntegerValue("TransactionNumber")
                Dim vAdjustmentType As BatchInfo.AdjustmentTypes
                If vBatch.BatchType = CareServices.BatchTypes.Cash Then vAdjustmentType = BatchInfo.AdjustmentTypes.CashBatchConfirmation Else vAdjustmentType = BatchInfo.AdjustmentTypes.GIKConfirmation
                FormHelper.RunTraderApplication(vTA, vAdjustmentType)
                vRefresh = True
              End If
            Case Else
              Dim vParams As ParameterList = FormHelper.ShowApplicationParameters(CareServices.FunctionParameterTypes.fptConfirmProvisionalTransaction)
              If vParams.Count > 0 Then
                If vParams.Contains("Date") Then vList("Date") = vParams("Date")
                vList("TransactionDate") = vParams("TransactionDate")
                Dim vResult As ParameterList = DataHelper.ConfirmProvisionalTransaction(vList)
                ShowInformationMessage(InformationMessages.ImTransactionConfirmed, vResult("Result"))
                vRefresh = True
              End If
          End Select

        Case BaseFinancialMenu.FinancialMenuItems.fmiReinstateProvisionalTrans
          DataHelper.CancelReinstateProvisionalTransaction(IntegerValue(pDataRow("BatchNumber")), IntegerValue(pDataRow("TransactionNumber")), False)
          vRefresh = True

        Case BaseFinancialMenu.FinancialMenuItems.fmiConfirmPaymentPlan
          Dim vList As New ParameterList(True)
          vList("PaymentPlanNumber") = pDataRow("PaymentPlanNumber").ToString
          DataHelper.ProcessPaymentPlanMenu(CareServices.XMLPaymentPlanMenuTypes.xpmtConfirmProvisionalPaymentPlan, vList)
          vRefresh = True
        Case BaseFinancialMenu.FinancialMenuItems.fmiUnlockFundraisingRequest
          Dim vList As New ParameterList(True)
          vList("FundraisingRequestNumber") = pDataRow("FundraisingRequestNumber").ToString
          vList("FundraisingStatus") = AppValues.ControlValue(AppValues.ControlValues.fundraising_status)
          vList("StatusChangeReason") = "Unlocked"
          DataHelper.UpdateItem(CareServices.XMLMaintenanceControlTypes.xmctFundraising, vList)
          vRefresh = True
        Case BaseFinancialMenu.FinancialMenuItems.fmiGoToActions
          If mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactFundraising Then
            If mvDataSet4 IsNot Nothing AndAlso mvDataSet4.Tables.Contains("DataRow") Then
              Dim vDataSet As DataSet = mvDataSet4.Copy
              vDataSet.Tables("DataRow").DefaultView.RowFilter = "ISNULL(ScheduledPaymentNumber,'') = '' AND ISNULL(CompletedOn,'') = ''"
              Dim vTable As DataTable = vDataSet.Tables("DataRow").DefaultView.ToTable
              If vTable.Rows.Count > 0 Then
                vDataSet.Tables.Remove("DataRow")
                vDataSet.Tables.Add(vTable)
                Dim vForm As New frmSelectListItem(vDataSet, frmSelectListItem.ListItemTypes.litFundraisingActions)
                If vForm.ShowDialog = System.Windows.Forms.DialogResult.OK Then
                  dts.DisplayGrid(4).SelectRow("ActionNumber", vForm.SelectedRow.ToString)
                  dts.SelectTab(5)
                End If
              End If
            End If
          End If
        Case BaseFinancialMenu.FinancialMenuItems.fmiNewAdHocAction
          If mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactFundraising Then HandleMenuClick(False, MenuSource.dgr)
        Case BaseFinancialMenu.FinancialMenuItems.fmiNewActionFromTemplate
          If mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactFundraising Then
            Dim vList As New ParameterList
            vList("FundraisingRequestNumber") = dgr.GetValue(dgr.CurrentDataRow, "FundraisingRequestNumber")
            vList("Logname") = dgr.GetValue(dgr.CurrentDataRow, "Logname")
            FormHelper.NewActionFromTemplate(Me, mvContactInfo.ContactNumber, vList)
          End If
        Case FinancialMenu.FinancialMenuItems.fmiAddGiftAidDeclaration
          Dim vList As New ParameterList
          Dim vReturnList As New ParameterList
          Dim vDR As DataRow = DataHelper.GetRowFromDataSet(DataHelper.GetGiftAidData(CareServices.XMLGiftAidDataSelectionTypes.xgdtGiftAidEarliestStartDate, 0))
          If mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactProcessedTransactions OrElse mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactUnProcessedTransactions OrElse
           mvDataType = CareNetServices.XMLContactDataSelectionTypes.xcdtContactDeliveryTransactions OrElse
           mvDataType = CareNetServices.XMLContactDataSelectionTypes.xcdtContactSalesTransactions Then
            vList("StartDate") = SelectRowItem(dgr.CurrentRow, "TransactionDate")
            vList("EndDate") = SelectRowItem(dgr.CurrentRow, "TransactionDate")
            vList("FromTransaction") = "Y"
          Else
            vList("StartDate") = SelectRowItem(dgr.CurrentRow, "StartDate")
            vList("EndDate") = AppValues.TodaysDate
          End If
          If CDate(vList("StartDate")) < CDate(vDR.Item("GiftAidEarliestStartDate").ToString) Then
            vList("StartDate") = vDR.Item("GiftAidEarliestStartDate").ToString
            If vList("EndDate").Length > 0 Then
              If CDate(vList("EndDate")) < CDate(vList("StartDate")) Then vList("EndDate") = vList("StartDate")
            End If
          End If
          If mvDataSet.Tables("DataRow").Columns.Contains("PaymentPlanNumber") Then
            vList("PaymentPlanNumber") = SelectRowItem(dgr.CurrentRow, "PaymentPlanNumber")
          End If
          If mvDataSet.Tables("DataRow").Columns.Contains("TransactionNumber") Then
            vList("TransactionNumber") = SelectRowItem(dgr.CurrentRow, "TransactionNumber")
          End If
          If mvDataSet.Tables("DataRow").Columns.Contains("BatchNumber") Then
            vList("BatchNumber") = SelectRowItem(dgr.CurrentRow, "BatchNumber")
          End If
          If mvDataSet.Tables("DataRow").Columns.Contains("SourceCode") Then
            vList("Source") = SelectRowItem(dgr.CurrentDataRow, "SourceCode")
          End If
          vList("Donations") = "Y"
          If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.ga_membership_tax_reclaim, False) Then
            vList("Members") = "Y"
          Else
            vList("Members") = "N"
          End If
          Dim vResultList As ParameterList = FormHelper.ShowApplicationParameters(EditPanelInfo.OtherPanelTypes.optGiftAidDeclaration, Nothing, vList, "Create Gift Aid Declaration")
          If vResultList IsNot Nothing AndAlso vResultList.Count > 0 Then
            vResultList("ContactNumber") = mvContactInfo.ContactNumber.ToString
            If vList.Contains("PaymentPlanNumber") Then vResultList("PaymentPlanNumber") = vList("PaymentPlanNumber")
            If vList.Contains("BatchNumber") Then vResultList("BatchNumber") = vList("BatchNumber")
            If vList.Contains("TransactionNumber") Then vResultList("TransactionNumber") = vList("TransactionNumber")

            vReturnList = DataHelper.AddItem(CareServices.XMLMaintenanceControlTypes.xmctGiftAidDeclarations, vResultList)
            If vReturnList.Contains("DeclarationNumber") AndAlso vReturnList("DeclarationNumber").Length > 0 Then
              ShowInformationMessage(InformationMessages.ImGiftAidDeclarationCreated, vReturnList("DeclarationNumber"))
              vRefresh = True
            ElseIf vReturnList.Contains("Message") AndAlso vReturnList("Message").Length > 0 Then
              ShowInformationMessage(vReturnList("Message"))
              vRefresh = False
            End If
          End If

        Case BaseFinancialMenu.FinancialMenuItems.fmiGoToTransaction
          mvTransactionLinkMenu_MenuSelected(pItem, pDataRow, pChangeDetails, pFinancialMenu)

        Case BaseFinancialMenu.FinancialMenuItems.fmiChangeClaimDate
          mvTransactionLinkMenu_MenuSelected(pItem, pDataRow, pChangeDetails, pFinancialMenu)

        Case BaseFinancialMenu.FinancialMenuItems.fmiRecalcLoanInterest
          Dim vDefaultList As New ParameterList
          vDefaultList("CalculationDate") = Today.ToShortDateString()
          Dim vDate As Date = CDate(pDataRow("InterestCalculatedDate").ToString).AddDays(1)
          vDefaultList("MinDate") = vDate.ToShortDateString
          If vDate > Today Then vDefaultList("CalculationDate") = vDate.ToShortDateString()
          vDate = CDate(pDataRow("ExpiryDate").ToString)
          If vDate > Today.AddYears(1) Then vDate = Today.AddYears(1)
          vDefaultList("MaxDate") = vDate.ToShortDateString
          Dim vParams As ParameterList = FormHelper.ShowApplicationParameters(CareNetServices.FunctionParameterTypes.fptReCalculateLoanInterest, vDefaultList, Nothing, Me)
          If vParams IsNot Nothing AndAlso vParams.Count > 0 Then
            Dim vCursor As New BusyCursor()
            vParams.AddSystemColumns()
            Dim vResult As ParameterList = DataHelper.RecalculateLoanInterest(IntegerValue(pDataRow("LoanNumber").ToString), vParams)
            If vResult IsNot Nothing AndAlso vResult.ContainsKey("Result") Then
              ShowInformationMessage(InformationMessages.ImLoanInterestCalculated)
              vRefresh = True
            End If
            vCursor.Dispose()
          End If

        Case BaseFinancialMenu.FinancialMenuItems.fmiDisplayTransactionsAll, BaseFinancialMenu.FinancialMenuItems.fmiDisplayTransactionsFullyAllocatedOnly, BaseFinancialMenu.FinancialMenuItems.fmiDisplayTransactionsUnallocatedOnly
          Select Case pItem
            Case BaseFinancialMenu.FinancialMenuItems.fmiDisplayTransactionsFullyAllocatedOnly
              mvFinancialMenu.DisplayTransactionsAllocationType = "F"
            Case BaseFinancialMenu.FinancialMenuItems.fmiDisplayTransactionsUnallocatedOnly
              mvFinancialMenu.DisplayTransactionsAllocationType = "U"
            Case Else
              mvFinancialMenu.DisplayTransactionsAllocationType = "A"
          End Select
          RefreshData(CDBNETCL.CareNetServices.XMLMaintenanceControlTypes.xmctNone)

        Case BaseFinancialMenu.FinancialMenuItems.fmiPreviewInvoice, BaseFinancialMenu.FinancialMenuItems.fmiPrintReceipt
          RunMailmerge(pItem, pDataRow)

      End Select

      If vRefresh Then RefreshData()
    Catch vCareException As CareException
      Select Case vCareException.ErrorNumber
        Case CareException.ErrorNumbers.enPaymenPlanReinstatementError,
         CareException.ErrorNumbers.enFMTValidationError,
         CareException.ErrorNumbers.enTraderUnsupportedFeature,
         CareException.ErrorNumbers.enBookingBatchNotPosted,
         CareException.ErrorNumbers.enBookingAlreadyCancelled,
         CareException.ErrorNumbers.enDelegatesAlreadyAttended,
         CareException.ErrorNumbers.enRecordAlreadyProcessed,
         CareException.ErrorNumbers.enCannotAdjustPaymentStatus,
         CareException.ErrorNumbers.enCannotReverseOrderAnalysis,
         CareException.ErrorNumbers.enCannotAdjustZeroBalancePP,
         CareException.ErrorNumbers.enAdjustmentError,
         CareException.ErrorNumbers.enOriginalBatchOrTransPurged,
         CareException.ErrorNumbers.enOriginalPaymentPartProcessed,
         CareException.ErrorNumbers.enCannotFindFinancialHistoryDetails,
         CareException.ErrorNumbers.enAllocationsBatchUnposted,
         CareException.ErrorNumbers.enInvoiceAllocationsUnposted,
         CareException.ErrorNumbers.enCannotChangeCentrePriceMismatch,
         CareException.ErrorNumbers.enScheduleClashInBooking,
         CareException.ErrorNumbers.enScheduleClashExistingBooking,
         CareException.ErrorNumbers.enExamBooking,
         CareException.ErrorNumbers.enDirectDebitReferenceNotUnique,
         CareException.ErrorNumbers.enCreditNoteAllocNoNumber,
         CareException.ErrorNumbers.enCannotRemoveSundryCreditNoteReversal,
         CareException.ErrorNumbers.enPaymentPlanUpdated
          ShowInformationMessage(vCareException.Message)
        Case CareException.ErrorNumbers.enInvoiceAllocationError, CareException.ErrorNumbers.enAllocateOrUnallocateCreditNote, CareException.ErrorNumbers.enUnallocateCreditNote
          Dim vContinue As Boolean = False
          Dim vUnallocateCreditNote As Boolean = False
          If vCareException.ErrorNumber = CareException.ErrorNumbers.enInvoiceAllocationError OrElse vCareException.ErrorNumber = CareException.ErrorNumbers.enUnallocateCreditNote Then
            If ShowQuestion(vCareException.Message, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
              vContinue = True
              If vCareException.ErrorNumber = CareException.ErrorNumbers.enUnallocateCreditNote Then vUnallocateCreditNote = True
            End If
          ElseIf vCareException.ErrorNumber = CareException.ErrorNumbers.enAllocateOrUnallocateCreditNote Then
            Dim vDialogueResult As System.Windows.Forms.DialogResult = ShowQuestion(vCareException.Message, MessageBoxButtons.YesNoCancel)
            If vDialogueResult <> System.Windows.Forms.DialogResult.Cancel Then
              vContinue = True
              vUnallocateCreditNote = vDialogueResult = System.Windows.Forms.DialogResult.No
            End If
          End If
          If vContinue Then
            Dim vExamBooking As Boolean = vInvoiceParams.ContainsKey("ExamBookingId")
            If vInvoiceParams.ToString.Length > 0 Then
              vInvoiceParams.Add("AllocationsChecked", "Y")
              If vUnallocateCreditNote Then vInvoiceParams.Add("UnallocateCreditNote", "Y")
            End If
            Dim vReturnList As ParameterList = Nothing
            Try
              If vExamBooking Then
                vReturnList = ExamsDataHelper.CancelExamBooking(vInvoiceParams)
              Else
                vReturnList = DataHelper.CancelEventBooking(vInvoiceParams)
              End If
            Catch vCareEx As CareException
              If vCareEx.ErrorNumber = CareException.ErrorNumbers.enCCAuthorisationFailed OrElse vCareEx.ErrorNumber = CareException.ErrorNumbers.enCardAuthorisationUnexpectedTimeout Then
                ShowInformationMessage(vCareEx.Message)
              Else
                Throw vCareEx
              End If
            End Try
            If vReturnList.Contains("Message") Then
              ShowInformationMessage(vReturnList("Message"))
            ElseIf vReturnList.Contains("ProcessWaitingList") Then
              Dim vEventInfo As New CareEventInfo(DataHelper.GetEventData(CareServices.XMLEventDataSelectionTypes.xedtEventInformation, SelectRowItemNumber(dgr.CurrentRow, "EventNumber")).Tables("DataRow").Rows(0))
              Dim vWaiting As New frmWaitingList(vEventInfo)
              vWaiting.ShowDialog(Me)
            End If
            If vExamBooking Then
              ShowInformationMessage(InformationMessages.ImExamBookingCancelled)
            Else
              ShowInformationMessage(InformationMessages.ImEventBookingCancelled)
            End If
            RefreshCard()
          End If
        Case CareException.ErrorNumbers.enGADDatesOverlap
          ShowInformationMessage(vCareException.Message)
        Case Else
          DataHelper.HandleException(vCareException)
      End Select
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub GoToPaymentPlanType(ByVal pType As CareServices.XMLContactDataSelectionTypes, ByVal pDataRow As DataRow, ByRef pRefresh As Boolean)
    Dim vPaymentPlanNumber As Integer = IntegerValue(pDataRow("PaymentPlanNumber").ToString())

    Dim vTable As DataTable = DataHelper.GetTableFromDataSet(DataHelper.GetPaymentPlanData(CareServices.XMLPaymentPlanDataSelectionTypes.xpdtPaymentPlanContactInformation, vPaymentPlanNumber))
    Dim vRowName As String
    Dim vCancel As Boolean = False
    Select Case pType
      Case CareServices.XMLContactDataSelectionTypes.xcdtContactCreditCardAuthorities
        vRowName = "CreditCardAuthority"
      Case CareServices.XMLContactDataSelectionTypes.xcdtContactDirectDebits
        vRowName = "DirectDebit"
      Case CareServices.XMLContactDataSelectionTypes.xcdtContactPaymentPlans
        vRowName = "PaymentPlan"
      Case CareServices.XMLContactDataSelectionTypes.xcdtContactStandingOrders
        vRowName = "StandingOrder"
      Case CareServices.XMLContactDataSelectionTypes.xcdtContactCovenants
        vRowName = "Covenant"
      Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactLoans
        vRowName = "Loan"
      Case Else
        vRowName = "Member"
    End Select
    If vTable IsNot Nothing Then
      Dim vContactNumber As Integer
      Dim vMemberCount As Integer
      For Each vContactRow As DataRow In vTable.Rows
        If vContactRow("ContactType").ToString = "Member" AndAlso vContactRow("CancellationReason").ToString.Length = 0 Then
          vMemberCount += 1
        End If
        If vContactRow("ContactType").ToString = vRowName Then
          Dim vRowContact As Integer = IntegerValue(vContactRow("ContactNumber").ToString)
          If vContactNumber = 0 OrElse vRowContact = vContactNumber Then
            vContactNumber = vRowContact
          End If
        End If
      Next
      If vContactNumber = 0 Then
        ShowInformationMessage(InformationMessages.ImPaymentPlanInvalid)
        pRefresh = False
        Exit Sub
      ElseIf pType = CareServices.XMLContactDataSelectionTypes.xcdtContactMembershipDetails AndAlso vMemberCount > 1 AndAlso vContactNumber = mvContactInfo.ContactNumber Then
        ShowInformationMessage(InformationMessages.ImMultipleMembersIncluding)
      ElseIf pType = CareServices.XMLContactDataSelectionTypes.xcdtContactMembershipDetails AndAlso vMemberCount > 1 AndAlso vContactNumber <> mvContactInfo.ContactNumber Then
        ShowInformationMessage(InformationMessages.ImMultipleMembersExcluding)
        pRefresh = False
        Exit Sub
      ElseIf vContactNumber <> mvContactInfo.ContactNumber Then
        Dim vContactInfo As New ContactInfo(vContactNumber)
        If ShowQuestion(InformationMessages.ImDifferentContact, MessageBoxButtons.YesNo, vContactInfo.ContactName) = System.Windows.Forms.DialogResult.Yes Then
          Dim vForm As Form = FormHelper.ShowCardIndex(pType, vContactNumber, False)
          If vForm IsNot Nothing Then
            DirectCast(vForm, frmCardSet).SelectRowItem("PaymentPlanNumber", vPaymentPlanNumber)
            pRefresh = False
          End If
          Exit Sub
        Else
          'Jira 673: Cancel navigation to Payment Plan node if user selects No in question box
          vCancel = True
        End If
      End If
    End If
    If Not vCancel Then
      sel.SetSelectionType(pType)
      SelectRowItem("PaymentPlanNumber", vPaymentPlanNumber)
    End If
    pRefresh = False
  End Sub

  Private Sub GoToBankAccount(ByVal pBankDetailsNumber As Integer, ByRef pRefresh As Boolean)
    Dim vCancel As Boolean = False
    Dim vList As New ParameterList(True)
    vList("BankDetailsNumber") = pBankDetailsNumber.ToString
    Dim vTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtContactAccounts, vList)
    If vTable IsNot Nothing Then
      Dim vContactNumber As Integer = IntegerValue(vTable.Rows(0).Item("ContactNumber").ToString)
      If vContactNumber <> mvContactInfo.ContactNumber Then
        Dim vContactInfo As New ContactInfo(vContactNumber)
        If ShowQuestion(InformationMessages.ImDifferentContact, MessageBoxButtons.YesNo, vContactInfo.ContactName) = System.Windows.Forms.DialogResult.Yes Then
          Dim vForm As Form = FormHelper.ShowCardIndex(CareNetServices.XMLContactDataSelectionTypes.xcdtContactBankAccounts, vContactNumber, False)
          If vForm IsNot Nothing Then
            DirectCast(vForm, frmCardSet).SelectRowItem("BankDetailsNumber", pBankDetailsNumber)
            pRefresh = False
          End If
          Exit Sub
        Else
          'Cancel navigation to Bank Account node if user selects No in question box
          vCancel = True
        End If
      End If
    End If
    If Not vCancel Then
      sel.SetSelectionType(CareNetServices.XMLContactDataSelectionTypes.xcdtContactBankAccounts)
      SelectRowItem("BankDetailsNumber", pBankDetailsNumber)
    End If
    pRefresh = False
  End Sub

  Public Sub SelectRowItem(ByVal pName As String, ByVal pNumber As Integer)
    Dim vRow As Integer = dgr.FindRow(pName, pNumber.ToString)
    If vRow >= 0 Then dgr.SelectRow(vRow)
  End Sub

  Private Function SelectTransaction(ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer) As Boolean
    Return SelectTransaction(pBatchNumber, pTransactionNumber, 0)
  End Function

  Public Function SelectTransaction(ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer, ByVal pLineNumber As Integer) As Boolean
    Dim vBatchCol As Integer = dgr.GetColumn("BatchNumber")
    Dim vTransCol As Integer = dgr.GetColumn("TransactionNumber")
    Dim vLineCol As Integer
    If pLineNumber > 0 Then vLineCol = dgr.GetColumn("LineNumber")
    If vBatchCol > -1 AndAlso vTransCol > -1 Then
      For vRow As Integer = 0 To dgr.DataRowCount - 1
        If IntegerValue(dgr.GetValue(vRow, vBatchCol)) = pBatchNumber AndAlso IntegerValue(dgr.GetValue(vRow, vTransCol)) = pTransactionNumber Then
          If pLineNumber > 0 AndAlso vLineCol >= 0 Then
            If IntegerValue(dgr.GetValue(vRow, vLineCol)) = pLineNumber Then
              dgr.SelectRow(vRow)
              Return True
            End If
          Else
            dgr.SelectRow(vRow)
            If mvDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactProcessedTransactions Then
              If pLineNumber > 0 Then
                Dim vLineRow As Integer = dts.DisplayGrid(0).FindRow("LineNumber", pLineNumber.ToString)
                If vLineRow > 0 Then dts.DisplayGrid(0).SelectRow(vLineRow)
              End If
            End If
            Return True
          End If
        End If
      Next
    End If
  End Function

  Public Function SelectPurchaseOrder(ByVal pPONumber As Integer, ByVal pPaymentNumber As Integer) As Boolean

    Dim vPOCol As Integer = dts.DisplayGrid(1).GetColumn("PurchaseOrderNumber")
    Dim vPNCol As Integer = dts.DisplayGrid(1).GetColumn("PaymentNumber")

    If vPOCol > -1 AndAlso vPNCol > -1 Then
      For vRow As Integer = 0 To dts.DisplayGrid(1).DataRowCount - 1
        If IntegerValue(dts.DisplayGrid(1).GetValue(vRow, vPOCol)) = pPONumber AndAlso IntegerValue(dts.DisplayGrid(1).GetValue(vRow, vPNCol)) = pPaymentNumber Then
          dts.DisplayGrid(1).SelectRow(vRow)
          Return True
        End If
      Next
    End If
  End Function
  Public Function SelectPurchaseInvoice(ByVal pPurchaseInvoiceNumber As Integer) As Boolean
    Dim vPOCol As Integer = dgr.GetColumn("PurchaseInvoiceNumber")

    If vPOCol > -1 Then
      For vRow As Integer = 0 To dgr.DataRowCount - 1
        If IntegerValue(dgr.GetValue(vRow, vPOCol)) = pPurchaseInvoiceNumber Then
          dgr.SelectRow(vRow)
          Return True
        End If
      Next
    End If
  End Function


  Private Sub mvAnalysisFinancialMenu_MenuSelected(ByVal pItem As FinancialMenu.FinancialMenuItems, ByVal pDataRow As DataRow) Handles mvAnalysisFinancialMenu.MenuSelected
    Try
      Dim vRefresh As Boolean = False
      Select Case pItem
        Case FinancialMenu.FinancialMenuItems.fmiGoToBackOrders
          sel.SetSelectionType(CareServices.XMLContactDataSelectionTypes.xcdtContactBackOrders)
          SelectTransaction(mvAnalysisFinancialMenu.TargetBatchNumber, mvAnalysisFinancialMenu.TargetTransactionNumber, mvAnalysisFinancialMenu.TargetLineNumber)
        Case FinancialMenu.FinancialMenuItems.fmiGoToChangedBy
          SelectTransaction(mvAnalysisFinancialMenu.AdjustmentBatchNumber, mvAnalysisFinancialMenu.AdjustmentTransactionNumber)
        Case FinancialMenu.FinancialMenuItems.fmiGoToChanges
          SelectTransaction(mvAnalysisFinancialMenu.AdjustmentWasBatchNumber, mvAnalysisFinancialMenu.AdjustmentWasTransactionNumber)
        Case FinancialMenu.FinancialMenuItems.fmiGoToCovenant
          GoToPaymentPlanType(CareServices.XMLContactDataSelectionTypes.xcdtContactCovenants, pDataRow, vRefresh)
        Case FinancialMenu.FinancialMenuItems.fmiGoToDespatch
          Dim vList As New ParameterList(True, False)
          vList.IntegerValue("BatchNumber") = mvAnalysisFinancialMenu.TargetBatchNumber
          vList.IntegerValue("TransactionNumber") = mvAnalysisFinancialMenu.TargetTransactionNumber
          vList.IntegerValue("LineNumber") = mvAnalysisFinancialMenu.TargetLineNumber
          vList("SystemColumns") = "N"
          Dim vRow As DataRow = DataHelper.GetRowFromDataSet(DataHelper.GetFinancialProcessingData(CareNetServices.XMLFinancialProcessingDataSelectionTypes.xbdstDespatchData, vList))
          If vRow IsNot Nothing Then
            SelectContactShowTransaction(IntegerValue(vRow("ContactNumber").ToString), mvAnalysisFinancialMenu.TargetBatchNumber, mvAnalysisFinancialMenu.TargetTransactionNumber, mvAnalysisFinancialMenu.TargetLineNumber, CareNetServices.XMLContactDataSelectionTypes.xcdtContactDespatchNotes)
          End If
        Case FinancialMenu.FinancialMenuItems.fmiGoToEvent, FinancialMenu.FinancialMenuItems.fmiGoToEvent2,
             FinancialMenu.FinancialMenuItems.fmiGoToEvent3,
             FinancialMenu.FinancialMenuItems.fmiGoToEvent4, FinancialMenu.FinancialMenuItems.fmiGoToEvent5
          FormHelper.ShowEventIndex(mvAnalysisFinancialMenu.EventNumber(pItem - FinancialMenu.FinancialMenuItems.fmiGoToEvent))

        Case FinancialMenu.FinancialMenuItems.fmiAddEventFinancialLink, FinancialMenu.FinancialMenuItems.fmiAddEventFinancialLink2,
             FinancialMenu.FinancialMenuItems.fmiAddEventFinancialLink3,
             FinancialMenu.FinancialMenuItems.fmiAddEventFinancialLink4, FinancialMenu.FinancialMenuItems.fmiAddEventFinancialLink5
          Dim vList As New ParameterList(True)
          vList("EventGroup") = DataHelper.EventGroups(pItem - FinancialMenu.FinancialMenuItems.fmiAddEventFinancialLink).Code
          Dim vEventNumber As Integer = FormHelper.ShowFinder(CareServices.XMLDataFinderTypes.xdftEvents, vList, Me)
          If vEventNumber > 0 Then DataHelper.AddEventFinancialLink(vEventNumber, mvAnalysisFinancialMenu.TargetBatchNumber, mvAnalysisFinancialMenu.TargetTransactionNumber, mvAnalysisFinancialMenu.TargetLineNumber)

        Case FinancialMenu.FinancialMenuItems.fmiRemoveEventFinancialLink, FinancialMenu.FinancialMenuItems.fmiRemoveEventFinancialLink2,
             FinancialMenu.FinancialMenuItems.fmiRemoveEventFinancialLink3,
             FinancialMenu.FinancialMenuItems.fmiRemoveEventFinancialLink4, FinancialMenu.FinancialMenuItems.fmiRemoveEventFinancialLink5
          DataHelper.DeleteEventFinancialLink(mvAnalysisFinancialMenu.EventNumber(pItem - FinancialMenu.FinancialMenuItems.fmiRemoveEventFinancialLink), mvAnalysisFinancialMenu.TargetBatchNumber, mvAnalysisFinancialMenu.TargetTransactionNumber, mvAnalysisFinancialMenu.TargetLineNumber)

        Case FinancialMenu.FinancialMenuItems.fmiGoToLinks
          Select Case mvAnalysisFinancialMenu.LineType
            Case "S"
              sel.SetSelectionType(CareServices.XMLContactDataSelectionTypes.xcdtContactXactionSentOnBehalfOf)
              SelectTransaction(mvAnalysisFinancialMenu.TargetBatchNumber, mvAnalysisFinancialMenu.TargetTransactionNumber, mvAnalysisFinancialMenu.TargetLineNumber)
            Case "H"
              sel.SetSelectionType(CareServices.XMLContactDataSelectionTypes.xcdtContactXactionPaidInBy)
              SelectTransaction(mvAnalysisFinancialMenu.TargetBatchNumber, mvAnalysisFinancialMenu.TargetTransactionNumber, mvAnalysisFinancialMenu.TargetLineNumber)
            Case "G"
              sel.SetSelectionType(CareServices.XMLContactDataSelectionTypes.xcdtContactXactionInMemoriamDonated)
              SelectTransaction(mvAnalysisFinancialMenu.TargetBatchNumber, mvAnalysisFinancialMenu.TargetTransactionNumber, mvAnalysisFinancialMenu.TargetLineNumber)
          End Select
        Case FinancialMenu.FinancialMenuItems.fmiGoToMembership
          GoToPaymentPlanType(CareServices.XMLContactDataSelectionTypes.xcdtContactMembershipDetails, pDataRow, vRefresh)
        Case FinancialMenu.FinancialMenuItems.fmiGoToPayPlan
          GoToPaymentPlanType(CareServices.XMLContactDataSelectionTypes.xcdtContactPaymentPlans, pDataRow, vRefresh)
        Case FinancialMenu.FinancialMenuItems.fmiGoToPreTaxPledge
          sel.SetSelectionType(CareServices.XMLContactDataSelectionTypes.xcdtContactPledges)
        Case FinancialMenu.FinancialMenuItems.fmiGoToPostTaxPledge
          sel.SetSelectionType(CareServices.XMLContactDataSelectionTypes.xcdtContactPostPGPledges)
        Case FinancialMenu.FinancialMenuItems.fmiRefund, FinancialMenu.FinancialMenuItems.fmiReverse
          Dim vTransDate As String = dgr.GetValue(dgr.CurrentRow, "TransactionDate")
          Dim vTransSign As String = dgr.GetValue(dgr.CurrentRow, "TransactionSign")
          Dim vStock As Boolean = BooleanValue(dts.DisplayGrid(0).GetValue(dts.DisplayGrid(0).CurrentRow, "StockProduct"))
          Dim vList As New ParameterList(True, True)
          vList("BatchNumber") = dgr.GetValue(dgr.CurrentRow, "BatchNumber")
          vList("TransactionNumber") = dgr.GetValue(dgr.CurrentRow, "TransactionNumber")
          vList("LineNumber") = dts.DisplayGrid(0).GetValue(dts.DisplayGrid(0).CurrentRow, "LineNumber")
          vList("Product") = dts.DisplayGrid(0).GetValue(dts.DisplayGrid(0).CurrentRow, "Product")
          vList("Rate") = dts.DisplayGrid(0).GetValue(dts.DisplayGrid(0).CurrentRow, "Rate")
          vList("Warehouse") = dts.DisplayGrid(0).GetValue(dts.DisplayGrid(0).CurrentRow, "Warehouse")
          vList("Quantity") = dts.DisplayGrid(0).GetValue(dts.DisplayGrid(0).CurrentRow, "Quantity")
          vList("Issued") = dts.DisplayGrid(0).GetValue(dts.DisplayGrid(0).CurrentRow, "Issued")
          vList("StockItem") = CBoolYN(vStock)
          vList("ContainsSalesLedgerItems") = CBoolYN(BooleanValue(pDataRow.Item("InvoicePayment").ToString))
          vList("CanPartRefund") = CBoolYN(mvAnalysisFinancialMenu.CanPartRefund)
          vList("LineTotal") = pDataRow.Item("Amount").ToString
          If pItem = FinancialMenu.FinancialMenuItems.fmiReverse Then
            FormHelper.RunFinancialAdjustments(CareServices.AdjustmentTypes.atReverse, vList, vTransDate, vTransSign, vStock)
          Else
            FormHelper.RunFinancialAdjustments(CareServices.AdjustmentTypes.atRefund, vList, vTransDate, vTransSign, vStock)
          End If
          vRefresh = True
        Case BaseFinancialMenu.FinancialMenuItems.fmiAddFundraisingPaymentLink
          Dim vList As New ParameterList(True, True)
          vList("BatchNumber") = dgr.GetValue(dgr.CurrentRow, "BatchNumber")
          vList.IntegerValue("ContactNumber") = mvContactInfo.ContactNumber
          Dim vScheduledNumber As Integer = FormHelper.ShowFinder(CareServices.XMLDataFinderTypes.xdftFundraisingPaymentScheduleFinder, vList, Me)
          If vScheduledNumber > 0 Then
            vList.IntegerValue("ScheduledPaymentNumber") = vScheduledNumber
            vList("TransactionNumber") = dgr.GetValue(dgr.CurrentRow, "TransactionNumber")
            vList("LineNumber") = dts.DisplayGrid(0).GetValue(dts.DisplayGrid(0).CurrentRow, "LineNumber")
            vList("Amount") = dts.DisplayGrid(0).GetValue(dts.DisplayGrid(0).CurrentRow, "Amount")
            vList("TransactionDate") = dgr.GetValue(dgr.CurrentRow, "TransactionDate")
            vList("Source") = dts.DisplayGrid(0).GetValue(dts.DisplayGrid(0).CurrentRow, "Source")
            DataHelper.AddFundraisingPaymentLink(vList)
            ShowInformationMessage(InformationMessages.ImFundraisingPaymentLinkCreated, vList("ScheduledPaymentNumber"))
          End If
      End Select
      If vRefresh Then RefreshCard()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub mvTransactionLinkMenu_MenuSelected(ByVal pItem As FinancialMenu.FinancialMenuItems, ByVal pDataRow As DataRow, ByVal pChangeDetails As Boolean, ByVal pFinancialMenu As BaseFinancialMenu) Handles mvTransactionLinkMenu.MenuSelected
    Try
      Select Case pItem
        Case BaseFinancialMenu.FinancialMenuItems.fmiChangeClaimDate
          Dim vDataSet As DataSet = DataHelper.GetPotentialClaimDatesForPayment(IntegerValue(pDataRow("ScheduledPaymentNumber")))
          Dim vForm As New frmSelectListItem(vDataSet, frmSelectListItem.ListItemTypes.litClaimDates)
          If vForm.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            DataHelper.UpdateClaimDateForPayment(IntegerValue(pDataRow("ScheduledPaymentNumber")), CStr(vDataSet.Tables("DataRow").Rows(vForm.SelectedRow)("ClaimDate")))
            pDataRow("ClaimDate") = CStr(vDataSet.Tables("DataRow").Rows(vForm.SelectedRow)("ClaimDate"))
          End If
        Case FinancialMenu.FinancialMenuItems.fmiGoToTransaction
          Select Case mvDataType
            Case CareServices.XMLContactDataSelectionTypes.xcdtContactXactionInMemoriamReceived,
                 CareServices.XMLContactDataSelectionTypes.xcdtContactXactionHandledBy,
                 CareServices.XMLContactDataSelectionTypes.xcdtContactXactionContributedTo,
                  CareNetServices.XMLContactDataSelectionTypes.xcdtContactLegacyBequests,
                  CareNetServices.XMLContactDataSelectionTypes.xcdtContactFundraising,
                  CareNetServices.XMLContactDataSelectionTypes.xcdtContactCreditCustomers,
                  CareNetServices.XMLContactDataSelectionTypes.xcdtContactExamDetails,
                  CareNetServices.XMLContactDataSelectionTypes.xcdtContactPaymentPlans,
                  CareNetServices.XMLContactDataSelectionTypes.xcdtContactMembershipDetails
              Dim vContactNumber As Integer
              If pDataRow.Table.Columns.Contains("ContactNumber") Then
                vContactNumber = IntegerValue(pDataRow("ContactNumber"))
              ElseIf pDataRow.Table.Columns.Contains("PayerContactNumber") Then
                vContactNumber = IntegerValue(pDataRow("PayerContactNumber"))
              End If
              SelectContactShowTransaction(vContactNumber, pFinancialMenu.TargetBatchNumber, pFinancialMenu.TargetTransactionNumber, pFinancialMenu.TargetLineNumber)
            Case Else
              SelectContactShowTransaction(0, pFinancialMenu.TargetBatchNumber, pFinancialMenu.TargetTransactionNumber, pFinancialMenu.TargetLineNumber)
              'ShowFinancialTransaction(pFinancialMenu.TargetBatchNumber, pFinancialMenu.TargetTransactionNumber, pFinancialMenu.TargetLineNumber)
          End Select
      End Select
    Catch vException As CareException
      Select Case vException.ErrorNumber
        Case CareException.ErrorNumbers.enClaimDateCannotChangeNoDD,
          CareException.ErrorNumbers.enClaimDateCannotChangeNotReconciled,
          CareException.ErrorNumbers.enClaimDateCannotChangeStatus
          ShowInformationMessage(vException.Message)
      End Select
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub SelectContactShowTransaction(ByVal pContactNumber As Integer, ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer, ByVal pLineNumber As Integer, Optional pType As CareServices.XMLContactDataSelectionTypes = CareNetServices.XMLContactDataSelectionTypes.xcdtContactProcessedTransactions)
    If pContactNumber = 0 Then
      Dim vList As New ParameterList(True)
      vList.IntegerValue("BatchNumber") = pBatchNumber
      vList.IntegerValue("TransactionNumber") = pTransactionNumber
      Dim vDataRow As DataRow = DataHelper.GetRowFromDataSet(DataHelper.GetTransactionData(CareNetServices.XMLTransactionDataSelectionTypes.xtdtTransactionDetails, vList))
      If vDataRow IsNot Nothing Then
        pContactNumber = IntegerValue(vDataRow("ContactNumber").ToString)
      End If
    End If

    If pContactNumber > 0 AndAlso pContactNumber <> mvContactInfo.ContactNumber Then
      Dim vContactInfo As New ContactInfo(pContactNumber)
      If ShowQuestion(InformationMessages.ImDifferentContact, MessageBoxButtons.YesNo, vContactInfo.ContactName) = System.Windows.Forms.DialogResult.Yes Then
        Dim vForm As Form = FormHelper.ShowCardIndex(pType, pContactNumber, False)
        If vForm IsNot Nothing Then
          DirectCast(vForm, frmCardSet).ShowFinancialTransaction(pType, pBatchNumber, pTransactionNumber, pLineNumber)
        End If
      End If
    Else
      ShowFinancialTransaction(pType, pBatchNumber, pTransactionNumber, pLineNumber)
    End If
  End Sub

  Public Sub ShowFinancialTransaction(pType As CareServices.XMLContactDataSelectionTypes, ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer, ByVal pLineNumber As Integer)
    'Look in processed transactions first
    sel.SetSelectionType(pType)
    If SelectTransaction(pBatchNumber, pTransactionNumber, pLineNumber) = False AndAlso pType = CareNetServices.XMLContactDataSelectionTypes.xcdtContactProcessedTransactions Then
      'If not found look in unprocessed transactions
      sel.SetSelectionType(CareServices.XMLContactDataSelectionTypes.xcdtContactUnProcessedTransactions)
      SelectTransaction(pBatchNumber, pTransactionNumber, pLineNumber)
    End If
  End Sub


  Private Sub mvJournalLinkMenu_MenuSelected(ByVal pItem As FinancialMenu.FinancialMenuItems, ByVal pDataRow As DataRow, ByVal pChangeDetails As Boolean, ByVal pFinancialMenu As BaseFinancialMenu) Handles mvJournalLinkMenu.MenuSelected
    Try
      Dim vNumber As Integer = IntegerValue(pDataRow("Select1"))
      Dim vType As CareServices.XMLContactDataSelectionTypes
      Dim vSelectionType As String = ""
      Select Case pItem
        Case FinancialMenu.FinancialMenuItems.fmiGoToCC
          vType = CareServices.XMLContactDataSelectionTypes.xcdtContactCreditCardAuthorities
          vSelectionType = "CreditCardAuthorityNumber"
        Case FinancialMenu.FinancialMenuItems.fmiGoToCovenant
          vType = CareServices.XMLContactDataSelectionTypes.xcdtContactCovenants
          vSelectionType = "CovenantNumber"
        Case FinancialMenu.FinancialMenuItems.fmiGoToDD
          vType = CareServices.XMLContactDataSelectionTypes.xcdtContactDirectDebits
          vSelectionType = "DirectDebitNumber"
        Case FinancialMenu.FinancialMenuItems.fmiGoToMembership
          vType = CareServices.XMLContactDataSelectionTypes.xcdtContactMembershipDetails
          vSelectionType = "MembershipNumber"
        Case FinancialMenu.FinancialMenuItems.fmiGoToPayPlan
          vType = CareServices.XMLContactDataSelectionTypes.xcdtContactPaymentPlans
          vSelectionType = "PaymentPlanNumber"
        Case FinancialMenu.FinancialMenuItems.fmiGoToSO
          vType = CareServices.XMLContactDataSelectionTypes.xcdtContactStandingOrders
          vSelectionType = "StandingOrderNumber"
      End Select
      If vSelectionType.Length > 0 Then
        If mvContactInfo.ContactNumber = 0 OrElse TypeOf (Me) Is frmCardDisplay Then
          Dim vContactNumber As Integer
          If mvContactInfo.ContactNumber = 0 Then
            vContactNumber = IntegerValue(pDataRow("ContactNumber"))
          Else
            vContactNumber = mvContactInfo.ContactNumber
          End If
          Dim vForm As Form = FormHelper.ShowCardIndex(vType, vContactNumber, False)
          If vForm IsNot Nothing Then
            DirectCast(vForm, frmCardSet).SelectRowItem(vSelectionType, vNumber)
          End If
        Else
          sel.SetSelectionType(vType)
          SelectRowItem(vSelectionType, vNumber)
        End If
      End If

    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
  Private Sub mvServiceBookingMenu_MenuSelected(ByVal pItem As ServiceBookingMenu.ServiceBookingMenuItems, ByVal pDataRow As DataRow, ByVal pChangeDetails As Boolean) Handles mvServiceBookingMenu.MenuSelected
    Try
      Dim vDefaults As New ParameterList
      Dim vRenameList As New ParameterList
      Dim vPaymentPlanNumber As Integer = 0
      Dim vBatchNumber As Integer = 0
      Dim vCancelList As New ParameterList(True)
      Dim vCreateNegativeSB As Boolean = False
      Dim vList As New ParameterList
      Select Case pItem
        Case ServiceBookingMenu.ServiceBookingMenuItems.sbiCancel
          If DataHelper.UserInfo.AccessLevel = UserInfo.UserAccessLevel.ualReadOnly Then
            ShowInformationMessage(InformationMessages.ImCancellationUnavailable) '"Cancellation is not available from here")
          Else
            vPaymentPlanNumber = IntegerValue(pDataRow("PaymentPlanNumber").ToString)
            If vPaymentPlanNumber > 0 Then
              Dim vPaymentPlan As New PaymentPlanInfo(vPaymentPlanNumber)
              If vPaymentPlan.DirectDebitStatus = "Y" Then vDefaults("CancelDirectDebit") = "Y"
              If vPaymentPlan.StandingOrderStatus = "Y" Then vDefaults("CancelStandingOrder") = "Y"
              If vPaymentPlan.CreditCardStatus = "Y" Then vDefaults("CancelCreditCardAuthority") = "Y"
              vDefaults("CancelStatusDate") = AppValues.TodaysDate
              vRenameList("CancelledOn") = ControlText.LblCancellationDate
            Else
              vBatchNumber = IntegerValue(pDataRow("BatchNumber").ToString)
              Dim vBatch As New BatchInfo(vBatchNumber)
              vCreateNegativeSB = Not (vBatch.BatchType = CareNetServices.BatchTypes.Cash OrElse vBatch.BatchType = CareNetServices.BatchTypes.CashWithInvoice) 'not paid by invoice so automatically create -ve SB data
              If vCreateNegativeSB Then
                Dim vWhereList As New ParameterList(True)
                vWhereList("ContactNumber") = pDataRow("BookingContactNumber").ToString
                vWhereList("StopCode") = String.Empty
                Dim vTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtCreditCustomers, vWhereList)
                'booking contact is a credit customer, give the user the option of automatically creating -ve data
                If vTable IsNot Nothing AndAlso vTable.Rows.Count > 0 Then vCreateNegativeSB = True
              End If
              If Not vCreateNegativeSB Then
                'booking contact is a credit customer, give the user the option of automatically creating -ve data
                If ShowQuestion(QuestionMessages.QmManuallyCreateServiceBookingCredit, MessageBoxButtons.YesNo) = DialogResult.Yes Then
                  vCreateNegativeSB = True
                Else
                  vCreateNegativeSB = False
                End If
              End If
            End If
            vCancelList("CreateNegativeSB") = CBoolYN(vCreateNegativeSB)
            vCancelList("ServiceBookingNumber") = pDataRow("ServiceBookingNumber").ToString
            vDefaults("CancellationDate") = pDataRow("EndDate").ToString
            vDefaults("ServiceBookingCancellation") = "Y"
            vList = FormHelper.ShowApplicationParameters(CareServices.FunctionParameterTypes.fptCancelPaymentPlan, vDefaults, vRenameList)
            If vList.Count > 0 Then
              vCancelList("CancellationReason") = vList("CancellationReason").ToString
              vCancelList("CancelledOn") = vList("CancellationDate").ToString
              If vList.ContainsKey("Source") Then vCancelList("Source") = vList("Source").ToString
              If vList.Contains("CancelDirectDebit") Then vCancelList("DirectDebit") = vList("CancelDirectDebit").ToString
              If vList.Contains("CancelStandingOrder") Then vCancelList("StandingOrder") = vList("CancelStandingOrder").ToString
              If vList.Contains("CancelCreditCardAuthority") Then vCancelList("CCCA") = vList("CancelCreditCardAuthority").ToString
              DataHelper.CancelServiceBooking(vCancelList)
            End If
            RefreshCard()
          End If
        Case ServiceBookingMenu.ServiceBookingMenuItems.sbiGoToServiceContact
          If IntegerValue(pDataRow("ServiceContactNumber").ToString) > 0 Then FormHelper.ShowCardIndex(CareServices.XMLContactDataSelectionTypes.xcdtNone, IntegerValue(pDataRow("ServiceContactNumber").ToString), True, False)
        Case ServiceBookingMenu.ServiceBookingMenuItems.sbiSBGoToRelatedContact
          If IntegerValue(pDataRow("RelatedContactNumber").ToString) > 0 Then FormHelper.ShowCardIndex(CareServices.XMLContactDataSelectionTypes.xcdtNone, IntegerValue(pDataRow("RelatedContactNumber").ToString), True, False)
      End Select
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub


  Private Sub ShowApplicationParametersHandler(ByVal pParentForm As MaintenanceParentForm, ByVal pApplicationParameterType As CareServices.FunctionParameterTypes, ByRef pParams As ParameterList)
    Dim vDefaults As New ParameterList
    Select Case pApplicationParameterType
      Case CareServices.FunctionParameterTypes.fptLeavePosition, CareServices.FunctionParameterTypes.fptMovePosition
        vDefaults("Finished") = AppValues.TodaysDateAddDays(-1)
        If pParams.Contains("ValidFrom") AndAlso pParams.Contains("ContactNumber") AndAlso pParams.Contains("AddressNumber") Then
          Dim vMinLeave As Date
          If Date.TryParse(pParams("ValidFrom"), vMinLeave) Then
            If Date.Parse(vDefaults("Finished")) < vMinLeave Then
              vDefaults("Finished") = vMinLeave.ToString(AppValues.DateFormat)
            End If
            vDefaults("MinimumLeavingDate") = vMinLeave.ToString(AppValues.DateFormat)
          End If
          Dim vTable As DataTable = DataHelper.GetTableFromDataSet(DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactAddresses, pParams.IntegerValue("ContactNumber")))
          For Each vDataRow As DataRow In vTable.Rows
            If IntegerValue(vDataRow.Item("AddressNumber").ToString) = pParams.IntegerValue("AddressNumber") Then
              Dim vMaxLeave As Date
              If Date.TryParse(vDataRow.Item("ValidTo").ToString, vMaxLeave) Then
                If Date.Parse(vDefaults("Finished")) > vMaxLeave Then
                  vDefaults("Finished") = vMaxLeave.ToString(AppValues.DateFormat)
                End If
                If Date.TryParse(vDefaults("Finished"), vMaxLeave) Then vDefaults("MaximumLeavingDate") = vMaxLeave.ToString(AppValues.DateFormat)
              End If
              Exit For
            End If
          Next
          If pApplicationParameterType = CareServices.FunctionParameterTypes.fptMovePosition AndAlso vTable.Rows.Count < 2 Then
            vDefaults("ChangeSite") = "N"
          End If
        End If
        vDefaults("Started") = Date.Parse(vDefaults("Finished")).AddDays(1).ToString(AppValues.DateFormat)
        If mvContactInfo.ContactType = CDBNETCL.ContactInfo.ContactTypes.ctOrganisation Then vDefaults("DisableChangeOrganisation") = "Y"
        pParams.Clear()
    End Select
    Dim vResultList As ParameterList = FormHelper.ShowApplicationParameters(pApplicationParameterType, vDefaults, , pParentForm)
    If vResultList IsNot Nothing AndAlso vResultList.Count > 0 Then
      If pApplicationParameterType = CareServices.FunctionParameterTypes.fptRemoveFutureMembershipType Then
        pParams("PriceIndicator") = vResultList("RunType")
      Else
        pParams = vResultList
      End If
    End If
  End Sub

  Private Sub mvPurchaseOrderMenu_MenuSelected(ByVal pItem As BaseFinancialMenu.FinancialMenuItems, ByVal pDataRow As System.Data.DataRow, ByVal pChangeDetails As Boolean, ByVal pFinancialMenu As BaseFinancialMenu) Handles mvPurchaseOrderMenu.MenuSelected
    Try
      Dim vRefresh As Boolean = False

      Select Case pItem
        Case FinancialMenu.FinancialMenuItems.fmiCancel
          Select Case mvDataType
            Case CareServices.XMLContactDataSelectionTypes.xcdtContactPurchaseOrders
              Dim vPONumber As Integer = IntegerValue(SelectRowItem(dgr.CurrentRow, "PurchaseOrderNumber"))
              If vPONumber > 0 Then
                Dim vList As New ParameterList
                vList = FormHelper.ShowApplicationParameters(CareServices.FunctionParameterTypes.fptCancellationReasonSourceAndDate, vList)
                If vList.Contains("CancellationReason") AndAlso vList("CancellationReason").Length > 0 Then
                  vList.IntegerValue("PurchaseOrderNumber") = vPONumber
                  vList = DataHelper.UpdateItem(CareServices.XMLMaintenanceControlTypes.xmctPurchaseOrders, vList)
                  vRefresh = True
                End If
              End If
          End Select
        Case BaseFinancialMenu.FinancialMenuItems.fmiAmendPurchaseOrder
          Dim vPONumber As Integer = IntegerValue(SelectRowItem(dgr.CurrentRow, "PurchaseOrderNumber"))
          If vPONumber > 0 Then
            Dim vList As New ParameterList
            vList.IntegerValue("PurchaseOrderNumber") = vPONumber
            Dim vTraderApplication As TraderApplication
            vTraderApplication = New TraderApplication(IntegerValue(AppValues.ConfigurationValue(AppValues.ConfigurationValues.trader_application_pom)))
            FormHelper.RunTraderApplication(vTraderApplication, vList)
            vRefresh = True
          End If
        Case FinancialMenu.FinancialMenuItems.fmiReinstate
          Select Case mvDataType
            Case CareServices.XMLContactDataSelectionTypes.xcdtContactPurchaseOrders
              Dim vPONumber As Integer = IntegerValue(SelectRowItem(dgr.CurrentRow, "PurchaseOrderNumber"))
              If vPONumber > 0 Then
                Dim vList As New ParameterList(True)
                vList("PurchaseOrderNumber") = vPONumber.ToString
                vList("CancellationReason") = ""
                DataHelper.UpdateItem(CareServices.XMLMaintenanceControlTypes.xmctPurchaseOrders, vList)
                vRefresh = True
              End If
          End Select
        Case BaseFinancialMenu.FinancialMenuItems.fmiAuthorisePurchaseOrder
          Dim vPONumber As Integer = IntegerValue(SelectRowItem(dgr.CurrentRow, "PurchaseOrderNumber"))
          If vPONumber > 0 Then
            Dim vList As New ParameterList(True)
            vList("PurchaseOrderNumber") = vPONumber.ToString
            vList("AuthorisedOn") = Date.Now.ToString(AppValues.DateTimeFormat)
            DataHelper.UpdateItem(CareServices.XMLMaintenanceControlTypes.xmctPurchaseOrders, vList)
            ShowInformationMessage(InformationMessages.ImAuthorisationComplete)
            vRefresh = True
          End If

      End Select
      If vRefresh Then RefreshCard()

    Catch vCareException As CareException
      Select Case vCareException.ErrorNumber
        Case CareException.ErrorNumbers.enTraderUnsupportedFeature
          ShowInformationMessage(vCareException.Message)
        Case Else
          DataHelper.HandleException(vCareException)
      End Select
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub mvPurchaseOrderPaymentMenu_MenuSelected(ByVal pItem As BaseFinancialMenu.FinancialMenuItems, ByVal pDataRow As System.Data.DataRow, ByVal pChangeDetails As Boolean, ByVal pFinancialMenu As BaseFinancialMenu) Handles mvPurchaseOrderPaymentMenu.MenuSelected
    Try
      Dim vRefresh As Boolean = False
      Select Case pItem
        Case BaseFinancialMenu.FinancialMenuItems.fmiNew, BaseFinancialMenu.FinancialMenuItems.fmiEdit
          Dim vList As New ParameterList
          vList("PurchaseOrderNumber") = dgr.GetValue(dgr.CurrentRow, "PurchaseOrderNumber")
          If pItem = BaseFinancialMenu.FinancialMenuItems.fmiNew Then
            vList("NonHistoricPopType") = "Y"
          End If
          Dim vForm As New frmCardMaintenance(Me, mvContactInfo, CareServices.XMLContactDataSelectionTypes.xcdtNone, mvDataSet3, pItem = BaseFinancialMenu.FinancialMenuItems.fmiEdit, dts.DisplayGrid(1).CurrentRow, CareServices.XMLMaintenanceControlTypes.xmctPurchaseOrderPayments, vList)
          ShowMaintenanceForm(vForm)

        Case BaseFinancialMenu.FinancialMenuItems.fmiAddPurchaseOrderPaymentReceipt
          Dim vList As New ParameterList
          vList("PurchaseOrderNumber") = dgr.GetValue(dgr.CurrentRow, "PurchaseOrderNumber")
          Dim vAmount As Double = DoubleValue(pDataRow("ExpectedReceiptAmount").ToString)
          If vAmount = 0 Then vAmount = DoubleValue(pDataRow("Amount").ToString)
          vList("Amount") = CStr(-vAmount)
          vList("ReceiptForPaymentNumber") = pDataRow("PaymentNumber").ToString
          vList("NonHistoricPopType") = "Y"
          Dim vForm As New frmCardMaintenance(Me, mvContactInfo, CareServices.XMLContactDataSelectionTypes.xcdtNone, mvDataSet3, pItem = BaseFinancialMenu.FinancialMenuItems.fmiEdit, dts.DisplayGrid(1).CurrentRow, CareServices.XMLMaintenanceControlTypes.xmctPurchaseOrderPayments, vList)
          ShowMaintenanceForm(vForm)


        Case BaseFinancialMenu.FinancialMenuItems.fmiAuthorise
          If dgr.GetValue(dgr.CurrentRow, "CancellationReason").Length > 0 Then
            ShowInformationMessage(InformationMessages.ImPOCancelled)
          Else
            Dim vList As ParameterList = FormHelper.ShowApplicationParameters(CareServices.FunctionParameterTypes.fptAuthorisePOPayment)
            If vList.Contains("AuthorisationStatus") Then
              vList.IntegerValue("PaymentNumber") = IntegerValue(dts.DisplayGrid(1).GetValue(dts.DisplayGrid(1).CurrentRow, "PaymentNumber"))
              vList.IntegerValue("PurchaseOrderNumber") = IntegerValue(dts.DisplayGrid(1).GetValue(dts.DisplayGrid(1).CurrentRow, "PurchaseOrderNumber"))
              DataHelper.UpdateItem(CareServices.XMLMaintenanceControlTypes.xmctPurchaseOrderPayments, vList)
              vRefresh = True
            End If
          End If
          'BR17340
        Case BaseFinancialMenu.FinancialMenuItems.fmiGoToPopRevChangedBy, BaseFinancialMenu.FinancialMenuItems.fmiGoToPopRevGoToChanges
          Dim mvParams As New ParameterList(True)
          mvParams("PurchaseOrderNumber") = dgr.GetValue(dgr.CurrentRow, "PurchaseOrderNumber")
          mvParams("PaymentNumber") = pDataRow("PaymentNumber").ToString
          Dim vResults As ParameterList = DataHelper.ReversePOPMenuSelection(mvParams)
          SelectPurchaseOrder(CInt(vResults("PONumRev")), CInt(vResults("PaymentNumRev")))

        Case BaseFinancialMenu.FinancialMenuItems.fmiCancelPOP
          Dim vPONumber As Integer = IntegerValue(SelectRowItem(dgr.CurrentRow, "PurchaseOrderNumber"))
          If vPONumber > 0 Then
            Dim vList As New ParameterList
            Dim vResult As DialogResult = System.Windows.Forms.DialogResult.Yes
            Dim vAdjustDate As String = AppValues.ConfigurationValue(AppValues.ConfigurationValues.fp_adjust_transaction_date)
            If LCase(vAdjustDate) = "today" Then
              vList("CancelledOn") = Now.ToString
            ElseIf LCase(vAdjustDate) = "original" Then
              vList("CancelledOn") = dgr.GetValue(dgr.CurrentRow, "StartDate") 'need to get original transaction date
            End If

            vList = FormHelper.ShowApplicationParameters(CareServices.FunctionParameterTypes.fptCancellationReasonSourceAndDate, vList)
            If vList.Contains("CancellationReason") AndAlso vList("CancellationReason").Length > 0 Then
              vList("PurchaseOrderNumber") = CStr(vPONumber)
              vList.IntegerValue("PaymentNumber") = IntegerValue(dts.DisplayGrid(1).GetValue(dts.DisplayGrid(1).CurrentRow, "PaymentNumber"))

              'see if there are multiple records affected
              vList("MultipleRecsCheck") = "Y"
              Dim vResults As ParameterList = DataHelper.ReversePurchaseOrderPayment(vList)
              If vResults.ContainsKey("MultipleRecsNo") And vResults("MultipleRecsNo").ToString <> "0" Then
                vResult = ShowQuestion(QuestionMessages.QmConfirmPOPMultipleRecs, MessageBoxButtons.YesNo, vResults("MultipleRecsNo"))
              End If

              If vResult = System.Windows.Forms.DialogResult.Yes Then
                vList("MultipleRecsCheck") = "N"
                Dim vDT As DataTable = Nothing
                Dim vProcess As New AsyncProcessHandler(AsyncProcessHandler.AsyncProcessHandlerTypes.ReversePurchaseOrderPayment, vList)
                vDT = DataHelper.GetTableFromDataSet(vProcess.GetDataSetFromResult)
                vRefresh = True
              End If
            End If
          End If

        Case BaseFinancialMenu.FinancialMenuItems.fmiPOPAnalysis
          If SelectRowItemNumber(dgr.CurrentRow, "PurchaseOrderNumber") > 0 AndAlso SelectRowItem(dgr.CurrentRow, "CancellationReason").Length = 0 Then
            Dim vDefaults As New ParameterList
            vDefaults("DueDate") = Now.ToString
            If AppValues.ConfigurationValue(AppValues.ConfigurationValues.fp_adjust_transaction_date, "today").ToLower = "original" Then vDefaults("DueDate") = pDataRow.Item("DueDate").ToString
            vDefaults("OriginalPoPaymentType") = pDataRow.Item("PaymentType").ToString
            vDefaults("OriginalDueDate") = pDataRow.Item("DueDate").ToString
            vDefaults("PaymentTypeRequired") = SelectRowItem(dgr.CurrentRow, "POPaymentTypeRequired")
            vDefaults("NonHistoricPopType") = "Y"
            Dim vResults As ParameterList = FormHelper.ShowApplicationParameters(CareNetServices.FunctionParameterTypes.fptPOPAnalysis, vDefaults, , Me)
            If vResults IsNot Nothing AndAlso vResults.Count > 0 Then
              vResults.IntegerValue("PurchaseOrderNumber") = SelectRowItemNumber(dgr.CurrentRow, "PurchaseOrderNumber")
              vResults.IntegerValue("PaymentNumber") = IntegerValue(pDataRow.Item("PaymentNumber").ToString)
              vResults = DataHelper.ReanalysePurchaseOrderPayment(vResults)
              ShowInformationMessage(InformationMessages.ImPOPAnalysisSuccessful, vResults("PurchaseOrderNumber"), vResults("PurchaseInvoiceNumber"), vResults("PaymentNumber"))
              vRefresh = True
            End If
          End If

      End Select
      If vRefresh Then RefreshCard()

    Catch vCareException As CareException
      Select Case vCareException.ErrorNumber
        Case CareException.ErrorNumbers.enPOPPayByBacsNoDefaultBankAccount
          ShowErrorMessage(InformationMessages.ImPOPPayByBacsNoSingleDefaultBankAccount)
        Case Else
          DataHelper.HandleException(vCareException)
      End Select
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub mvEventDelegateMenu_MenuSelected(ByVal pItem As BaseFinancialMenu.FinancialMenuItems, ByVal pDataRow As System.Data.DataRow, ByVal pChangeDetails As Boolean, ByVal pFinancialMenu As BaseFinancialMenu) Handles mvContactEventDelegateMenu.MenuSelected
    Try
      Select Case pItem
        Case BaseFinancialMenu.FinancialMenuItems.fmiGoToEvent
          FormHelper.ShowEventIndex(mvContactEventDelegateMenu.EventNumber(0))
        Case BaseFinancialMenu.FinancialMenuItems.fmiSupplementaryInformation
          Dim vSource As String
          If mvContactEventDelegateMenu.EventDelegateInfo.BookingTransactionExists Then
            vSource = mvContactEventDelegateMenu.EventDelegateInfo.TransactionSource
          Else
            vSource = mvContactEventDelegateMenu.CareEventInfo(0).SourceCode
          End If
          ShowDelegateDataSheet(Me, mvContactInfo, "D", vSource, mvContactEventDelegateMenu.CareEventInfo(0).ActivityGroup, mvContactEventDelegateMenu.CareEventInfo(0).RelationshipGroup, mvContactEventDelegateMenu.EventDelegateInfo.ContactName, mvContactEventDelegateMenu.EventDelegateInfo, False)
          dgr.SelectRow(dgr.CurrentRow)
      End Select
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub mvViewMailingDocumentMenu_MenuSelected(ByVal pItem As BaseFinancialMenu.FinancialMenuItems, ByVal pDataRow As System.Data.DataRow, ByVal pChangeDetails As Boolean, ByVal pFinancialMenu As BaseFinancialMenu) Handles mvViewMailingDocumentMenu.MenuSelected
    Try
      Select Case pItem
        Case BaseFinancialMenu.FinancialMenuItems.fmiViewMailingDocument
          ' Do the work here - see what we can display and inform user
          ViewMailingDocument(pDataRow)
        Case BaseFinancialMenu.FinancialMenuItems.fmiDeleteMailingDocument
          If Settings.ConfirmDelete AndAlso ShowQuestion(QuestionMessages.QmConfirmDelete, MessageBoxButtons.OKCancel) = System.Windows.Forms.DialogResult.Cancel Then Exit Sub
          DeleteContactMailingDocument(pDataRow)
          dgr.SetRowVisible(dgr.CurrentRow, False)
        Case BaseFinancialMenu.FinancialMenuItems.fmiUnfulfillMailingDocument
          If ShowQuestion(QuestionMessages.QmConfirmSetUnfulfilled, MessageBoxButtons.YesNo, "1") = System.Windows.Forms.DialogResult.No Then Exit Sub
          UnfulfillContactMailingDocument(pDataRow)
          Dim vCurrentRowValue As String = pDataRow("MailingNumber").ToString
          RefreshData()
          dgr.SelectRow("MailingNumber", vCurrentRowValue)
      End Select
    Catch vCareException As CareException
      Select Case vCareException.ErrorNumber
        Case CareException.ErrorNumbers.enMailingNumberConfig
          ShowErrorMessage(vCareException.Message)
        Case CareException.ErrorNumbers.enCMDAlreadyUnfulfilled, CareException.ErrorNumbers.enCMDCannotDeleteFulfilled,
        CareException.ErrorNumbers.enCMDFulfillmentHistoryInvalid, CareException.ErrorNumbers.enCMDInvalidDeleted
          ShowErrorMessage(vCareException.Message)
          RefreshData()
        Case Else
          DataHelper.HandleException(vCareException)
      End Select
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub mvPurchaseInvoiceChequeMenu_MenuSelected(ByVal pItem As BaseFinancialMenu.FinancialMenuItems, ByVal pDataRow As System.Data.DataRow, ByVal pChangeDetails As Boolean, ByVal pFinancialMenu As BaseFinancialMenu) Handles mvPurchaseInvoiceChequeMenu.MenuSelected
    Try
      Select Case pItem
        Case BaseFinancialMenu.FinancialMenuItems.fmiReissueCheque
          If ShowQuestion(QuestionMessages.QmChequeReissue, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No Then Exit Sub
          ReissueCheque(pDataRow)
          Dim vForm As New frmCardMaintenance(CType(CareNetServices.XMLMaintenanceControlTypes.xmctPurchaseInvoiceChequePayee, CareServices.XMLMaintenanceControlTypes), pDataRow)
          vForm.Parent = Nothing
          vForm.TopLevel = True
          vForm.ShowDialog()
          RefreshCard()
        Case BaseFinancialMenu.FinancialMenuItems.fmiChangeChequePayee
          Dim vForm As New frmCardMaintenance(CType(CareNetServices.XMLMaintenanceControlTypes.xmctPurchaseInvoiceChequePayee, CareServices.XMLMaintenanceControlTypes), pDataRow)
          vForm.Parent = Nothing
          vForm.TopLevel = True
          If vForm.ShowDialog <> DialogResult.Cancel Then RefreshCard()
        Case BaseFinancialMenu.FinancialMenuItems.fmiChequeSetStatus
          Dim vParamList As ParameterList = FormHelper.ShowApplicationParameters(CareNetServices.FunctionParameterTypes.fptSetChequeStatus)
          If vParamList.ContainsKey("ChequeStatus") Then
            Dim vParameterList As New ParameterList(True)
            vParameterList("ChequeReferenceNumber") = pDataRow("ChequeReferenceNumber").ToString
            vParameterList("ChequeStatus") = vParamList("ChequeStatus")
            DataHelper.UpdateCheque(vParameterList)
            RefreshCard()
          End If
        Case BaseFinancialMenu.FinancialMenuItems.fmiGoToPopRevChangedBy, BaseFinancialMenu.FinancialMenuItems.fmiGoToPopRevGoToChanges
          Dim mvParams As New ParameterList(True)
          mvParams("ChequeReferenceNumber") = dgr.GetValue(dgr.CurrentRow, "ChequeReferenceNumber")
          Dim vResults As ParameterList = DataHelper.ReversePOPMenuSelection(mvParams)
          SelectPurchaseInvoice(CInt(vResults("InvNumberRev")))
      End Select

    Catch vCareException As CareException
      Select Case vCareException.ErrorNumber
        Case CareException.ErrorNumbers.enChequeInvalid, CareException.ErrorNumbers.enChequeNumberNotSet, CareException.ErrorNumbers.enChequeNotPrinted,
        CareException.ErrorNumbers.enChequeReconciled
          ShowErrorMessage(vCareException.Message)
          RefreshCard()
        Case Else
          DataHelper.HandleException(vCareException)
      End Select
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Function ProcessContactMailingsMenu(ByVal pDataRow As DataRow) As Boolean
    ' Decides if the ContactMailings popup menu can be shown and populates MLV's 
    ' TODO - Handle multiple menu items
    Dim vMailingDocUnfulfilled As Boolean
    With pDataRow
      mvMailingNumber = IntegerValue(pDataRow("MailingNumber"))
      Dim vMailingType As String = pDataRow("Type").ToString
      mvMailingDocFulfilled = (vMailingType = "Printed")
      mvMailingDocOutgoing = (vMailingType = "Outgoing")
      vMailingDocUnfulfilled = (vMailingType = "Pending")
      If mvMailingDocOutgoing Then
        mvMailingDocOutgoing = AppValues.ConfigurationOption(AppValues.ConfigurationOptions.ma_auto_name_mailing_files, False)
        If mvMailingDocOutgoing Then
          mvMailingFilename = pDataRow("MailingFilename").ToString
        Else
          mvMailingFilename = ""
        End If
        'Do not allow the user to attempt to view Mailing Data if no filename has been found.
        mvMailingDocOutgoing = (mvMailingFilename.Length > 0)
        mvMailingDocReqVal = IntegerValue(pDataRow("ContactNumber"))
      Else
        mvMailingDocReqVal = mvMailingNumber
      End If
    End With
    ProcessContactMailingsMenu = (mvMailingNumber > 0) AndAlso (mvMailingDocFulfilled Or mvMailingDocOutgoing Or vMailingDocUnfulfilled)
  End Function

  Private Sub ViewMailingDocument(ByVal pDataRow As DataRow)
    Dim vMailingDisplay As String = AppValues.ConfigurationValue(AppValues.ConfigurationValues.ml_mailing_display)
    Dim vUseCAREForm As Boolean

    Dim vEmailDocument As Boolean = IntegerValue(pDataRow("CommunicationNumber").ToString) > 0
    Dim vParamList As New ParameterList(True)
    vParamList("MailingNumber") = pDataRow("MailingNumber").ToString
    Dim vMailingHistoryDocumentCount As Integer = DataHelper.GetCount(CareNetServices.XMLGetCountTypes.xgctMailingHistoryDocuments, vParamList)
    If vMailingHistoryDocumentCount > 0 AndAlso vEmailDocument = False Then
      Dim vMailingHistoryDocumentNumber As Integer = IntegerValue(pDataRow("MailingNumber"))
      Dim vMailingHistoryDocument As String = DataHelper.GetMailingHistoryDocumentFile(vMailingHistoryDocumentNumber, ".doc", False)
      Dim vExternalApplication As ExternalApplication
      Dim vMergeFileName As String = pDataRow("MailingFilename").ToString
      Dim vParameters As New ParameterList(True)
      If Not File.Exists(vMergeFileName) Then
        ShowWarningMessage(InformationMessages.ImMailingDataFileNotFound, vMergeFileName)
        Exit Sub
      End If
      vParameters("Mailing") = pDataRow("Mailing").ToString
      vParameters("MailingDocumentNumber") = pDataRow("MailingNumber").ToString
      vExternalApplication = GetDocumentApplication(".doc")
      If mvMailingDocReqVal.ToString.Length > 0 Then
        Dim vCSVFile As New FileReader(vMergeFileName)
        Dim vLineData As String = vCSVFile.ReadLine()
        Dim vStreamWriter As StreamWriter = Nothing
        Dim vContactIndex As Integer = 0
        Dim vOutFileName As String = DataHelper.GetTempFile(".csv")
        vStreamWriter = New StreamWriter(vOutFileName)
        For vIndex As Integer = 0 To vCSVFile.FieldCount - 1
          If vCSVFile.Fields(vIndex).ToString = "Contact Number" Then vContactIndex = vIndex
        Next
        vStreamWriter.WriteLine(vLineData)
        Dim vFound As Boolean
        While Not vCSVFile.EndOfFile And Not vFound
          vLineData = vCSVFile.ReadLine()
          If vCSVFile.Fields(vContactIndex).ToString = mvMailingDocReqVal.ToString Then
            vStreamWriter.WriteLine(vLineData)
          End If
        End While
        vStreamWriter.Close()
        vExternalApplication.MergeStandardDocument(vMailingHistoryDocument, vOutFileName, New DataTable(), True)
      Else
        vExternalApplication.MergeStandardDocument(vMailingHistoryDocument, vMergeFileName, New DataTable(), True)
      End If
    Else
      If mvMailingDocOutgoing Then
        vUseCAREForm = (vMailingDisplay = "MIXED" Or vMailingDisplay = "CARE_FORM")
      ElseIf mvMailingDocFulfilled Then
        'For contact mailing documents vMailingNumber is the Mailing Document Number.
        mvMailingFilename = pDataRow("MailingFilename").ToString()
        If mvMailingFilename.IsNullOrWhitespace Then
          mvMailingFilename = AppValues.GetMailingFileName(IntegerValue(pDataRow("FulfillmentNumber")))
        End If
        vUseCAREForm = (vMailingDisplay = "CARE_FORM")
      End If
      Dim vOutFileName As String = ""
      Dim vDataSet As DataSet = Nothing
      Dim vParams As ParameterList
      If mvMailingFilename.Length > 0 Then
        Dim vShowDocument As Boolean
        If Not File.Exists(mvMailingFilename) Then
          ShowWarningMessage(InformationMessages.ImMailingDataFileNotFound, mvMailingFilename)
        Else
          vShowDocument = True
          Dim vCSVFile As New FileReader(mvMailingFilename)
          Dim vLineData As String = vCSVFile.ReadLine()
          Dim vStreamWriter As StreamWriter = Nothing
          If vEmailDocument OrElse Not vUseCAREForm Then
            vOutFileName = DataHelper.GetTempFile(".csv")
            vStreamWriter = New StreamWriter(vOutFileName)
          End If
          vParams = New ParameterList()
          Dim vFields As ArrayListEx = vCSVFile.Fields
          Dim vIndex As Integer
          Dim vFieldRequired As Integer
          Dim vParamName As String
          For vIndex = 0 To vCSVFile.FieldCount - 1
            vParamName = vCSVFile.Item(vIndex)
            While vParams.Contains(vParamName)
              vParamName = vParamName & "_"
            End While
            vParams.Add(vParamName, "")
            If vParamName = "Contact Number" Then vFieldRequired = vIndex
          Next
          If vEmailDocument Then
            vStreamWriter.WriteLine(vLineData)
          ElseIf Not mvMailingDocOutgoing Then
            If Not vUseCAREForm Then vStreamWriter.WriteLine(vLineData)
            vFieldRequired = 0
          End If
          Dim vFound As Boolean
          While Not vCSVFile.EndOfFile And Not vFound
            vLineData = vCSVFile.ReadLine()
            If IntegerValue(vCSVFile.Item(vFieldRequired)) = mvMailingDocReqVal Then
              If vEmailDocument OrElse Not vUseCAREForm Then
                If mvMailingDocOutgoing And Not vEmailDocument Then
                  For vIndex = 0 To vFields.Count - 1
                    vStreamWriter.WriteLine(vFields(vIndex).ToString & Space(40 - vFields(vIndex).ToString.Length) & vCSVFile.Item(vIndex))
                  Next
                Else
                  vStreamWriter.WriteLine(vLineData)
                End If
                vStreamWriter.Close()
              Else
                ' Need to build the dataset for use with the 'care' form
                vDataSet = New DataSet()
                vDataSet.Tables.Add("DataRow")
                With vDataSet.Tables("DataRow")
                  .Columns.Add("Field")
                  .Columns.Add("Value")
                  For vIndex = 0 To vFields.Count - 1
                    .Rows.Add(vFields(vIndex).ToString, vCSVFile.Item(vIndex))
                  Next
                End With
              End If
              vFound = True
            End If
          End While
          If Not vFound Then
            ShowWarningMessage(InformationMessages.ImMailingDataNotFound, mvMailingFilename)
            If Not vUseCAREForm Then vStreamWriter.Close()
          End If
          vCSVFile.CloseFile()
        End If
        If vShowDocument Then
          Dim vExternalApplication As ExternalApplication
          Dim vSelectForm As frmSelectItems
          If mvMailingDocOutgoing Then
            If vEmailDocument Then
              Dim vStdDocument As String = "HTML"
              Dim vList As New ParameterList(True)
              vList("MailingNumber") = pDataRow("MailingNumber").ToString
              Dim vRow As DataRow = DataHelper.GetLookupDataRow(CareNetServices.XMLLookupDataTypes.xldtEmailJobs, vList)
              If vRow IsNot Nothing Then vStdDocument = vRow("StandardDocument").ToString
              Dim vSDList As New ParameterList(True)
              vSDList("StandardDocument") = vStdDocument
              vRow = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtStandardDocuments, vSDList).Rows(0)
              Dim vApplication As ExternalApplication = GetDocumentApplication(vRow.Item("DocFileExtension").ToString)
              Dim vMailingNumber As Integer = 0
              If vMailingHistoryDocumentCount > 0 Then vMailingNumber = IntegerValue(pDataRow("MailingNumber"))
              vApplication.MergeStandardDocument(vRow.Item("StandardDocument").ToString, vRow.Item("DocFileExtension").ToString, vOutFileName, False, True, False, vMailingNumber)
            ElseIf vUseCAREForm Then
              vSelectForm = New frmSelectItems(vDataSet, frmSelectItems.SelectItemsTypes.sitContactMailings)
              vSelectForm.ShowDialog(Me)
            Else
              vExternalApplication = GetDocumentApplication(".doc")
              AddHandler vExternalApplication.ActionComplete, AddressOf WordActionComplete
              Dim vFileInfo As New FileInfo(vOutFileName)
              vExternalApplication.ViewDocument(vFileInfo)
            End If
          Else
            If Not vUseCAREForm Then
              Dim vDocumentName As String = DataHelper.GetContactMailingDocumentFile(IntegerValue(pDataRow("MailingNumber")), ".doc")
              vExternalApplication = GetDocumentApplication(".doc")
              Dim vList As New ParameterList(True)
              vList("Mailing") = pDataRow("Mailing").ToString
              vList("MailingDocumentNumber") = pDataRow("MailingNumber").ToString
              vExternalApplication.MergeStandardDocument(vDocumentName, vOutFileName, DataHelper.GetTableFromDataSet(DataHelper.GetMailingDocumentParagraphs(vList)))
            Else
              vSelectForm = New frmSelectItems(vDataSet, frmSelectItems.SelectItemsTypes.sitContactMailings)
              vSelectForm.ShowDialog(Me)
            End If
          End If
          ' Temp file is deleted when the external application (Word) closes (WordActionComplete) below
          ' TODO: Consider other 'temp' files
        End If
      Else
        ShowWarningMessage(InformationMessages.ImMailingDataFileNotFound, mvMailingFilename)
      End If
    End If
  End Sub

  Private Sub WordActionComplete(ByVal pAction As ExternalApplication.DocumentActions, ByVal pFilename As String)
    DocumentApplication = Nothing
    Select Case pAction
      Case ExternalApplication.DocumentActions.daEditing
        ' NYI
      Case ExternalApplication.DocumentActions.daPrinting, ExternalApplication.DocumentActions.daViewing
        If File.Exists(pFilename) Then File.Delete(pFilename)
    End Select
  End Sub

  Private Sub DeleteContactMailingDocument(ByVal pDataRow As DataRow)
    DataHelper.DeleteContactMailingDocument(IntegerValue(pDataRow("MailingNumber")))
  End Sub

  Private Sub UnfulfillContactMailingDocument(ByVal pDataRow As DataRow)
    DataHelper.SetMailingDocumentUnfulfilled(IntegerValue(pDataRow("FulfillmentNumber")), "", IntegerValue(pDataRow("MailingNumber")))
  End Sub

  Private Sub ReissueCheque(ByVal pDataRow As DataRow)
    DataHelper.ReissueCheque(IntegerValue(pDataRow("ChequeReferenceNumber")))
  End Sub

  Private Function FundraisingActionsMaintenance(ByVal pSource As MenuSource, ByVal pEdit As Boolean) As frmCardMaintenance
    Dim vList As New ParameterList
    vList("FundraisingRequestNumber") = dgr.GetValue(dgr.CurrentDataRow, "FundraisingRequestNumber")
    vList("Logname") = dgr.GetValue(dgr.CurrentDataRow, "Logname")
    If pSource = MenuSource.dgr OrElse pSource = MenuSource.actiondgr4 Then
      vList("ActionDesc") = dgr.GetValue(dgr.CurrentDataRow, "RequestDescription")
    ElseIf pSource = MenuSource.actiondgr0 AndAlso dts.DisplayGrid(0).GetValue(dts.DisplayGrid(0).CurrentDataRow, "ScheduledPaymentNumber").Length > 0 Then
      vList("ActionDesc") = dts.DisplayGrid(0).GetValue(dts.DisplayGrid(0).CurrentDataRow, "ScheduledPaymentDesc")
      vList("ScheduledPaymentNumber") = dts.DisplayGrid(0).GetValue(dts.DisplayGrid(0).CurrentDataRow, "ScheduledPaymentNumber")
      vList("HideGrid") = "Y"
    End If
    Return New frmCardMaintenance(Me, mvContactInfo, mvDataType, mvDataSet4, pEdit, dts.DisplayGrid(4).CurrentDataRow, CareServices.XMLMaintenanceControlTypes.xmctAction, vList)
  End Function

#Region "Accessibility"

  Public Overrides Sub SetActiveChildControl()
    If Me.Enabled Then
      If dgr.ContainsFocus Then
        mvActiveChildControl = dgr
      Else
        mvActiveChildControl = Nothing
      End If
    End If
  End Sub
  Public Overrides ReadOnly Property GetActiveChildControl() As Control
    Get
      Return mvActiveChildControl
    End Get
  End Property
  Private Sub dgrMenuStrip_Closed(ByVal sender As Object, ByVal e As System.Windows.Forms.ToolStripDropDownClosedEventArgs) Handles dgrMenuStrip.Closed, dgr0MenuStrip.Closed, dgr1MenuStrip.Closed, dplMenuStrip.Closed, dgr2MenuStrip.Closed, dgr3MenuStrip.Closed, dgr4MenuStrip.Closed, dgr5MenuStrip.Closed
    If mvAccessibilityRoleReset Then dgr.AccessibleRole = mvGridAccessibleRole
  End Sub
  Private Sub dgrMenuStrip_Closing(ByVal sender As Object, ByVal e As System.Windows.Forms.ToolStripDropDownClosingEventArgs) Handles dgrMenuStrip.Closing, dgr0MenuStrip.Closing, dgr1MenuStrip.Closing, dplMenuStrip.Closing, dgr2MenuStrip.Closing, dgr3MenuStrip.Closing, dgr4MenuStrip.Closing, dgr5MenuStrip.Closing
    mvRightSplitterAccessibleRole = splRight.AccessibleRole
    splRight.AccessibleRole = System.Windows.Forms.AccessibleRole.None
    dgr.ResetAccessibility()
    dgr.ReadAccessibility = False
    mvGridAccessibleRole = dgr.AccessibleRole
    dgr.AccessibleRole = System.Windows.Forms.AccessibleRole.None
    mvAccessibilityRoleReset = True
  End Sub
#End Region

  Private Sub mvFrmNewMember_MemberActionCompleted(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles mvFrmNewMember.FormClosed
    Dim vPaymentPlanNumber As Integer = CType(sender, frmNewMember).PPNumber
    mvDataSet4 = DataHelper.GetPaymentPlanData(CareServices.XMLPaymentPlanDataSelectionTypes.xpdtPaymentPlanMembers, vPaymentPlanNumber)
    dts.DisplayGrid(2).Populate(mvDataSet4)
  End Sub
  Private Sub mvFinancialSubMenu_MenuSelected(ByVal pItem As BaseFinancialMenu.FinancialMenuItems, ByVal pDataRow As System.Data.DataRow) Handles mvFinancialSubMenu.MenuSelected
    Try
      Select Case pItem
        Case BaseFinancialMenu.FinancialMenuItems.fmiRemoveAllocations
          Dim vRefresh As Boolean
          RemoveInvoiceAllocations(mvDataSet3, dts.SubDisplayGrid(0).CurrentRow, vRefresh)
          If vRefresh Then RefreshCard()
        Case BaseFinancialMenu.FinancialMenuItems.fmiReplaceMember
          Dim vPaymentPlanMemberInfo As New PaymentPlanMemberInfo(pDataRow)
          mvFrmNewMember = New frmNewMember(vPaymentPlanMemberInfo, False)
          mvFrmNewMember.Show()
        Case BaseFinancialMenu.FinancialMenuItems.fmiAddMember
          Dim vPaymentPlanMemberInfo As New PaymentPlanMemberInfo(pDataRow)
          mvFrmNewMember = New frmNewMember(vPaymentPlanMemberInfo, True)
          mvFrmNewMember.Show()
        Case BaseFinancialMenu.FinancialMenuItems.fmiGoToTransaction
          mvTransactionLinkMenu_MenuSelected(pItem, pDataRow, False, mvFinancialSubMenu)
      End Select
    Catch vCareException As CareException
      Select Case vCareException.ErrorNumber
        Case CareException.ErrorNumbers.enCannotRemoveSundryCreditNoteReversal
          ShowInformationMessage(vCareException.Message)
        Case Else
          DataHelper.HandleException(vCareException)
      End Select
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub sel_RevertCustomisationChanges(ByVal Sender As Object, ByVal pResult As String) Handles sel.RevertCustomisationChanges

    Dim vParams As New ParameterList(True)
    Try
      If ShowQuestion(QuestionMessages.QmRevertModule, MessageBoxButtons.OKCancel) = DialogResult.OK Then
        vParams.Add("DataSelectionType", sel.DataSelectionType.ToString)
        vParams.Add("AccessMethod", "S")
        vParams.Add(mvContactInfo.ContactGroupParameterName, mvContactInfo.ContactGroup)
        vParams.Add("Logname", DataHelper.UserInfo.Logname.ToString)
        vParams.Add("Department", DataHelper.UserInfo.Department.ToString)
        vParams.Add("Client", DataHelper.GetClientCode())
        vParams.Add("WebPageItemNumber", "")
        DataHelper.DeleteItem(CareNetServices.XMLMaintenanceControlTypes.xmctDisplayListItem, vParams)
        Dim vEntityGroup As EntityGroup = DataHelper.ContactAndOrganisationGroups(mvContactInfo.ContactGroup)
        vEntityGroup.ResetSelectionPages()
        sel.Init(mvContactInfo, True)
      End If
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub dplMenuCustomise_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles dplMenuCustomise.Click
    HandleMenuClick(True, MenuSource.dplCustomise)
    RefreshCard()
  End Sub

  Private Sub dplMenuRevert_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles dplMenuRevert.Click
    HandleMenuClick(True, MenuSource.dplRevert)
    RefreshCard()
  End Sub

  Private Sub hdr_RefreshHeader(ByVal sender As Object) Handles hdr.RefreshHeader
    Dim vDataSet As DataSet = DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactHeaderInformation, mvContactInfo.ContactNumber)
    hdr.Populate(vDataSet, mvContactInfo)
    splTop.SplitterDistance = hdr.Height
  End Sub

  Private Sub RemoveInvoiceAllocations(ByVal pDataSet As DataSet, ByVal pRow As Integer, ByRef pRefresh As Boolean)
    Dim vType As String = GetDataRow(pDataSet, pRow).Item("TransactionType").ToString
    Dim vMessage As String
    Select Case vType
      Case "Invoice"
        vMessage = QuestionMessages.QmRemoveAllocationsInvoice
      Case "Payment"
        vMessage = QuestionMessages.QmRemoveAllocationsPayment
      Case Else
        vMessage = QuestionMessages.QmRemoveAllocationsCreditNote
    End Select
    If ShowQuestion(vMessage, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
      Dim vList As New ParameterList(True)
      vList("RemoveAllocations") = "Y"
      Dim vReturnList As ParameterList = DataHelper.UpdateInvoice(CInt(GetDataRow(pDataSet, pRow).Item("StoredInvoiceNumber")), vList)
      Dim vSB As New StringBuilder
      vSB.AppendLine(InformationMessages.ImAllocationsRemoved)
      If vReturnList.ContainsKey("Reference") Then vSB.AppendLine(GetInformationMessage(InformationMessages.ImAllocationsReference, vReturnList("Reference")))
      ShowInformationMessage(vSB.ToString)
      pRefresh = True
    End If
  End Sub

  Private Sub BeforeSelect(ByVal pSender As Object, ByRef pCancel As Boolean) Handles sel.BeforeSelect
    If mvCareWebBrowser IsNot Nothing AndAlso mvCareWebBrowser.ConfirmNavigateAway Then
      If Not ConfirmCancel() Then pCancel = True
      If Not pCancel Then mvCareWebBrowser = Nothing
    End If
  End Sub

  Private Sub mvDocumentMenu_ShowRelatedDocument(ByVal Sender As Object) Handles mvDocumentMenu.ShowRelatedDocument
    If mvDocumentMenu.DocumentNumber > 0 Then
      Dim vList As New ParameterList
      vList.IntegerValue("CommunicationsLogNumber1") = mvDocumentMenu.DocumentNumber
      vList("FinderCaption") = ControlText.FrmRelatedDocumentsFinder
      FormHelper.ShowFinder(CareNetServices.XMLDataFinderTypes.xdftDocuments, vList)
    End If
  End Sub

  Private Sub mvPurchaseInvoiceMenu_MenuSelected(pItem As BaseFinancialMenu.FinancialMenuItems, pDataRow As System.Data.DataRow, pChangeDetails As Boolean, pFinancialMenu As BaseFinancialMenu) Handles mvPurchaseInvoiceMenu.MenuSelected
    Try
      Dim vRefresh As Boolean = False

      Select Case pItem
        Case FinancialMenu.FinancialMenuItems.fmiGoToPopRevChangedBy, FinancialMenu.FinancialMenuItems.fmiGoToPopRevGoToChanges
          Dim mvParams As New ParameterList(True)
          mvParams("PurchaseInvoiceNumber") = dgr.GetValue(dgr.CurrentRow, "PurchaseInvoiceNumber")
          Dim vResults As ParameterList = DataHelper.ReversePOPMenuSelection(mvParams)
          SelectPurchaseInvoice(CInt(vResults("InvNumberRev")))
      End Select
      If vRefresh Then RefreshCard()

    Catch vCareException As CareException
      Select Case vCareException.ErrorNumber
        Case CareException.ErrorNumbers.enTraderUnsupportedFeature
          ShowInformationMessage(vCareException.Message)
        Case Else
          DataHelper.HandleException(vCareException)
      End Select
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Public Sub EntitySelected(pSender As Object, pEntityNumber As Integer, Optional pEntityType As HistoryEntityTypes = CDBNETCL.HistoryEntityTypes.hetContacts) Implements IDashboardTabContainer.EntitySelected
    MainHelper.NavigateHistoryItem(pEntityType, pEntityNumber, True)
  End Sub

  ''' <summary>Add and remove additional menu items to <see cref="dgr0MenuStrip">dgr0MenuStrip</see> depending upon the Data Selection Type.</summary>
  Private Sub dgr0MenuStripBuilder()
    'Now build menu

    If mvDgr0IndependentItemNames Is Nothing Then 'No appending done yet. Let's note menu items for this object 
      mvDgr0IndependentItemNames = New List(Of String)
      For Each vToolstripItem As ToolStripItem In dgr0MenuStrip.Items
        mvDgr0IndependentItemNames.Add(vToolstripItem.Name)
      Next
    End If

    'Delete menu items that are for previous associated object
    Dim vDependentToolStripItems As New List(Of String)
    For Each vToolstripItem As ToolStripItem In dgr0MenuStrip.Items
      If Not mvDgr0IndependentItemNames.Contains(vToolstripItem.Name) Then
        vDependentToolStripItems.Add(vToolstripItem.Name)
      End If
    Next
    For Each vName As String In vDependentToolStripItems
      dgr0MenuStrip.Items.RemoveByKey(vName)
    Next

    If dts.DisplayGrid(0).DataRowCount > 0 Then
      'Merge object and new associated object menu items into menu 
      If mvDataType = CareNetServices.XMLContactDataSelectionTypes.xcdtContactMeetings Then
        'MEETING DOCUMENTS
        'Add menu items if meeting documents exist - if display grid row has a ContactType = "" then this is a link to a meeting document 
        Dim vContactTypeColumn As Integer = dts.DisplayGrid(0).GetColumn("ContactType")
        Dim vContactNumberColumn As Integer = dts.DisplayGrid(0).GetColumn("ContactNumber")
        If vContactTypeColumn >= 0 AndAlso vContactNumberColumn >= 0 Then
          If dts.DisplayGrid(0).GetValue(dts.DisplayGrid(0).CurrentRow, "ContactType").ToString.Length = 0 Then 'Since MeetingDocument links has no ContactType value set by the server this line will pick links to document
            Dim vCommunicationsLogNumber As Integer
            vCommunicationsLogNumber = IntegerValue(dts.DisplayGrid(0).GetValue(dts.DisplayGrid(0).CurrentRow, "ContactNumber")) 'The server had put communications_log_number in ContactNumber column. 
            mvDgr0ContextCommandHelper = New ContextCommandHelper(ContextCommandHelper.AssociatedObject.cciCommunicationsLog, vCommunicationsLogNumber)
            If mvDgr0ContextCommandHelper.ContextCommandItems.Count > 0 Then
              MenuToolbarCommand.SetAccessControl(mvDgr0ContextCommandHelper.ContextCommandItems)
              Dim vGridMenuItem As ToolStripItem = dgr0MenuStrip.Items("Grid")
              dgr0MenuStrip.Items.RemoveByKey("Grid")
              For Each vItem As MenuToolbarCommand In mvDgr0ContextCommandHelper.ContextCommandItems
                vItem.OnClick = AddressOf dgr0MenuStrip_MenuHandler
                Dim vMenuItem As ToolStripMenuItem = vItem.MenuStripItem
                If Not dgr0MenuStrip.Items.ContainsKey(vMenuItem.Name) Then
                  vMenuItem.Visible = Not vItem.HideItem
                  If vItem.HideItem Then vMenuItem.Enabled = False
                  dgr0MenuStrip.Items.Add(vMenuItem)
                End If
              Next
              If vGridMenuItem IsNot Nothing Then dgr0MenuStrip.Items.Add(vGridMenuItem)
            End If
            'Enable or disable menu items for associated object
            mvDgr0ContextCommandHelper.SetVisibleItems(dgr0MenuStrip)
          End If
        End If
      End If

    End If
  End Sub

  Private Sub dgr0MenuStrip_MenuHandler(ByVal sender As System.Object, ByVal e As System.EventArgs)
    Dim vCursor As New BusyCursor
    Try
      Dim vContextCommand As MenuToolbarCommand = DirectCast(DirectCast(sender, ToolStripMenuItem).Tag, MenuToolbarCommand)
      mvDgr0ContextCommandHelper.MenuHandler(vContextCommand, Me, e)
    Catch vEx As CareException
      If vEx.ErrorNumber = CareException.ErrorNumbers.enCannotDeleteDocument Then
        ShowInformationMessage(vEx.Message)
      Else
        DataHelper.HandleException(vEx)
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try

  End Sub

  Private Sub dgr_TimesheetSelected(pSender As Object, pRow As Integer, pContactPositionTimesheetNumber As Integer) Handles dts.TimesheetSelected
    HandleMenuClick(True, MenuSource.dgr5)
  End Sub

  Private Sub RunMailmerge(ByVal pMenuItem As FinancialMenu.FinancialMenuItems, ByVal pDataRow As DataRow)
    Dim vList As New ParameterList(True)
    Dim vMergeType As CareNetServices.TraderMailmergeType = CareNetServices.TraderMailmergeType.tmtInvoice
    Dim vSTDDocumentCode As String = String.Empty

    Select Case pMenuItem
      Case BaseFinancialMenu.FinancialMenuItems.fmiPreviewInvoice
        vList.IntegerValue("BatchNumbers") = IntegerValue(pDataRow("BatchNumber").ToString)
        vList.IntegerValue("TransactionNumbers") = IntegerValue(pDataRow("TransactionNumber").ToString)
        'From and to invoice numbers need to be zero
        vList.IntegerValue("FromInvoiceNumber") = 0
        vList.IntegerValue("ToInvoiceNumber") = 0
        vList("Company") = SelectRowItem(dgr.CurrentRow, "Company")
        vList("PrintPreview") = "Y"
        vList("SalesLedgerPreviewInvoice") = "Y"
        vSTDDocumentCode = AppValues.ControlValue(AppValues.ControlTables.financial_controls, AppValues.ControlValues.preview_invoice_std_document)

      Case BaseFinancialMenu.FinancialMenuItems.fmiPrintReceipt
        vList.IntegerValue("BatchNumber") = IntegerValue(pDataRow("BatchNumber").ToString)
        vList.IntegerValue("TransactionNumber") = IntegerValue(pDataRow("TransactionNumber").ToString)
        vMergeType = CareNetServices.TraderMailmergeType.tmtReceipt
        vSTDDocumentCode = AppValues.ControlValue(AppValues.ControlTables.financial_controls, AppValues.ControlValues.receipt_print_std_document)

    End Select

    If vSTDDocumentCode.Length > 0 Then
      Dim vFileName As String = DataHelper.GetTempFile(".csv")
      If DataHelper.GetTraderMailingFile(vMergeType, vList, vFileName) Then
        'We have the mailmerge file, now need to perform the mailmerge
        'This will run Invoice Mailmerge with the standard document
        vList = New ParameterList(True)
        vList("StandardDocument") = vSTDDocumentCode
        Dim vRow As DataRow = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtStandardDocuments, vList).Rows(0)
        Dim vExtApplication As ExternalApplication = GetDocumentApplication(vRow.Item("DocFileExtension").ToString)
        'Print preview
        vExtApplication.MergeStandardDocument(vRow.Item("StandardDocument").ToString, vRow.Item("DocFileExtension").ToString, vFileName, False, True, False, True)
      End If
    End If

  End Sub
End Class
