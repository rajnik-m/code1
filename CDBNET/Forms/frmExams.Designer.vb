<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmExams
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
    Me.components = New System.ComponentModel.Container()
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmExams))
    Me.splTop = New System.Windows.Forms.SplitContainer()
    Me.splBottom = New System.Windows.Forms.SplitContainer()
    Me.sel = New CDBNETCL.ExamSelector()
    Me.cboSessions = New System.Windows.Forms.ComboBox()
    Me.pnlCommands = New CDBNETCL.PanelEx()
    Me.imgExemptions = New CDBNETCL.ImageButton()
    Me.imgSessions = New CDBNETCL.ImageButton()
    Me.imgCentres = New CDBNETCL.ImageButton()
    Me.imgPersonnel = New CDBNETCL.ImageButton()
    Me.imgCourse = New CDBNETCL.ImageButton()
    Me.tabMaster = New System.Windows.Forms.TabControl()
    Me.tabMain = New System.Windows.Forms.TabPage()
    Me.splRight = New System.Windows.Forms.SplitContainer()
    Me.UpperEditPanel = New CDBNETCL.EditPanel()
    Me.dgr = New CDBNETCL.DisplayGrid()
    Me.dpl = New CDBNETCL.DisplayPanel()
    Me.dgrDetails = New CDBNETCL.DisplayGrid()
    Me.epl = New CDBNETCL.EditPanel()
    Me.selExams = New CDBNETCL.ExamSelector()
    Me.bpl = New CDBNETCL.ButtonPanel()
    Me.cmdLink = New System.Windows.Forms.Button()
    Me.cmdAnalysis = New System.Windows.Forms.Button()
    Me.cmdSave = New System.Windows.Forms.Button()
    Me.cmdAllocate = New System.Windows.Forms.Button()
    Me.cmdUnallocate = New System.Windows.Forms.Button()
    Me.cmdSelectAll = New System.Windows.Forms.Button()
    Me.cmdUnSelectAll = New System.Windows.Forms.Button()
    Me.cmdNew = New System.Windows.Forms.Button()
    Me.cmdNewChild = New System.Windows.Forms.Button()
    Me.cmdDelete = New System.Windows.Forms.Button()
    Me.cmdClose = New System.Windows.Forms.Button()
    Me.tabCustomForm = New System.Windows.Forms.TabPage()
    Me.pnlCustomFormPage = New CDBNETCLPages.GridPanelPage()
    Me.dgrMenuStrip = New System.Windows.Forms.ContextMenuStrip(Me.components)
    Me.dgrMenuNew = New System.Windows.Forms.ToolStripMenuItem()
    Me.dgrMenuEdit = New System.Windows.Forms.ToolStripMenuItem()
    Me.dgrMenuNewDocument = New System.Windows.Forms.ToolStripMenuItem()
    Me.dgr2MenuStrip = New System.Windows.Forms.ContextMenuStrip(Me.components)
    Me.dgr2MenuNew = New System.Windows.Forms.ToolStripMenuItem()
    Me.dgr2MenuEdit = New System.Windows.Forms.ToolStripMenuItem()
    Me.dplMenuStrip = New System.Windows.Forms.ContextMenuStrip(Me.components)
    Me.dplMenuNew = New System.Windows.Forms.ToolStripMenuItem()
    Me.dplMenuEdit = New System.Windows.Forms.ToolStripMenuItem()
    Me.dplMenuCustomise = New System.Windows.Forms.ToolStripMenuItem()
    Me.dplMenuRevert = New System.Windows.Forms.ToolStripMenuItem()
    Me.dgr1MenuStrip = New System.Windows.Forms.ContextMenuStrip(Me.components)
    Me.dgr1MenuNew = New System.Windows.Forms.ToolStripMenuItem()
    Me.dgr1MenuEdit = New System.Windows.Forms.ToolStripMenuItem()
    Me.dgr0MenuStrip = New System.Windows.Forms.ContextMenuStrip(Me.components)
    Me.dgr0MenuNew = New System.Windows.Forms.ToolStripMenuItem()
    Me.dgr0MenuEdit = New System.Windows.Forms.ToolStripMenuItem()
    Me.splTab = New System.Windows.Forms.Splitter()
    Me.dspTabGrid = New CDBNETCL.DisplayTabSet()
    CType(Me.splTop, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.splTop.Panel2.SuspendLayout()
    Me.splTop.SuspendLayout()
    CType(Me.splBottom, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.splBottom.Panel1.SuspendLayout()
    Me.splBottom.Panel2.SuspendLayout()
    Me.splBottom.SuspendLayout()
    Me.pnlCommands.SuspendLayout()
    Me.tabMaster.SuspendLayout()
    Me.tabMain.SuspendLayout()
    CType(Me.splRight, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.splRight.Panel1.SuspendLayout()
    Me.splRight.Panel2.SuspendLayout()
    Me.splRight.SuspendLayout()
    Me.bpl.SuspendLayout()
    Me.tabCustomForm.SuspendLayout()
    Me.dgrMenuStrip.SuspendLayout()
    Me.dgr2MenuStrip.SuspendLayout()
    Me.dplMenuStrip.SuspendLayout()
    Me.dgr1MenuStrip.SuspendLayout()
    Me.dgr0MenuStrip.SuspendLayout()
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
    Me.splTop.Size = New System.Drawing.Size(783, 446)
    Me.splTop.SplitterDistance = 62
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
    Me.splBottom.Panel1.Controls.Add(Me.cboSessions)
    Me.splBottom.Panel1.Controls.Add(Me.pnlCommands)
    '
    'splBottom.Panel2
    '
    Me.splBottom.Panel2.Controls.Add(Me.tabMaster)
    Me.splBottom.Size = New System.Drawing.Size(783, 380)
    Me.splBottom.SplitterDistance = 184
    Me.splBottom.TabIndex = 0
    '
    'sel
    '
    Me.sel.AutoSelectParents = False
    Me.sel.BackColor = System.Drawing.Color.Transparent
    Me.sel.Dock = System.Windows.Forms.DockStyle.Fill
    Me.sel.ExamMaintenance = False
    Me.sel.Location = New System.Drawing.Point(0, 21)
    Me.sel.Name = "sel"
    Me.sel.Size = New System.Drawing.Size(184, 164)
    Me.sel.TabIndex = 3
    Me.sel.TreeContextMenu = Nothing
    '
    'cboSessions
    '
    Me.cboSessions.Dock = System.Windows.Forms.DockStyle.Top
    Me.cboSessions.FormattingEnabled = True
    Me.cboSessions.Location = New System.Drawing.Point(0, 0)
    Me.cboSessions.Name = "cboSessions"
    Me.cboSessions.Size = New System.Drawing.Size(184, 21)
    Me.cboSessions.TabIndex = 2
    '
    'pnlCommands
    '
    Me.pnlCommands.BackColor = System.Drawing.Color.Transparent
    Me.pnlCommands.Controls.Add(Me.imgExemptions)
    Me.pnlCommands.Controls.Add(Me.imgSessions)
    Me.pnlCommands.Controls.Add(Me.imgCentres)
    Me.pnlCommands.Controls.Add(Me.imgPersonnel)
    Me.pnlCommands.Controls.Add(Me.imgCourse)
    Me.pnlCommands.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.pnlCommands.Location = New System.Drawing.Point(0, 185)
    Me.pnlCommands.Name = "pnlCommands"
    Me.pnlCommands.Size = New System.Drawing.Size(184, 195)
    Me.pnlCommands.TabIndex = 1
    '
    'imgExemptions
    '
    Me.imgExemptions.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.imgExemptions.BackColor = System.Drawing.Color.Transparent
    Me.imgExemptions.DialogResult = System.Windows.Forms.DialogResult.OK
    Me.imgExemptions.DrawBorder = False
    Me.imgExemptions.Location = New System.Drawing.Point(3, 156)
    Me.imgExemptions.Name = "imgExemptions"
    Me.imgExemptions.NormalImage = CType(resources.GetObject("imgExemptions.NormalImage"), System.Drawing.Image)
    Me.imgExemptions.ShowFocusRect = True
    Me.imgExemptions.ShowSelectedState = False
    Me.imgExemptions.Size = New System.Drawing.Size(178, 32)
    Me.imgExemptions.SizeMode = CDBNETCL.ImageButton.ImageButtonSizeMode.AutoSize
    Me.imgExemptions.TabIndex = 4
    Me.imgExemptions.Tag = "SCXMBX"
    Me.imgExemptions.Text = "Exemptions"
    Me.imgExemptions.TextAlign = System.Drawing.ContentAlignment.MiddleRight
    Me.imgExemptions.TransparentColor = System.Drawing.Color.Transparent
    '
    'imgSessions
    '
    Me.imgSessions.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.imgSessions.BackColor = System.Drawing.Color.Transparent
    Me.imgSessions.DialogResult = System.Windows.Forms.DialogResult.OK
    Me.imgSessions.DrawBorder = False
    Me.imgSessions.Location = New System.Drawing.Point(3, 118)
    Me.imgSessions.Name = "imgSessions"
    Me.imgSessions.NormalImage = CType(resources.GetObject("imgSessions.NormalImage"), System.Drawing.Image)
    Me.imgSessions.ShowFocusRect = True
    Me.imgSessions.ShowSelectedState = False
    Me.imgSessions.Size = New System.Drawing.Size(178, 32)
    Me.imgSessions.SizeMode = CDBNETCL.ImageButton.ImageButtonSizeMode.AutoSize
    Me.imgSessions.TabIndex = 3
    Me.imgSessions.Tag = "SCXMBS"
    Me.imgSessions.Text = "Sessions"
    Me.imgSessions.TextAlign = System.Drawing.ContentAlignment.MiddleRight
    Me.imgSessions.TransparentColor = System.Drawing.Color.Transparent
    '
    'imgCentres
    '
    Me.imgCentres.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.imgCentres.BackColor = System.Drawing.Color.Transparent
    Me.imgCentres.DialogResult = System.Windows.Forms.DialogResult.OK
    Me.imgCentres.DrawBorder = False
    Me.imgCentres.Location = New System.Drawing.Point(3, 80)
    Me.imgCentres.Name = "imgCentres"
    Me.imgCentres.NormalImage = CType(resources.GetObject("imgCentres.NormalImage"), System.Drawing.Image)
    Me.imgCentres.ShowFocusRect = True
    Me.imgCentres.ShowSelectedState = False
    Me.imgCentres.Size = New System.Drawing.Size(178, 32)
    Me.imgCentres.SizeMode = CDBNETCL.ImageButton.ImageButtonSizeMode.AutoSize
    Me.imgCentres.TabIndex = 2
    Me.imgCentres.Tag = "SCXMBE"
    Me.imgCentres.Text = "Centres"
    Me.imgCentres.TextAlign = System.Drawing.ContentAlignment.MiddleRight
    Me.imgCentres.TransparentColor = System.Drawing.Color.Transparent
    '
    'imgPersonnel
    '
    Me.imgPersonnel.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.imgPersonnel.BackColor = System.Drawing.Color.Transparent
    Me.imgPersonnel.DialogResult = System.Windows.Forms.DialogResult.OK
    Me.imgPersonnel.DrawBorder = False
    Me.imgPersonnel.Location = New System.Drawing.Point(3, 42)
    Me.imgPersonnel.Name = "imgPersonnel"
    Me.imgPersonnel.NormalImage = CType(resources.GetObject("imgPersonnel.NormalImage"), System.Drawing.Image)
    Me.imgPersonnel.ShowFocusRect = True
    Me.imgPersonnel.ShowSelectedState = False
    Me.imgPersonnel.Size = New System.Drawing.Size(178, 32)
    Me.imgPersonnel.SizeMode = CDBNETCL.ImageButton.ImageButtonSizeMode.AutoSize
    Me.imgPersonnel.TabIndex = 1
    Me.imgPersonnel.Tag = "SCXMBP"
    Me.imgPersonnel.Text = "Personnel"
    Me.imgPersonnel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
    Me.imgPersonnel.TransparentColor = System.Drawing.Color.Transparent
    '
    'imgCourse
    '
    Me.imgCourse.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.imgCourse.BackColor = System.Drawing.Color.Transparent
    Me.imgCourse.DialogResult = System.Windows.Forms.DialogResult.OK
    Me.imgCourse.DrawBorder = False
    Me.imgCourse.Location = New System.Drawing.Point(3, 4)
    Me.imgCourse.Name = "imgCourse"
    Me.imgCourse.NormalImage = CType(resources.GetObject("imgCourse.NormalImage"), System.Drawing.Image)
    Me.imgCourse.ShowFocusRect = True
    Me.imgCourse.ShowSelectedState = False
    Me.imgCourse.Size = New System.Drawing.Size(178, 32)
    Me.imgCourse.SizeMode = CDBNETCL.ImageButton.ImageButtonSizeMode.AutoSize
    Me.imgCourse.TabIndex = 0
    Me.imgCourse.Tag = "SCXMBC"
    Me.imgCourse.Text = "Courses"
    Me.imgCourse.TextAlign = System.Drawing.ContentAlignment.MiddleRight
    Me.imgCourse.TransparentColor = System.Drawing.Color.Transparent
    '
    'tabMaster
    '
    Me.tabMaster.Appearance = System.Windows.Forms.TabAppearance.FlatButtons
    Me.tabMaster.Controls.Add(Me.tabMain)
    Me.tabMaster.Controls.Add(Me.tabCustomForm)
    Me.tabMaster.Dock = System.Windows.Forms.DockStyle.Fill
    Me.tabMaster.Location = New System.Drawing.Point(0, 0)
    Me.tabMaster.Margin = New System.Windows.Forms.Padding(0)
    Me.tabMaster.Name = "tabMaster"
    Me.tabMaster.Padding = New System.Drawing.Point(0, 0)
    Me.tabMaster.SelectedIndex = 0
    Me.tabMaster.Size = New System.Drawing.Size(595, 380)
    Me.tabMaster.TabIndex = 2
    '
    'tabMain
    '
    Me.tabMain.Controls.Add(Me.splRight)
    Me.tabMain.Location = New System.Drawing.Point(4, 25)
    Me.tabMain.Margin = New System.Windows.Forms.Padding(0)
    Me.tabMain.Name = "tabMain"
    Me.tabMain.Size = New System.Drawing.Size(587, 351)
    Me.tabMain.TabIndex = 1
    Me.tabMain.Text = "Main"
    Me.tabMain.UseVisualStyleBackColor = True
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
    Me.splRight.Panel1.Controls.Add(Me.UpperEditPanel)
    Me.splRight.Panel1.Controls.Add(Me.dgr)
    '
    'splRight.Panel2
    '
    Me.splRight.Panel2.Controls.Add(Me.dspTabGrid)
    Me.splRight.Panel2.Controls.Add(Me.splTab)
    Me.splRight.Panel2.Controls.Add(Me.dpl)
    Me.splRight.Panel2.Controls.Add(Me.dgrDetails)
    Me.splRight.Panel2.Controls.Add(Me.epl)
    Me.splRight.Panel2.Controls.Add(Me.selExams)
    Me.splRight.Panel2.Controls.Add(Me.bpl)
    Me.splRight.Size = New System.Drawing.Size(587, 351)
    Me.splRight.SplitterDistance = 114
    Me.splRight.TabIndex = 1
    '
    'UpperEditPanel
    '
    Me.UpperEditPanel.AddressChanged = False
    Me.UpperEditPanel.BackColor = System.Drawing.Color.Transparent
    Me.UpperEditPanel.DataChanged = False
    Me.UpperEditPanel.DefaultSaveFolder = ""
    Me.UpperEditPanel.Dock = System.Windows.Forms.DockStyle.Fill
    Me.UpperEditPanel.Location = New System.Drawing.Point(0, 0)
    Me.UpperEditPanel.Margin = New System.Windows.Forms.Padding(0)
    Me.UpperEditPanel.Name = "UpperEditPanel"
    Me.UpperEditPanel.Recipients = Nothing
    Me.UpperEditPanel.Size = New System.Drawing.Size(587, 114)
    Me.UpperEditPanel.SuppressDrawing = False
    Me.UpperEditPanel.TabIndex = 1
    Me.UpperEditPanel.TabSelectedIndex = 0
    Me.UpperEditPanel.Visible = False
    '
    'dgr
    '
    Me.dgr.AccessibleName = "Display Grid"
    Me.dgr.ActiveColumn = 0
    Me.dgr.AllowColumnResize = True
    Me.dgr.AllowSorting = True
    Me.dgr.AutoSetHeight = False
    Me.dgr.AutoSetRowHeight = False
    Me.dgr.DisplayTitle = Nothing
    Me.dgr.Dock = System.Windows.Forms.DockStyle.Fill
    Me.dgr.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!)
    Me.dgr.Location = New System.Drawing.Point(0, 0)
    Me.dgr.MaintenanceDesc = Nothing
    Me.dgr.Margin = New System.Windows.Forms.Padding(0)
    Me.dgr.MaxGridRows = 8
    Me.dgr.MultipleSelect = False
    Me.dgr.Name = "dgr"
    Me.dgr.RowCount = 10
    Me.dgr.ShowIfEmpty = False
    Me.dgr.Size = New System.Drawing.Size(587, 114)
    Me.dgr.SuppressHyperLinkFormat = False
    Me.dgr.TabIndex = 0
    '
    'dpl
    '
    Me.dpl.AccessibleName = "Display Panel"
    Me.dpl.BackColor = System.Drawing.Color.Transparent
    Me.dpl.DataSelectionType = CDBNETCL.CareNetServices.XMLContactDataSelectionTypes.xcdtNone
    Me.dpl.Dock = System.Windows.Forms.DockStyle.Top
    Me.dpl.Location = New System.Drawing.Point(0, 0)
    Me.dpl.Name = "dpl"
    Me.dpl.Size = New System.Drawing.Size(587, 80)
    Me.dpl.TabIndex = 4
    '
    'dgrDetails
    '
    Me.dgrDetails.AccessibleName = "Display Grid"
    Me.dgrDetails.ActiveColumn = 0
    Me.dgrDetails.AllowColumnResize = True
    Me.dgrDetails.AllowSorting = True
    Me.dgrDetails.AutoSetHeight = False
    Me.dgrDetails.AutoSetRowHeight = False
    Me.dgrDetails.DisplayTitle = Nothing
    Me.dgrDetails.Dock = System.Windows.Forms.DockStyle.Fill
    Me.dgrDetails.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!)
    Me.dgrDetails.Location = New System.Drawing.Point(0, 0)
    Me.dgrDetails.MaintenanceDesc = Nothing
    Me.dgrDetails.MaxGridRows = 8
    Me.dgrDetails.MultipleSelect = False
    Me.dgrDetails.Name = "dgrDetails"
    Me.dgrDetails.RowCount = 10
    Me.dgrDetails.ShowIfEmpty = False
    Me.dgrDetails.Size = New System.Drawing.Size(587, 194)
    Me.dgrDetails.SuppressHyperLinkFormat = False
    Me.dgrDetails.TabIndex = 2
    '
    'epl
    '
    Me.epl.AddressChanged = False
    Me.epl.BackColor = System.Drawing.Color.Transparent
    Me.epl.DataChanged = False
    Me.epl.DefaultSaveFolder = ""
    Me.epl.Dock = System.Windows.Forms.DockStyle.Fill
    Me.epl.Location = New System.Drawing.Point(0, 0)
    Me.epl.Name = "epl"
    Me.epl.Recipients = Nothing
    Me.epl.Size = New System.Drawing.Size(587, 194)
    Me.epl.SuppressDrawing = False
    Me.epl.TabIndex = 0
    Me.epl.TabSelectedIndex = 0
    '
    'selExams
    '
    Me.selExams.AutoSelectParents = False
    Me.selExams.BackColor = System.Drawing.Color.Transparent
    Me.selExams.Dock = System.Windows.Forms.DockStyle.Fill
    Me.selExams.ExamMaintenance = False
    Me.selExams.Location = New System.Drawing.Point(0, 0)
    Me.selExams.Name = "selExams"
    Me.selExams.Size = New System.Drawing.Size(587, 194)
    Me.selExams.TabIndex = 1
    Me.selExams.TreeContextMenu = Nothing
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdLink)
    Me.bpl.Controls.Add(Me.cmdAnalysis)
    Me.bpl.Controls.Add(Me.cmdSave)
    Me.bpl.Controls.Add(Me.cmdAllocate)
    Me.bpl.Controls.Add(Me.cmdUnallocate)
    Me.bpl.Controls.Add(Me.cmdSelectAll)
    Me.bpl.Controls.Add(Me.cmdUnSelectAll)
    Me.bpl.Controls.Add(Me.cmdNew)
    Me.bpl.Controls.Add(Me.cmdNewChild)
    Me.bpl.Controls.Add(Me.cmdDelete)
    Me.bpl.Controls.Add(Me.cmdClose)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bpl.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bpl.Location = New System.Drawing.Point(0, 194)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(587, 39)
    Me.bpl.TabIndex = 1
    '
    'cmdLink
    '
    Me.cmdLink.Location = New System.Drawing.Point(2, 6)
    Me.cmdLink.Name = "cmdLink"
    Me.cmdLink.Size = New System.Drawing.Size(52, 27)
    Me.cmdLink.TabIndex = 9
    Me.cmdLink.Text = "Link"
    Me.cmdLink.UseVisualStyleBackColor = True
    '
    'cmdAnalysis
    '
    Me.cmdAnalysis.Location = New System.Drawing.Point(55, 6)
    Me.cmdAnalysis.Name = "cmdAnalysis"
    Me.cmdAnalysis.Size = New System.Drawing.Size(52, 27)
    Me.cmdAnalysis.TabIndex = 8
    Me.cmdAnalysis.Text = "Analysis"
    Me.cmdAnalysis.UseVisualStyleBackColor = True
    '
    'cmdSave
    '
    Me.cmdSave.Location = New System.Drawing.Point(108, 6)
    Me.cmdSave.Name = "cmdSave"
    Me.cmdSave.Size = New System.Drawing.Size(52, 27)
    Me.cmdSave.TabIndex = 2
    Me.cmdSave.Text = "Save"
    Me.cmdSave.UseVisualStyleBackColor = True
    '
    'cmdAllocate
    '
    Me.cmdAllocate.Location = New System.Drawing.Point(161, 6)
    Me.cmdAllocate.Name = "cmdAllocate"
    Me.cmdAllocate.Size = New System.Drawing.Size(52, 27)
    Me.cmdAllocate.TabIndex = 7
    Me.cmdAllocate.Text = "Allocate"
    Me.cmdAllocate.UseVisualStyleBackColor = True
    '
    'cmdUnallocate
    '
    Me.cmdUnallocate.Location = New System.Drawing.Point(214, 6)
    Me.cmdUnallocate.Name = "cmdUnallocate"
    Me.cmdUnallocate.Size = New System.Drawing.Size(52, 27)
    Me.cmdUnallocate.TabIndex = 6
    Me.cmdUnallocate.Text = "UnAllocate"
    Me.cmdUnallocate.UseVisualStyleBackColor = True
    '
    'cmdSelectAll
    '
    Me.cmdSelectAll.Location = New System.Drawing.Point(267, 6)
    Me.cmdSelectAll.Name = "cmdSelectAll"
    Me.cmdSelectAll.Size = New System.Drawing.Size(52, 27)
    Me.cmdSelectAll.TabIndex = 5
    Me.cmdSelectAll.Text = "Select All"
    Me.cmdSelectAll.UseVisualStyleBackColor = True
    '
    'cmdUnSelectAll
    '
    Me.cmdUnSelectAll.Location = New System.Drawing.Point(320, 6)
    Me.cmdUnSelectAll.Name = "cmdUnSelectAll"
    Me.cmdUnSelectAll.Size = New System.Drawing.Size(52, 27)
    Me.cmdUnSelectAll.TabIndex = 4
    Me.cmdUnSelectAll.Text = "UnSelect All"
    Me.cmdUnSelectAll.UseVisualStyleBackColor = True
    '
    'cmdNew
    '
    Me.cmdNew.Location = New System.Drawing.Point(373, 6)
    Me.cmdNew.Name = "cmdNew"
    Me.cmdNew.Size = New System.Drawing.Size(52, 27)
    Me.cmdNew.TabIndex = 1
    Me.cmdNew.Text = "New"
    Me.cmdNew.UseVisualStyleBackColor = True
    '
    'cmdNewChild
    '
    Me.cmdNewChild.Location = New System.Drawing.Point(426, 6)
    Me.cmdNewChild.Name = "cmdNewChild"
    Me.cmdNewChild.Size = New System.Drawing.Size(52, 27)
    Me.cmdNewChild.TabIndex = 3
    Me.cmdNewChild.Text = "New Child"
    Me.cmdNewChild.UseVisualStyleBackColor = True
    '
    'cmdDelete
    '
    Me.cmdDelete.Location = New System.Drawing.Point(479, 6)
    Me.cmdDelete.Name = "cmdDelete"
    Me.cmdDelete.Size = New System.Drawing.Size(52, 27)
    Me.cmdDelete.TabIndex = 0
    Me.cmdDelete.Text = "Delete"
    Me.cmdDelete.UseVisualStyleBackColor = True
    '
    'cmdClose
    '
    Me.cmdClose.Location = New System.Drawing.Point(532, 6)
    Me.cmdClose.Name = "cmdClose"
    Me.cmdClose.Size = New System.Drawing.Size(52, 27)
    Me.cmdClose.TabIndex = 10
    Me.cmdClose.Text = "Close"
    Me.cmdClose.UseVisualStyleBackColor = True
    '
    'tabCustomForm
    '
    Me.tabCustomForm.Controls.Add(Me.pnlCustomFormPage)
    Me.tabCustomForm.Location = New System.Drawing.Point(4, 25)
    Me.tabCustomForm.Margin = New System.Windows.Forms.Padding(0)
    Me.tabCustomForm.Name = "tabCustomForm"
    Me.tabCustomForm.Size = New System.Drawing.Size(587, 351)
    Me.tabCustomForm.TabIndex = 2
    Me.tabCustomForm.Text = "Custom Form"
    Me.tabCustomForm.UseVisualStyleBackColor = True
    '
    'pnlCustomFormPage
    '
    Me.pnlCustomFormPage.DataContext = Nothing
    Me.pnlCustomFormPage.Dock = System.Windows.Forms.DockStyle.Fill
    Me.pnlCustomFormPage.IDataContext = Nothing
    Me.pnlCustomFormPage.Location = New System.Drawing.Point(0, 0)
    Me.pnlCustomFormPage.Margin = New System.Windows.Forms.Padding(0)
    Me.pnlCustomFormPage.Name = "pnlCustomFormPage"
    Me.pnlCustomFormPage.Size = New System.Drawing.Size(587, 351)
    Me.pnlCustomFormPage.TabIndex = 0
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
    'dgr0MenuStrip
    '
    Me.dgr0MenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.dgr0MenuNew, Me.dgr0MenuEdit})
    Me.dgr0MenuStrip.Name = "dgrMenuStrip"
    Me.dgr0MenuStrip.Size = New System.Drawing.Size(108, 48)
    '
    'dgr0MenuNew
    '
    Me.dgr0MenuNew.Name = "dgr0MenuNew"
    Me.dgr0MenuNew.Size = New System.Drawing.Size(107, 22)
    Me.dgr0MenuNew.Text = "&New..."
    '
    'dgr0MenuEdit
    '
    Me.dgr0MenuEdit.Name = "dgr0MenuEdit"
    Me.dgr0MenuEdit.Size = New System.Drawing.Size(107, 22)
    Me.dgr0MenuEdit.Text = "&Edit..."
    '
    'splTab
    '
    Me.splTab.Dock = System.Windows.Forms.DockStyle.Top
    Me.splTab.Location = New System.Drawing.Point(0, 80)
    Me.splTab.Name = "splTab"
    Me.splTab.Size = New System.Drawing.Size(587, 8)
    Me.splTab.TabIndex = 5
    Me.splTab.TabStop = False
    '
    'dspTabGrid
    '
    Me.dspTabGrid.Dock = System.Windows.Forms.DockStyle.Fill
    Me.dspTabGrid.Location = New System.Drawing.Point(0, 88)
    Me.dspTabGrid.Margin = New System.Windows.Forms.Padding(0)
    Me.dspTabGrid.Name = "dspTabGrid"
    Me.dspTabGrid.PanelVisible = False
    Me.dspTabGrid.Size = New System.Drawing.Size(587, 106)
    Me.dspTabGrid.TabDock = System.Windows.Forms.DockStyle.None
    Me.dspTabGrid.TabIndex = 6
    Me.dspTabGrid.TabVisible = False
    '
    'frmExams
    '
    Me.ClientSize = New System.Drawing.Size(783, 446)
    Me.Controls.Add(Me.splTop)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmExams"
    Me.splTop.Panel2.ResumeLayout(False)
    CType(Me.splTop, System.ComponentModel.ISupportInitialize).EndInit()
    Me.splTop.ResumeLayout(False)
    Me.splBottom.Panel1.ResumeLayout(False)
    Me.splBottom.Panel2.ResumeLayout(False)
    CType(Me.splBottom, System.ComponentModel.ISupportInitialize).EndInit()
    Me.splBottom.ResumeLayout(False)
    Me.pnlCommands.ResumeLayout(False)
    Me.tabMaster.ResumeLayout(False)
    Me.tabMain.ResumeLayout(False)
    Me.splRight.Panel1.ResumeLayout(False)
    Me.splRight.Panel2.ResumeLayout(False)
    CType(Me.splRight, System.ComponentModel.ISupportInitialize).EndInit()
    Me.splRight.ResumeLayout(False)
    Me.bpl.ResumeLayout(False)
    Me.tabCustomForm.ResumeLayout(False)
    Me.dgrMenuStrip.ResumeLayout(False)
    Me.dgr2MenuStrip.ResumeLayout(False)
    Me.dplMenuStrip.ResumeLayout(False)
    Me.dgr1MenuStrip.ResumeLayout(False)
    Me.dgr0MenuStrip.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents splTop As System.Windows.Forms.SplitContainer
  Friend WithEvents splBottom As System.Windows.Forms.SplitContainer
  Friend WithEvents splRight As System.Windows.Forms.SplitContainer
  Friend WithEvents dgr As CDBNETCL.DisplayGrid
  Friend WithEvents epl As CDBNETCL.EditPanel
  Friend WithEvents bpl As CDBNETCL.ButtonPanel
  Friend WithEvents cmdSave As System.Windows.Forms.Button
  Friend WithEvents cmdNew As System.Windows.Forms.Button
  Friend WithEvents cmdDelete As System.Windows.Forms.Button
  Friend WithEvents pnlCommands As CDBNETCL.PanelEx
  Friend WithEvents imgSessions As CDBNETCL.ImageButton
  Friend WithEvents imgCentres As CDBNETCL.ImageButton
  Friend WithEvents imgPersonnel As CDBNETCL.ImageButton
  Friend WithEvents imgCourse As CDBNETCL.ImageButton
  Friend WithEvents selExams As CDBNETCL.ExamSelector
  Friend WithEvents cmdNewChild As System.Windows.Forms.Button
  Friend WithEvents cboSessions As System.Windows.Forms.ComboBox
  Friend WithEvents imgExemptions As CDBNETCL.ImageButton
  Friend WithEvents cmdSelectAll As System.Windows.Forms.Button
  Friend WithEvents cmdUnSelectAll As System.Windows.Forms.Button
  Friend WithEvents dgrDetails As CDBNETCL.DisplayGrid
  Friend WithEvents cmdUnallocate As System.Windows.Forms.Button
  Friend WithEvents cmdAllocate As System.Windows.Forms.Button
  Friend WithEvents cmdLink As System.Windows.Forms.Button
  Friend WithEvents cmdAnalysis As System.Windows.Forms.Button
  Friend WithEvents cmdClose As System.Windows.Forms.Button
  Friend WithEvents UpperEditPanel As CDBNETCL.EditPanel
  Friend WithEvents dpl As CDBNETCL.DisplayPanel
  Friend WithEvents dgrMenuStrip As System.Windows.Forms.ContextMenuStrip
  Friend WithEvents dgrMenuNew As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents dgrMenuEdit As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents dgrMenuNewDocument As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents dgr2MenuStrip As System.Windows.Forms.ContextMenuStrip
  Friend WithEvents dgr2MenuNew As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents dgr2MenuEdit As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents dplMenuStrip As System.Windows.Forms.ContextMenuStrip
  Friend WithEvents dplMenuNew As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents dplMenuEdit As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents dplMenuCustomise As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents dplMenuRevert As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents dgr1MenuStrip As System.Windows.Forms.ContextMenuStrip
  Friend WithEvents dgr1MenuNew As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents dgr1MenuEdit As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents dgr0MenuStrip As System.Windows.Forms.ContextMenuStrip
  Friend WithEvents dgr0MenuNew As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents dgr0MenuEdit As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents tabMaster As System.Windows.Forms.TabControl
  Friend WithEvents tabMain As System.Windows.Forms.TabPage
  Friend WithEvents tabCustomForm As System.Windows.Forms.TabPage
  Friend WithEvents pnlCustomFormPage As CDBNETCLPages.GridPanelPage
  Friend WithEvents sel As CDBNETCL.ExamSelector
  Friend WithEvents splTab As System.Windows.Forms.Splitter
  Friend WithEvents dspTabGrid As CDBNETCL.DisplayTabSet

End Class
