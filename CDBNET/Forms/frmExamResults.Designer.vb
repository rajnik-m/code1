<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmExamResults
  Inherits CDBNETCL.MaintenanceParentForm

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
    Me.components = New System.ComponentModel.Container()
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmExamResults))
    Me.splMain = New System.Windows.Forms.SplitContainer()
    Me.bplTop = New CDBNETCL.ButtonPanel()
    Me.cmdSelect = New System.Windows.Forms.Button()
    Me.epl = New CDBNETCL.EditPanel()
    Me.splBottom = New System.Windows.Forms.SplitContainer()
    Me.splComponentSeparator = New System.Windows.Forms.Splitter()
    Me.dgrComponents = New CDBNETCL.DisplayGrid()
    Me.dgr = New CDBNETCL.DisplayGrid()
    Me.Condition = New System.Windows.Forms.TabControl()
    Me.TabPage1 = New System.Windows.Forms.TabPage()
    Me.pnlQuickEntry = New CDBNETCL.PanelEx()
    Me.lblInfo = New CDBNETCL.TransparentLabel()
    Me.txtResult = New CDBNETCL.TextLookupBox()
    Me.txtGrade = New CDBNETCL.TextLookupBox()
    Me.lblMark = New System.Windows.Forms.Label()
    Me.txtMark = New System.Windows.Forms.TextBox()
    Me.lblContactNumber = New System.Windows.Forms.Label()
    Me.txtContactNumber = New System.Windows.Forms.TextBox()
    Me.TabFindReplacePage = New System.Windows.Forms.TabPage()
    Me.TabFindReplaceHost = New System.Windows.Forms.TabControl()
    Me.TabFindReplaceMark = New System.Windows.Forms.TabPage()
    Me.lblReplaceWith = New System.Windows.Forms.Label()
    Me.lblFindWhat = New System.Windows.Forms.Label()
    Me.txtMarkReplace = New System.Windows.Forms.TextBox()
    Me.txtMarkFind = New System.Windows.Forms.TextBox()
    Me.TabFindReplaceGrade = New System.Windows.Forms.TabPage()
    Me.txtGradeReplace = New CDBNETCL.TextLookupBox()
    Me.Label1 = New System.Windows.Forms.Label()
    Me.Label2 = New System.Windows.Forms.Label()
    Me.txtGradeFind = New CDBNETCL.TextLookupBox()
    Me.TabFindReplaceResult = New System.Windows.Forms.TabPage()
    Me.Label3 = New System.Windows.Forms.Label()
    Me.Label4 = New System.Windows.Forms.Label()
    Me.txtResultFind = New CDBNETCL.TextLookupBox()
    Me.txtResultReplace = New CDBNETCL.TextLookupBox()
    Me.Panel1 = New System.Windows.Forms.Panel()
    Me.rdoResult = New System.Windows.Forms.RadioButton()
    Me.rdoGrade = New System.Windows.Forms.RadioButton()
    Me.rdoMark = New System.Windows.Forms.RadioButton()
    Me.cmdReplaceAll = New System.Windows.Forms.Button()
    Me.cmdReplace = New System.Windows.Forms.Button()
    Me.cmdFindNext = New System.Windows.Forms.Button()
    Me.TabPage3 = New System.Windows.Forms.TabPage()
    Me.PanelEx2 = New CDBNETCL.PanelEx()
    Me.txtChangeReason = New CDBNETCL.TextLookupBox()
    Me.lblChangeReason = New System.Windows.Forms.Label()
    Me.bplBottom = New CDBNETCL.ButtonPanel()
    Me.cmdSave = New System.Windows.Forms.Button()
    Me.cmdOK = New System.Windows.Forms.Button()
    Me.cmdCancel = New System.Windows.Forms.Button()
    CType(Me.splMain, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.splMain.Panel1.SuspendLayout()
    Me.splMain.Panel2.SuspendLayout()
    Me.splMain.SuspendLayout()
    Me.bplTop.SuspendLayout()
    CType(Me.splBottom, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.splBottom.Panel1.SuspendLayout()
    Me.splBottom.Panel2.SuspendLayout()
    Me.splBottom.SuspendLayout()
    Me.Condition.SuspendLayout()
    Me.TabPage1.SuspendLayout()
    Me.pnlQuickEntry.SuspendLayout()
    Me.TabFindReplacePage.SuspendLayout()
    Me.TabFindReplaceHost.SuspendLayout()
    Me.TabFindReplaceMark.SuspendLayout()
    Me.TabFindReplaceGrade.SuspendLayout()
    Me.TabFindReplaceResult.SuspendLayout()
    Me.Panel1.SuspendLayout()
    Me.TabPage3.SuspendLayout()
    Me.PanelEx2.SuspendLayout()
    Me.bplBottom.SuspendLayout()
    Me.SuspendLayout()
    '
    'splMain
    '
    Me.splMain.Dock = System.Windows.Forms.DockStyle.Fill
    Me.splMain.Location = New System.Drawing.Point(0, 0)
    Me.splMain.Name = "splMain"
    Me.splMain.Orientation = System.Windows.Forms.Orientation.Horizontal
    '
    'splMain.Panel1
    '
    Me.splMain.Panel1.Controls.Add(Me.bplTop)
    Me.splMain.Panel1.Controls.Add(Me.epl)
    '
    'splMain.Panel2
    '
    Me.splMain.Panel2.Controls.Add(Me.splBottom)
    Me.splMain.Size = New System.Drawing.Size(732, 573)
    Me.splMain.SplitterDistance = 156
    Me.splMain.TabIndex = 0
    '
    'bplTop
    '
    Me.bplTop.Controls.Add(Me.cmdSelect)
    Me.bplTop.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bplTop.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bplTop.Location = New System.Drawing.Point(0, 117)
    Me.bplTop.Name = "bplTop"
    Me.bplTop.Size = New System.Drawing.Size(732, 39)
    Me.bplTop.TabIndex = 1
    '
    'cmdSelect
    '
    Me.cmdSelect.Location = New System.Drawing.Point(318, 6)
    Me.cmdSelect.Name = "cmdSelect"
    Me.cmdSelect.Size = New System.Drawing.Size(96, 27)
    Me.cmdSelect.TabIndex = 0
    Me.cmdSelect.Text = "Select"
    Me.cmdSelect.UseVisualStyleBackColor = True
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
    Me.epl.Size = New System.Drawing.Size(732, 156)
    Me.epl.SuppressDrawing = False
    Me.epl.TabIndex = 0
    Me.epl.TabSelectedIndex = 0
    '
    'splBottom
    '
    Me.splBottom.Dock = System.Windows.Forms.DockStyle.Fill
    Me.splBottom.Location = New System.Drawing.Point(0, 0)
    Me.splBottom.Name = "splBottom"
    Me.splBottom.Orientation = System.Windows.Forms.Orientation.Horizontal
    '
    'splBottom.Panel1
    '
    Me.splBottom.Panel1.Controls.Add(Me.splComponentSeparator)
    Me.splBottom.Panel1.Controls.Add(Me.dgrComponents)
    Me.splBottom.Panel1.Controls.Add(Me.dgr)
    '
    'splBottom.Panel2
    '
    Me.splBottom.Panel2.Controls.Add(Me.Condition)
    Me.splBottom.Size = New System.Drawing.Size(732, 413)
    Me.splBottom.SplitterDistance = 250
    Me.splBottom.TabIndex = 5
    '
    'splComponentSeparator
    '
    Me.splComponentSeparator.BackColor = System.Drawing.Color.WhiteSmoke
    Me.splComponentSeparator.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.splComponentSeparator.Location = New System.Drawing.Point(0, 140)
    Me.splComponentSeparator.Name = "splComponentSeparator"
    Me.splComponentSeparator.Size = New System.Drawing.Size(732, 10)
    Me.splComponentSeparator.TabIndex = 6
    Me.splComponentSeparator.TabStop = False
    '
    'dgrComponents
    '
    Me.dgrComponents.AccessibleName = "Display Grid"
    Me.dgrComponents.ActiveColumn = 0
    Me.dgrComponents.AllowColumnResize = True
    Me.dgrComponents.AllowSorting = True
    Me.dgrComponents.AutoSetHeight = False
    Me.dgrComponents.AutoSetRowHeight = False
    Me.dgrComponents.DisplayTitle = Nothing
    Me.dgrComponents.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.dgrComponents.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!)
    Me.dgrComponents.Location = New System.Drawing.Point(0, 150)
    Me.dgrComponents.MaintenanceDesc = Nothing
    Me.dgrComponents.MaxGridRows = 8
    Me.dgrComponents.MinimumSize = New System.Drawing.Size(0, 100)
    Me.dgrComponents.MultipleSelect = False
    Me.dgrComponents.Name = "dgrComponents"
    Me.dgrComponents.RowCount = 10
    Me.dgrComponents.ShowIfEmpty = False
    Me.dgrComponents.Size = New System.Drawing.Size(732, 100)
    Me.dgrComponents.SuppressHyperLinkFormat = False
    Me.dgrComponents.TabIndex = 5
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
    Me.dgr.MaxGridRows = 8
    Me.dgr.MinimumSize = New System.Drawing.Size(0, 100)
    Me.dgr.MultipleSelect = False
    Me.dgr.Name = "dgr"
    Me.dgr.RowCount = 10
    Me.dgr.ShowIfEmpty = False
    Me.dgr.Size = New System.Drawing.Size(732, 250)
    Me.dgr.SuppressHyperLinkFormat = False
    Me.dgr.TabIndex = 1
    '
    'Condition
    '
    Me.Condition.Controls.Add(Me.TabPage1)
    Me.Condition.Controls.Add(Me.TabFindReplacePage)
    Me.Condition.Controls.Add(Me.TabPage3)
    Me.Condition.Dock = System.Windows.Forms.DockStyle.Fill
    Me.Condition.Location = New System.Drawing.Point(0, 0)
    Me.Condition.Name = "Condition"
    Me.Condition.SelectedIndex = 0
    Me.Condition.Size = New System.Drawing.Size(732, 159)
    Me.Condition.TabIndex = 6
    '
    'TabPage1
    '
    Me.TabPage1.Controls.Add(Me.pnlQuickEntry)
    Me.TabPage1.Location = New System.Drawing.Point(4, 25)
    Me.TabPage1.Name = "TabPage1"
    Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
    Me.TabPage1.Size = New System.Drawing.Size(724, 93)
    Me.TabPage1.TabIndex = 0
    Me.TabPage1.Text = "Details"
    Me.TabPage1.UseVisualStyleBackColor = True
    '
    'pnlQuickEntry
    '
    Me.pnlQuickEntry.BackColor = System.Drawing.Color.Transparent
    Me.pnlQuickEntry.Controls.Add(Me.lblInfo)
    Me.pnlQuickEntry.Controls.Add(Me.txtResult)
    Me.pnlQuickEntry.Controls.Add(Me.txtGrade)
    Me.pnlQuickEntry.Controls.Add(Me.lblMark)
    Me.pnlQuickEntry.Controls.Add(Me.txtMark)
    Me.pnlQuickEntry.Controls.Add(Me.lblContactNumber)
    Me.pnlQuickEntry.Controls.Add(Me.txtContactNumber)
    Me.pnlQuickEntry.Dock = System.Windows.Forms.DockStyle.Fill
    Me.pnlQuickEntry.Location = New System.Drawing.Point(3, 3)
    Me.pnlQuickEntry.Name = "pnlQuickEntry"
    Me.pnlQuickEntry.Size = New System.Drawing.Size(718, 87)
    Me.pnlQuickEntry.TabIndex = 6
    '
    'lblInfo
    '
    Me.lblInfo.AutoSize = True
    Me.lblInfo.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.lblInfo.Location = New System.Drawing.Point(0, 53)
    Me.lblInfo.Multiline = False
    Me.lblInfo.Name = "lblInfo"
    Me.lblInfo.Size = New System.Drawing.Size(751, 17)
    Me.lblInfo.TabIndex = 49
    Me.lblInfo.Text = "Grades, Marks and Results can only be updated from the grid, when multiple result" & _
    " entry configuration is switched on. "
    Me.lblInfo.Visible = False
    '
    'txtResult
    '
    Me.txtResult.ActiveOnly = False
    Me.txtResult.BackColor = System.Drawing.SystemColors.Control
    Me.txtResult.CustomFormNumber = 0
    Me.txtResult.Description = ""
    Me.txtResult.EnabledProperty = True
    Me.txtResult.ExamCentreId = 0
    Me.txtResult.ExamUnitLinkId = 0
    Me.txtResult.HasDependancies = False
    Me.txtResult.IsDesign = False
    Me.txtResult.Location = New System.Drawing.Point(238, 11)
    Me.txtResult.MaxLength = 32767
    Me.txtResult.MultipleValuesSupported = False
    Me.txtResult.Name = "txtResult"
    Me.txtResult.OriginalText = Nothing
    Me.txtResult.PreventHistoricalSelection = False
    Me.txtResult.ReadOnlyProperty = False
    Me.txtResult.Size = New System.Drawing.Size(410, 24)
    Me.txtResult.TabIndex = 48
    Me.txtResult.TextReadOnly = False
    Me.txtResult.TotalWidth = 408
    Me.txtResult.ValidationRequired = True
    Me.txtResult.WarningMessage = Nothing
    '
    'txtGrade
    '
    Me.txtGrade.ActiveOnly = False
    Me.txtGrade.BackColor = System.Drawing.SystemColors.Control
    Me.txtGrade.CustomFormNumber = 0
    Me.txtGrade.Description = ""
    Me.txtGrade.EnabledProperty = True
    Me.txtGrade.ExamCentreId = 0
    Me.txtGrade.ExamUnitLinkId = 0
    Me.txtGrade.HasDependancies = False
    Me.txtGrade.IsDesign = False
    Me.txtGrade.Location = New System.Drawing.Point(238, 11)
    Me.txtGrade.MaxLength = 32767
    Me.txtGrade.MultipleValuesSupported = False
    Me.txtGrade.Name = "txtGrade"
    Me.txtGrade.OriginalText = Nothing
    Me.txtGrade.PreventHistoricalSelection = False
    Me.txtGrade.ReadOnlyProperty = False
    Me.txtGrade.Size = New System.Drawing.Size(410, 24)
    Me.txtGrade.TabIndex = 47
    Me.txtGrade.TextReadOnly = False
    Me.txtGrade.TotalWidth = 408
    Me.txtGrade.ValidationRequired = True
    Me.txtGrade.WarningMessage = Nothing
    '
    'lblMark
    '
    Me.lblMark.AutoSize = True
    Me.lblMark.Location = New System.Drawing.Point(186, 14)
    Me.lblMark.Name = "lblMark"
    Me.lblMark.Size = New System.Drawing.Size(48, 17)
    Me.lblMark.TabIndex = 7
    Me.lblMark.Text = "Grade"
    '
    'txtMark
    '
    Me.txtMark.Location = New System.Drawing.Point(288, 11)
    Me.txtMark.Name = "txtMark"
    Me.txtMark.Size = New System.Drawing.Size(105, 22)
    Me.txtMark.TabIndex = 8
    '
    'lblContactNumber
    '
    Me.lblContactNumber.AutoSize = True
    Me.lblContactNumber.Location = New System.Drawing.Point(2, 14)
    Me.lblContactNumber.Name = "lblContactNumber"
    Me.lblContactNumber.Size = New System.Drawing.Size(56, 17)
    Me.lblContactNumber.TabIndex = 5
    Me.lblContactNumber.Text = "Contact"
    '
    'txtContactNumber
    '
    Me.txtContactNumber.Location = New System.Drawing.Point(62, 11)
    Me.txtContactNumber.Name = "txtContactNumber"
    Me.txtContactNumber.Size = New System.Drawing.Size(105, 22)
    Me.txtContactNumber.TabIndex = 6
    '
    'TabFindReplacePage
    '
    Me.TabFindReplacePage.Controls.Add(Me.TabFindReplaceHost)
    Me.TabFindReplacePage.Controls.Add(Me.Panel1)
    Me.TabFindReplacePage.Location = New System.Drawing.Point(4, 25)
    Me.TabFindReplacePage.Name = "TabFindReplacePage"
    Me.TabFindReplacePage.Padding = New System.Windows.Forms.Padding(3)
    Me.TabFindReplacePage.Size = New System.Drawing.Size(724, 130)
    Me.TabFindReplacePage.TabIndex = 3
    Me.TabFindReplacePage.Text = "Find & Replace"
    Me.TabFindReplacePage.UseVisualStyleBackColor = True
    '
    'TabFindReplaceHost
    '
    Me.TabFindReplaceHost.Controls.Add(Me.TabFindReplaceMark)
    Me.TabFindReplaceHost.Controls.Add(Me.TabFindReplaceGrade)
    Me.TabFindReplaceHost.Controls.Add(Me.TabFindReplaceResult)
    Me.TabFindReplaceHost.Dock = System.Windows.Forms.DockStyle.Fill
    Me.TabFindReplaceHost.ItemSize = New System.Drawing.Size(10, 12)
    Me.TabFindReplaceHost.Location = New System.Drawing.Point(3, 3)
    Me.TabFindReplaceHost.Name = "TabFindReplaceHost"
    Me.TabFindReplaceHost.SelectedIndex = 0
    Me.TabFindReplaceHost.Size = New System.Drawing.Size(518, 124)
    Me.TabFindReplaceHost.TabIndex = 0
    '
    'TabFindReplaceMark
    '
    Me.TabFindReplaceMark.Controls.Add(Me.lblReplaceWith)
    Me.TabFindReplaceMark.Controls.Add(Me.lblFindWhat)
    Me.TabFindReplaceMark.Controls.Add(Me.txtMarkReplace)
    Me.TabFindReplaceMark.Controls.Add(Me.txtMarkFind)
    Me.TabFindReplaceMark.Location = New System.Drawing.Point(4, 16)
    Me.TabFindReplaceMark.Name = "TabFindReplaceMark"
    Me.TabFindReplaceMark.Padding = New System.Windows.Forms.Padding(3)
    Me.TabFindReplaceMark.Size = New System.Drawing.Size(510, 104)
    Me.TabFindReplaceMark.TabIndex = 0
    Me.TabFindReplaceMark.Text = "Marks"
    Me.TabFindReplaceMark.UseVisualStyleBackColor = True
    '
    'lblReplaceWith
    '
    Me.lblReplaceWith.AutoSize = True
    Me.lblReplaceWith.Location = New System.Drawing.Point(5, 45)
    Me.lblReplaceWith.Name = "lblReplaceWith"
    Me.lblReplaceWith.Size = New System.Drawing.Size(92, 17)
    Me.lblReplaceWith.TabIndex = 61
    Me.lblReplaceWith.Text = "Replace With"
    '
    'lblFindWhat
    '
    Me.lblFindWhat.AutoSize = True
    Me.lblFindWhat.Location = New System.Drawing.Point(5, 8)
    Me.lblFindWhat.Name = "lblFindWhat"
    Me.lblFindWhat.Size = New System.Drawing.Size(72, 17)
    Me.lblFindWhat.TabIndex = 60
    Me.lblFindWhat.Text = "Find What"
    '
    'txtMarkReplace
    '
    Me.txtMarkReplace.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.txtMarkReplace.Location = New System.Drawing.Point(101, 44)
    Me.txtMarkReplace.Name = "txtMarkReplace"
    Me.txtMarkReplace.Size = New System.Drawing.Size(403, 22)
    Me.txtMarkReplace.TabIndex = 59
    '
    'txtMarkFind
    '
    Me.txtMarkFind.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.txtMarkFind.Location = New System.Drawing.Point(101, 8)
    Me.txtMarkFind.Name = "txtMarkFind"
    Me.txtMarkFind.Size = New System.Drawing.Size(403, 22)
    Me.txtMarkFind.TabIndex = 58
    '
    'TabFindReplaceGrade
    '
    Me.TabFindReplaceGrade.Controls.Add(Me.txtGradeReplace)
    Me.TabFindReplaceGrade.Controls.Add(Me.Label1)
    Me.TabFindReplaceGrade.Controls.Add(Me.Label2)
    Me.TabFindReplaceGrade.Controls.Add(Me.txtGradeFind)
    Me.TabFindReplaceGrade.Location = New System.Drawing.Point(4, 16)
    Me.TabFindReplaceGrade.Name = "TabFindReplaceGrade"
    Me.TabFindReplaceGrade.Padding = New System.Windows.Forms.Padding(3)
    Me.TabFindReplaceGrade.Size = New System.Drawing.Size(510, 67)
    Me.TabFindReplaceGrade.TabIndex = 1
    Me.TabFindReplaceGrade.Text = "Grade"
    Me.TabFindReplaceGrade.UseVisualStyleBackColor = True
    '
    'txtGradeReplace
    '
    Me.txtGradeReplace.ActiveOnly = False
    Me.txtGradeReplace.BackColor = System.Drawing.SystemColors.Control
    Me.txtGradeReplace.CustomFormNumber = 0
    Me.txtGradeReplace.Description = ""
    Me.txtGradeReplace.EnabledProperty = True
    Me.txtGradeReplace.ExamCentreId = 0
    Me.txtGradeReplace.ExamUnitLinkId = 0
    Me.txtGradeReplace.HasDependancies = False
    Me.txtGradeReplace.IsDesign = False
    Me.txtGradeReplace.Location = New System.Drawing.Point(101, 45)
    Me.txtGradeReplace.MaxLength = 32767
    Me.txtGradeReplace.MultipleValuesSupported = False
    Me.txtGradeReplace.Name = "txtGradeReplace"
    Me.txtGradeReplace.OriginalText = Nothing
    Me.txtGradeReplace.PreventHistoricalSelection = False
    Me.txtGradeReplace.ReadOnlyProperty = False
    Me.txtGradeReplace.Size = New System.Drawing.Size(403, 24)
    Me.txtGradeReplace.TabIndex = 64
    Me.txtGradeReplace.TextReadOnly = False
    Me.txtGradeReplace.TotalWidth = 408
    Me.txtGradeReplace.ValidationRequired = True
    Me.txtGradeReplace.WarningMessage = Nothing
    '
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.Location = New System.Drawing.Point(5, 45)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(92, 17)
    Me.Label1.TabIndex = 63
    Me.Label1.Text = "Replace With"
    '
    'Label2
    '
    Me.Label2.AutoSize = True
    Me.Label2.Location = New System.Drawing.Point(5, 8)
    Me.Label2.Name = "Label2"
    Me.Label2.Size = New System.Drawing.Size(72, 17)
    Me.Label2.TabIndex = 62
    Me.Label2.Text = "Find What"
    '
    'txtGradeFind
    '
    Me.txtGradeFind.ActiveOnly = False
    Me.txtGradeFind.BackColor = System.Drawing.SystemColors.Control
    Me.txtGradeFind.CustomFormNumber = 0
    Me.txtGradeFind.Description = ""
    Me.txtGradeFind.EnabledProperty = True
    Me.txtGradeFind.ExamCentreId = 0
    Me.txtGradeFind.ExamUnitLinkId = 0
    Me.txtGradeFind.HasDependancies = False
    Me.txtGradeFind.IsDesign = False
    Me.txtGradeFind.Location = New System.Drawing.Point(101, 8)
    Me.txtGradeFind.MaxLength = 32767
    Me.txtGradeFind.MultipleValuesSupported = False
    Me.txtGradeFind.Name = "txtGradeFind"
    Me.txtGradeFind.OriginalText = Nothing
    Me.txtGradeFind.PreventHistoricalSelection = False
    Me.txtGradeFind.ReadOnlyProperty = False
    Me.txtGradeFind.Size = New System.Drawing.Size(403, 24)
    Me.txtGradeFind.TabIndex = 50
    Me.txtGradeFind.TextReadOnly = False
    Me.txtGradeFind.TotalWidth = 408
    Me.txtGradeFind.ValidationRequired = True
    Me.txtGradeFind.WarningMessage = Nothing
    '
    'TabFindReplaceResult
    '
    Me.TabFindReplaceResult.Controls.Add(Me.Label3)
    Me.TabFindReplaceResult.Controls.Add(Me.Label4)
    Me.TabFindReplaceResult.Controls.Add(Me.txtResultFind)
    Me.TabFindReplaceResult.Controls.Add(Me.txtResultReplace)
    Me.TabFindReplaceResult.Location = New System.Drawing.Point(4, 16)
    Me.TabFindReplaceResult.Name = "TabFindReplaceResult"
    Me.TabFindReplaceResult.Padding = New System.Windows.Forms.Padding(3)
    Me.TabFindReplaceResult.Size = New System.Drawing.Size(510, 67)
    Me.TabFindReplaceResult.TabIndex = 2
    Me.TabFindReplaceResult.Text = "Result"
    Me.TabFindReplaceResult.UseVisualStyleBackColor = True
    '
    'Label3
    '
    Me.Label3.AutoSize = True
    Me.Label3.Location = New System.Drawing.Point(5, 45)
    Me.Label3.Name = "Label3"
    Me.Label3.Size = New System.Drawing.Size(92, 17)
    Me.Label3.TabIndex = 65
    Me.Label3.Text = "Replace With"
    '
    'Label4
    '
    Me.Label4.AutoSize = True
    Me.Label4.Location = New System.Drawing.Point(5, 8)
    Me.Label4.Name = "Label4"
    Me.Label4.Size = New System.Drawing.Size(72, 17)
    Me.Label4.TabIndex = 64
    Me.Label4.Text = "Find What"
    '
    'txtResultFind
    '
    Me.txtResultFind.ActiveOnly = False
    Me.txtResultFind.BackColor = System.Drawing.SystemColors.Control
    Me.txtResultFind.CustomFormNumber = 0
    Me.txtResultFind.Description = ""
    Me.txtResultFind.EnabledProperty = True
    Me.txtResultFind.ExamCentreId = 0
    Me.txtResultFind.ExamUnitLinkId = 0
    Me.txtResultFind.HasDependancies = False
    Me.txtResultFind.IsDesign = False
    Me.txtResultFind.Location = New System.Drawing.Point(101, 8)
    Me.txtResultFind.MaxLength = 32767
    Me.txtResultFind.MultipleValuesSupported = False
    Me.txtResultFind.Name = "txtResultFind"
    Me.txtResultFind.OriginalText = Nothing
    Me.txtResultFind.PreventHistoricalSelection = False
    Me.txtResultFind.ReadOnlyProperty = False
    Me.txtResultFind.Size = New System.Drawing.Size(403, 24)
    Me.txtResultFind.TabIndex = 53
    Me.txtResultFind.TextReadOnly = False
    Me.txtResultFind.TotalWidth = 408
    Me.txtResultFind.ValidationRequired = True
    Me.txtResultFind.WarningMessage = Nothing
    '
    'txtResultReplace
    '
    Me.txtResultReplace.ActiveOnly = False
    Me.txtResultReplace.BackColor = System.Drawing.SystemColors.Control
    Me.txtResultReplace.CustomFormNumber = 0
    Me.txtResultReplace.Description = ""
    Me.txtResultReplace.EnabledProperty = True
    Me.txtResultReplace.ExamCentreId = 0
    Me.txtResultReplace.ExamUnitLinkId = 0
    Me.txtResultReplace.HasDependancies = False
    Me.txtResultReplace.IsDesign = False
    Me.txtResultReplace.Location = New System.Drawing.Point(101, 45)
    Me.txtResultReplace.MaxLength = 32767
    Me.txtResultReplace.MultipleValuesSupported = False
    Me.txtResultReplace.Name = "txtResultReplace"
    Me.txtResultReplace.OriginalText = Nothing
    Me.txtResultReplace.PreventHistoricalSelection = False
    Me.txtResultReplace.ReadOnlyProperty = False
    Me.txtResultReplace.Size = New System.Drawing.Size(403, 24)
    Me.txtResultReplace.TabIndex = 54
    Me.txtResultReplace.TextReadOnly = False
    Me.txtResultReplace.TotalWidth = 408
    Me.txtResultReplace.ValidationRequired = True
    Me.txtResultReplace.WarningMessage = Nothing
    '
    'Panel1
    '
    Me.Panel1.Controls.Add(Me.rdoResult)
    Me.Panel1.Controls.Add(Me.rdoGrade)
    Me.Panel1.Controls.Add(Me.rdoMark)
    Me.Panel1.Controls.Add(Me.cmdReplaceAll)
    Me.Panel1.Controls.Add(Me.cmdReplace)
    Me.Panel1.Controls.Add(Me.cmdFindNext)
    Me.Panel1.Dock = System.Windows.Forms.DockStyle.Right
    Me.Panel1.Location = New System.Drawing.Point(521, 3)
    Me.Panel1.Name = "Panel1"
    Me.Panel1.Size = New System.Drawing.Size(200, 124)
    Me.Panel1.TabIndex = 1
    '
    'rdoResult
    '
    Me.rdoResult.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.rdoResult.AutoSize = True
    Me.rdoResult.Location = New System.Drawing.Point(116, 60)
    Me.rdoResult.Name = "rdoResult"
    Me.rdoResult.Size = New System.Drawing.Size(69, 21)
    Me.rdoResult.TabIndex = 58
    Me.rdoResult.Text = "Result"
    Me.rdoResult.UseVisualStyleBackColor = True
    '
    'rdoGrade
    '
    Me.rdoGrade.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.rdoGrade.AutoSize = True
    Me.rdoGrade.Location = New System.Drawing.Point(116, 32)
    Me.rdoGrade.Name = "rdoGrade"
    Me.rdoGrade.Size = New System.Drawing.Size(69, 21)
    Me.rdoGrade.TabIndex = 57
    Me.rdoGrade.Text = "Grade"
    Me.rdoGrade.UseVisualStyleBackColor = True
    '
    'rdoMark
    '
    Me.rdoMark.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.rdoMark.AutoSize = True
    Me.rdoMark.Checked = True
    Me.rdoMark.Location = New System.Drawing.Point(116, 6)
    Me.rdoMark.Name = "rdoMark"
    Me.rdoMark.Size = New System.Drawing.Size(60, 21)
    Me.rdoMark.TabIndex = 56
    Me.rdoMark.TabStop = True
    Me.rdoMark.Text = "Mark"
    Me.rdoMark.UseVisualStyleBackColor = True
    '
    'cmdReplaceAll
    '
    Me.cmdReplaceAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdReplaceAll.Location = New System.Drawing.Point(6, 58)
    Me.cmdReplaceAll.Name = "cmdReplaceAll"
    Me.cmdReplaceAll.Size = New System.Drawing.Size(86, 22)
    Me.cmdReplaceAll.TabIndex = 7
    Me.cmdReplaceAll.Text = "Replace All"
    Me.cmdReplaceAll.UseVisualStyleBackColor = True
    '
    'cmdReplace
    '
    Me.cmdReplace.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdReplace.Location = New System.Drawing.Point(6, 32)
    Me.cmdReplace.Name = "cmdReplace"
    Me.cmdReplace.Size = New System.Drawing.Size(86, 22)
    Me.cmdReplace.TabIndex = 6
    Me.cmdReplace.Text = "Replace"
    Me.cmdReplace.UseVisualStyleBackColor = True
    '
    'cmdFindNext
    '
    Me.cmdFindNext.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdFindNext.Location = New System.Drawing.Point(6, 5)
    Me.cmdFindNext.Name = "cmdFindNext"
    Me.cmdFindNext.Size = New System.Drawing.Size(86, 22)
    Me.cmdFindNext.TabIndex = 5
    Me.cmdFindNext.Text = "Find Next"
    Me.cmdFindNext.UseVisualStyleBackColor = True
    '
    'TabPage3
    '
    Me.TabPage3.Controls.Add(Me.PanelEx2)
    Me.TabPage3.Location = New System.Drawing.Point(4, 25)
    Me.TabPage3.Name = "TabPage3"
    Me.TabPage3.Padding = New System.Windows.Forms.Padding(3)
    Me.TabPage3.Size = New System.Drawing.Size(724, 93)
    Me.TabPage3.TabIndex = 2
    Me.TabPage3.Text = "Change Reason"
    Me.TabPage3.UseVisualStyleBackColor = True
    '
    'PanelEx2
    '
    Me.PanelEx2.BackColor = System.Drawing.Color.Transparent
    Me.PanelEx2.Controls.Add(Me.txtChangeReason)
    Me.PanelEx2.Controls.Add(Me.lblChangeReason)
    Me.PanelEx2.Dock = System.Windows.Forms.DockStyle.Fill
    Me.PanelEx2.Location = New System.Drawing.Point(3, 3)
    Me.PanelEx2.Name = "PanelEx2"
    Me.PanelEx2.Size = New System.Drawing.Size(718, 87)
    Me.PanelEx2.TabIndex = 7
    '
    'txtChangeReason
    '
    Me.txtChangeReason.ActiveOnly = False
    Me.txtChangeReason.BackColor = System.Drawing.SystemColors.Control
    Me.txtChangeReason.CustomFormNumber = 0
    Me.txtChangeReason.Description = ""
    Me.txtChangeReason.EnabledProperty = True
    Me.txtChangeReason.ExamCentreId = 0
    Me.txtChangeReason.ExamUnitLinkId = 0
    Me.txtChangeReason.HasDependancies = False
    Me.txtChangeReason.IsDesign = False
    Me.txtChangeReason.Location = New System.Drawing.Point(131, 19)
    Me.txtChangeReason.MaxLength = 32767
    Me.txtChangeReason.MultipleValuesSupported = False
    Me.txtChangeReason.Name = "txtChangeReason"
    Me.txtChangeReason.OriginalText = Nothing
    Me.txtChangeReason.PreventHistoricalSelection = False
    Me.txtChangeReason.ReadOnlyProperty = False
    Me.txtChangeReason.Size = New System.Drawing.Size(410, 24)
    Me.txtChangeReason.TabIndex = 48
    Me.txtChangeReason.TextReadOnly = False
    Me.txtChangeReason.TotalWidth = 408
    Me.txtChangeReason.ValidationRequired = True
    Me.txtChangeReason.WarningMessage = Nothing
    '
    'lblChangeReason
    '
    Me.lblChangeReason.AutoSize = True
    Me.lblChangeReason.Location = New System.Drawing.Point(16, 22)
    Me.lblChangeReason.Name = "lblChangeReason"
    Me.lblChangeReason.Size = New System.Drawing.Size(110, 17)
    Me.lblChangeReason.TabIndex = 7
    Me.lblChangeReason.Text = "Change Reason"
    '
    'bplBottom
    '
    Me.bplBottom.Controls.Add(Me.cmdSave)
    Me.bplBottom.Controls.Add(Me.cmdOK)
    Me.bplBottom.Controls.Add(Me.cmdCancel)
    Me.bplBottom.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bplBottom.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bplBottom.Location = New System.Drawing.Point(0, 573)
    Me.bplBottom.Name = "bplBottom"
    Me.bplBottom.Size = New System.Drawing.Size(732, 39)
    Me.bplBottom.TabIndex = 4
    '
    'cmdSave
    '
    Me.cmdSave.Location = New System.Drawing.Point(207, 6)
    Me.cmdSave.Name = "cmdSave"
    Me.cmdSave.Size = New System.Drawing.Size(96, 27)
    Me.cmdSave.TabIndex = 0
    Me.cmdSave.Text = "Save"
    Me.cmdSave.UseVisualStyleBackColor = True
    '
    'cmdOK
    '
    Me.cmdOK.Location = New System.Drawing.Point(318, 6)
    Me.cmdOK.Name = "cmdOK"
    Me.cmdOK.Size = New System.Drawing.Size(96, 27)
    Me.cmdOK.TabIndex = 0
    Me.cmdOK.Text = "OK"
    Me.cmdOK.UseVisualStyleBackColor = True
    '
    'cmdCancel
    '
    Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdCancel.Location = New System.Drawing.Point(429, 6)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(96, 27)
    Me.cmdCancel.TabIndex = 1
    Me.cmdCancel.Text = "Cancel"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'frmExamResults
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(732, 612)
    Me.Controls.Add(Me.splMain)
    Me.Controls.Add(Me.bplBottom)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmExamResults"
    Me.Text = "frmExamResu"
    Me.splMain.Panel1.ResumeLayout(False)
    Me.splMain.Panel2.ResumeLayout(False)
    CType(Me.splMain, System.ComponentModel.ISupportInitialize).EndInit()
    Me.splMain.ResumeLayout(False)
    Me.bplTop.ResumeLayout(False)
    Me.splBottom.Panel1.ResumeLayout(False)
    Me.splBottom.Panel2.ResumeLayout(False)
    CType(Me.splBottom, System.ComponentModel.ISupportInitialize).EndInit()
    Me.splBottom.ResumeLayout(False)
    Me.Condition.ResumeLayout(False)
    Me.TabPage1.ResumeLayout(False)
    Me.pnlQuickEntry.ResumeLayout(False)
    Me.pnlQuickEntry.PerformLayout()
    Me.TabFindReplacePage.ResumeLayout(False)
    Me.TabFindReplaceHost.ResumeLayout(False)
    Me.TabFindReplaceMark.ResumeLayout(False)
    Me.TabFindReplaceMark.PerformLayout()
    Me.TabFindReplaceGrade.ResumeLayout(False)
    Me.TabFindReplaceGrade.PerformLayout()
    Me.TabFindReplaceResult.ResumeLayout(False)
    Me.TabFindReplaceResult.PerformLayout()
    Me.Panel1.ResumeLayout(False)
    Me.Panel1.PerformLayout()
    Me.TabPage3.ResumeLayout(False)
    Me.PanelEx2.ResumeLayout(False)
    Me.PanelEx2.PerformLayout()
    Me.bplBottom.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents splMain As System.Windows.Forms.SplitContainer
  Friend WithEvents bplTop As CDBNETCL.ButtonPanel
  Friend WithEvents cmdSelect As System.Windows.Forms.Button
  Friend WithEvents epl As CDBNETCL.EditPanel
  Friend WithEvents splBottom As System.Windows.Forms.SplitContainer
  Friend WithEvents dgr As CDBNETCL.DisplayGrid
  Friend WithEvents dgrComponents As CDBNETCL.DisplayGrid
  Friend WithEvents bplBottom As CDBNETCL.ButtonPanel
  Friend WithEvents cmdSave As System.Windows.Forms.Button
  Friend WithEvents cmdOK As System.Windows.Forms.Button
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents Condition As System.Windows.Forms.TabControl
  Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
  Friend WithEvents pnlQuickEntry As CDBNETCL.PanelEx
  Friend WithEvents txtResult As CDBNETCL.TextLookupBox
  Friend WithEvents txtGrade As CDBNETCL.TextLookupBox
  Friend WithEvents lblMark As System.Windows.Forms.Label
  Friend WithEvents txtMark As System.Windows.Forms.TextBox
  Friend WithEvents lblContactNumber As System.Windows.Forms.Label
  Friend WithEvents txtContactNumber As System.Windows.Forms.TextBox
  Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
  Friend WithEvents PanelEx2 As CDBNETCL.PanelEx
  Friend WithEvents txtChangeReason As CDBNETCL.TextLookupBox
  Friend WithEvents lblChangeReason As System.Windows.Forms.Label
  Friend WithEvents lblInfo As CDBNETCL.TransparentLabel
  Friend WithEvents splComponentSeparator As System.Windows.Forms.Splitter
  Friend WithEvents TabFindReplacePage As System.Windows.Forms.TabPage
  Friend WithEvents TabFindReplaceHost As System.Windows.Forms.TabControl
  Friend WithEvents TabFindReplaceMark As System.Windows.Forms.TabPage
  Friend WithEvents lblReplaceWith As System.Windows.Forms.Label
  Friend WithEvents lblFindWhat As System.Windows.Forms.Label
  Friend WithEvents txtMarkReplace As System.Windows.Forms.TextBox
  Friend WithEvents txtMarkFind As System.Windows.Forms.TextBox
  Friend WithEvents TabFindReplaceGrade As System.Windows.Forms.TabPage
  Friend WithEvents txtGradeReplace As CDBNETCL.TextLookupBox
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents Label2 As System.Windows.Forms.Label
  Friend WithEvents txtGradeFind As CDBNETCL.TextLookupBox
  Friend WithEvents TabFindReplaceResult As System.Windows.Forms.TabPage
  Friend WithEvents Label3 As System.Windows.Forms.Label
  Friend WithEvents Label4 As System.Windows.Forms.Label
  Friend WithEvents txtResultFind As CDBNETCL.TextLookupBox
  Friend WithEvents txtResultReplace As CDBNETCL.TextLookupBox
  Friend WithEvents Panel1 As System.Windows.Forms.Panel
  Friend WithEvents rdoResult As System.Windows.Forms.RadioButton
  Friend WithEvents rdoGrade As System.Windows.Forms.RadioButton
  Friend WithEvents rdoMark As System.Windows.Forms.RadioButton
  Friend WithEvents cmdReplaceAll As System.Windows.Forms.Button
  Friend WithEvents cmdReplace As System.Windows.Forms.Button
  Friend WithEvents cmdFindNext As System.Windows.Forms.Button
End Class
