<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmImport
  Inherits CDBNETCL.ThemedForm

  'Form overrides dispose to clean up the component list.
  <System.Diagnostics.DebuggerNonUserCode()>
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
  <System.Diagnostics.DebuggerStepThrough()>
  Private Sub InitializeComponent()
    Me.components = New System.ComponentModel.Container()
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmImport))
    Me.pnlAddressUpdate = New System.Windows.Forms.Panel()
    Me.dgrMenuStrip = New System.Windows.Forms.ContextMenuStrip(Me.components)
    Me.dgrMenuDateFormat = New System.Windows.Forms.ToolStripMenuItem()
    Me.dgrMenuMapAttribute = New System.Windows.Forms.ToolStripMenuItem()
    Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
    Me.tabMain = New CDBNETCL.TabControl()
    Me.tbpData = New System.Windows.Forms.TabPage()
    Me.lnkHelp = New System.Windows.Forms.LinkLabel()
    Me.dgr = New CDBNETCL.DisplayGrid()
    Me.lblDataSource = New System.Windows.Forms.Label()
    Me.lblSource = New System.Windows.Forms.Label()
    Me.txtDataSource = New CDBNETCL.TextLookupBox()
    Me.txtSource = New CDBNETCL.TextLookupBox()
    Me.grpDedup = New System.Windows.Forms.GroupBox()
    Me.optDedupNone = New System.Windows.Forms.RadioButton()
    Me.optDedupAddressOnly = New System.Windows.Forms.RadioButton()
    Me.optDedupFull = New System.Windows.Forms.RadioButton()
    Me.cboSeparator = New System.Windows.Forms.ComboBox()
    Me.lblSeperator = New System.Windows.Forms.Label()
    Me.cboTables = New System.Windows.Forms.ComboBox()
    Me.lblTableDesc = New System.Windows.Forms.Label()
    Me.lblGroups = New System.Windows.Forms.Label()
    Me.cboGroups = New System.Windows.Forms.ComboBox()
    Me.chkIgnore = New System.Windows.Forms.CheckBox()
    Me.cboAttrs = New System.Windows.Forms.ComboBox()
    Me.lblAttribute = New System.Windows.Forms.Label()
    Me.lblColumn = New System.Windows.Forms.Label()
    Me.cboType = New System.Windows.Forms.ComboBox()
    Me.lblDataImportType = New System.Windows.Forms.Label()
    Me.lblKey = New System.Windows.Forms.Label()
    Me.chkKey = New System.Windows.Forms.CheckBox()
    Me.lblMapValue = New System.Windows.Forms.Label()
    Me.optMapValueLookup = New System.Windows.Forms.RadioButton()
    Me.optMapValueNull = New System.Windows.Forms.RadioButton()
    Me.tbpOptions = New System.Windows.Forms.TabPage()
    Me.tabSub = New CDBNETCL.TabControl()
    Me.tbpGeneralOpt = New System.Windows.Forms.TabPage()
    Me.spltGeneralOpt = New System.Windows.Forms.SplitContainer()
    Me.grpProcessingOptions = New System.Windows.Forms.GroupBox()
    Me.chkAmendmentHistory = New System.Windows.Forms.CheckBox()
    Me.lblControlNumberBlockSize = New System.Windows.Forms.Label()
    Me.txtControlNumberBlockSize = New System.Windows.Forms.TextBox()
    Me.txtReplaceQuestionMarkWith = New System.Windows.Forms.TextBox()
    Me.chkReplaceQuestionMark = New System.Windows.Forms.CheckBox()
    Me.chkCMDSupp = New System.Windows.Forms.CheckBox()
    Me.chkCreateCMD = New System.Windows.Forms.CheckBox()
    Me.chkValCodes = New System.Windows.Forms.CheckBox()
    Me.chkNoIndexes = New System.Windows.Forms.CheckBox()
    Me.chkControlNumbers = New System.Windows.Forms.CheckBox()
    Me.grpLogFileOptions = New System.Windows.Forms.GroupBox()
    Me.optMIRecordsOriginalFile = New System.Windows.Forms.RadioButton()
    Me.optMIRecordsSuccFromFirstImport = New System.Windows.Forms.RadioButton()
    Me.optMIRecordsSuccFromPrevImport = New System.Windows.Forms.RadioButton()
    Me.lblMultipleImport = New System.Windows.Forms.Label()
    Me.lblDefinitionFile = New System.Windows.Forms.Label()
    Me.chkNoFileName = New System.Windows.Forms.CheckBox()
    Me.chkLogConversion = New System.Windows.Forms.CheckBox()
    Me.chkLogDedupAudit = New System.Windows.Forms.CheckBox()
    Me.chkLogDups = New System.Windows.Forms.CheckBox()
    Me.chkLogWarn = New System.Windows.Forms.CheckBox()
    Me.chkLogCreate = New System.Windows.Forms.CheckBox()
    Me.tbpMultImpRuns = New System.Windows.Forms.TabPage()
    Me.lblSelectedImportType = New System.Windows.Forms.Label()
    Me.chkDupAsError = New System.Windows.Forms.CheckBox()
    Me.lstMultiImportSelected = New System.Windows.Forms.ListBox()
    Me.cmdRemoveDefinition = New System.Windows.Forms.Button()
    Me.cmdAddDefinition = New System.Windows.Forms.Button()
    Me.lblAvailableImportTypes = New System.Windows.Forms.Label()
    Me.lstMultiImportAvailable = New System.Windows.Forms.ListBox()
    Me.tbpCustomOpt = New System.Windows.Forms.TabPage()
    Me.pnlConAndOrg = New System.Windows.Forms.Panel()
    Me.grpDataOption = New System.Windows.Forms.GroupBox()
    Me.chkAllowBlankForOrgName = New System.Windows.Forms.CheckBox()
    Me.txtOrgNumber = New CDBNETCL.TextLookupBox()
    Me.lblOrganisation = New System.Windows.Forms.Label()
    Me.chkAddPosition = New System.Windows.Forms.CheckBox()
    Me.chkEmployee = New System.Windows.Forms.CheckBox()
    Me.chkCacheMailsort = New System.Windows.Forms.CheckBox()
    Me.chkDefAddrFromUnknown = New System.Windows.Forms.CheckBox()
    Me.chkDefSupp = New System.Windows.Forms.CheckBox()
    Me.chkCreateGridRefs = New System.Windows.Forms.CheckBox()
    Me.chkRePostcode = New System.Windows.Forms.CheckBox()
    Me.chkPAFAddress = New System.Windows.Forms.CheckBox()
    Me.chkSurnameFirst = New System.Windows.Forms.CheckBox()
    Me.chkCaps = New System.Windows.Forms.CheckBox()
    Me.chkDear = New System.Windows.Forms.CheckBox()
    Me.grpDupUpdate = New System.Windows.Forms.GroupBox()
    Me.lblUpdateSub = New System.Windows.Forms.Label()
    Me.chkNameGatheringIncentives = New System.Windows.Forms.CheckBox()
    Me.chkActivity = New System.Windows.Forms.CheckBox()
    Me.chkUpdateWithNull = New System.Windows.Forms.CheckBox()
    Me.chkUpdateAll = New System.Windows.Forms.CheckBox()
    Me.chkUpdate = New System.Windows.Forms.CheckBox()
    Me.grpDeDuplication = New System.Windows.Forms.GroupBox()
    Me.chkOrgNamePostCodeAddress = New System.Windows.Forms.CheckBox()
    Me.chkBankDetailsDedup = New System.Windows.Forms.CheckBox()
    Me.chkOrgAddressPotDup = New System.Windows.Forms.CheckBox()
    Me.chkSoundexDedup = New System.Windows.Forms.CheckBox()
    Me.chkAddressDedup = New System.Windows.Forms.CheckBox()
    Me.chkForeInitDeDup = New System.Windows.Forms.CheckBox()
    Me.chkTitleDeDup = New System.Windows.Forms.CheckBox()
    Me.chkEmailDedup = New System.Windows.Forms.CheckBox()
    Me.chkNumberDeDup = New System.Windows.Forms.CheckBox()
    Me.chkExtRefDeDup = New System.Windows.Forms.CheckBox()
    Me.chkExclUnkAdd = New System.Windows.Forms.CheckBox()
    Me.pnlTableImport = New System.Windows.Forms.Panel()
    Me.chkEmptyBeforeImport = New System.Windows.Forms.CheckBox()
    Me.grp = New System.Windows.Forms.GroupBox()
    Me.lblTableImport = New System.Windows.Forms.Label()
    Me.chkUpdateAllTableImport = New System.Windows.Forms.CheckBox()
    Me.chkUpdateTableImport = New System.Windows.Forms.CheckBox()
    Me.pnlPayment = New System.Windows.Forms.Panel()
    Me.lblNumberOfDays = New System.Windows.Forms.Label()
    Me.txtNumberOfDays = New System.Windows.Forms.TextBox()
    Me.chkMatchSchPayment = New System.Windows.Forms.CheckBox()
    Me.chkCreateAct = New System.Windows.Forms.CheckBox()
    Me.chkSkipZeroAmt = New System.Windows.Forms.CheckBox()
    Me.chkProcessIncentives = New System.Windows.Forms.CheckBox()
    Me.chkReference = New System.Windows.Forms.CheckBox()
    Me.chkAddTransactions = New System.Windows.Forms.CheckBox()
    Me.chkNoFromFile = New System.Windows.Forms.CheckBox()
    Me.chkGiftAidRecords = New System.Windows.Forms.CheckBox()
    Me.grpTypeOfPayment = New System.Windows.Forms.GroupBox()
    Me.optPaymentsUnposted = New System.Windows.Forms.RadioButton()
    Me.optPaymentsPostedToNominal = New System.Windows.Forms.RadioButton()
    Me.optPaymentsFinHistory = New System.Windows.Forms.RadioButton()
    Me.optPaymentsPostedToCB = New System.Windows.Forms.RadioButton()
    Me.pnlDocument = New System.Windows.Forms.Panel()
    Me.grpDupUpdateOpt = New System.Windows.Forms.GroupBox()
    Me.chkUpdateAllDoc = New System.Windows.Forms.CheckBox()
    Me.chkUpdateDoc = New System.Windows.Forms.CheckBox()
    Me.lblUpdateExisting = New System.Windows.Forms.Label()
    Me.pnlStock = New System.Windows.Forms.Panel()
    Me.optStockSet = New System.Windows.Forms.RadioButton()
    Me.optStockUpdate = New System.Windows.Forms.RadioButton()
    Me.pnlBankTransactions = New System.Windows.Forms.Panel()
    Me.chkDASImport = New System.Windows.Forms.CheckBox()
    Me.pnlAddrUpdate = New System.Windows.Forms.Panel()
    Me.chkCacheMailsortAddr = New System.Windows.Forms.CheckBox()
    Me.chkExtractAddr = New System.Windows.Forms.CheckBox()
    Me.tbpDefaults = New System.Windows.Forms.TabPage()
    Me.dgrDefaults = New CDBNETCL.DisplayGrid()
    Me.pnlDefaults = New System.Windows.Forms.Panel()
    Me.txtLookupDefValue = New CDBNETCL.TextLookupBox()
    Me.lblElse = New System.Windows.Forms.Label()
    Me.dtpckValue = New System.Windows.Forms.DateTimePicker()
    Me.cmdDefaultAdd = New System.Windows.Forms.Button()
    Me.cboPatternValue = New System.Windows.Forms.ComboBox()
    Me.cboDefAttrs = New System.Windows.Forms.ComboBox()
    Me.chkCtrlNo = New System.Windows.Forms.CheckBox()
    Me.chkIncPerLine = New System.Windows.Forms.CheckBox()
    Me.lblValue = New System.Windows.Forms.Label()
    Me.lblDefAttribute = New System.Windows.Forms.Label()
    Me.txtDefValue = New System.Windows.Forms.TextBox()
    Me.SplitControl = New System.Windows.Forms.SplitContainer()
    Me.bpl = New CDBNETCL.ButtonPanel()
    Me.cmdStop = New System.Windows.Forms.Button()
    Me.cmdTest = New System.Windows.Forms.Button()
    Me.cmdOk = New System.Windows.Forms.Button()
    Me.cmdCancel = New System.Windows.Forms.Button()
    Me.cmdSave = New System.Windows.Forms.Button()
    Me.lblJobNumber = New System.Windows.Forms.Label()
    Me.lblStatus = New System.Windows.Forms.Label()
    Me.dgrMenuStrip.SuspendLayout()
    CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SplitContainer1.Panel1.SuspendLayout()
    Me.SplitContainer1.Panel2.SuspendLayout()
    Me.SplitContainer1.SuspendLayout()
    Me.tabMain.SuspendLayout()
    Me.tbpData.SuspendLayout()
    Me.grpDedup.SuspendLayout()
    Me.tbpOptions.SuspendLayout()
    Me.tabSub.SuspendLayout()
    Me.tbpGeneralOpt.SuspendLayout()
    CType(Me.spltGeneralOpt, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.spltGeneralOpt.Panel1.SuspendLayout()
    Me.spltGeneralOpt.Panel2.SuspendLayout()
    Me.spltGeneralOpt.SuspendLayout()
    Me.grpProcessingOptions.SuspendLayout()
    Me.grpLogFileOptions.SuspendLayout()
    Me.tbpMultImpRuns.SuspendLayout()
    Me.tbpCustomOpt.SuspendLayout()
    Me.pnlConAndOrg.SuspendLayout()
    Me.grpDataOption.SuspendLayout()
    Me.grpDupUpdate.SuspendLayout()
    Me.grpDeDuplication.SuspendLayout()
    Me.pnlTableImport.SuspendLayout()
    Me.grp.SuspendLayout()
    Me.pnlPayment.SuspendLayout()
    Me.grpTypeOfPayment.SuspendLayout()
    Me.pnlDocument.SuspendLayout()
    Me.grpDupUpdateOpt.SuspendLayout()
    Me.pnlStock.SuspendLayout()
    Me.pnlBankTransactions.SuspendLayout()
    Me.pnlAddrUpdate.SuspendLayout()
    Me.tbpDefaults.SuspendLayout()
    Me.pnlDefaults.SuspendLayout()
    CType(Me.SplitControl, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SplitControl.Panel1.SuspendLayout()
    Me.SplitControl.Panel2.SuspendLayout()
    Me.SplitControl.SuspendLayout()
    Me.bpl.SuspendLayout()
    Me.SuspendLayout()
    '
    'pnlAddressUpdate
    '
    Me.pnlAddressUpdate.Location = New System.Drawing.Point(0, 0)
    Me.pnlAddressUpdate.Name = "pnlAddressUpdate"
    Me.pnlAddressUpdate.Size = New System.Drawing.Size(904, 510)
    Me.pnlAddressUpdate.TabIndex = 18
    '
    'dgrMenuStrip
    '
    Me.dgrMenuStrip.ImageScalingSize = New System.Drawing.Size(20, 20)
    Me.dgrMenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.dgrMenuDateFormat, Me.dgrMenuMapAttribute})
    Me.dgrMenuStrip.Name = "dgrMenuStrip"
    Me.dgrMenuStrip.Size = New System.Drawing.Size(181, 56)
    '
    'dgrMenuDateFormat
    '
    Me.dgrMenuDateFormat.Name = "dgrMenuDateFormat"
    Me.dgrMenuDateFormat.Size = New System.Drawing.Size(180, 26)
    Me.dgrMenuDateFormat.Text = "Date Format"
    '
    'dgrMenuMapAttribute
    '
    Me.dgrMenuMapAttribute.Name = "dgrMenuMapAttribute"
    Me.dgrMenuMapAttribute.Size = New System.Drawing.Size(180, 26)
    Me.dgrMenuMapAttribute.Text = "Map Attribute..."
    '
    'SplitContainer1
    '
    Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
    Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
    Me.SplitContainer1.Name = "SplitContainer1"
    Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
    '
    'SplitContainer1.Panel1
    '
    Me.SplitContainer1.Panel1.Controls.Add(Me.tabMain)
    '
    'SplitContainer1.Panel2
    '
    Me.SplitContainer1.Panel2.Controls.Add(Me.SplitControl)
    Me.SplitContainer1.Size = New System.Drawing.Size(929, 675)
    Me.SplitContainer1.SplitterDistance = 603
    Me.SplitContainer1.TabIndex = 19
    '
    'tabMain
    '
    Me.tabMain.Controls.Add(Me.tbpData)
    Me.tabMain.Controls.Add(Me.tbpOptions)
    Me.tabMain.Controls.Add(Me.tbpDefaults)
    Me.tabMain.Dock = System.Windows.Forms.DockStyle.Fill
    Me.tabMain.ItemSize = New System.Drawing.Size(267, 22)
    Me.tabMain.Location = New System.Drawing.Point(0, 0)
    Me.tabMain.Name = "tabMain"
    Me.tabMain.SelectedIndex = 0
    Me.tabMain.Size = New System.Drawing.Size(929, 603)
    Me.tabMain.SizeMode = System.Windows.Forms.TabSizeMode.Fixed
    Me.tabMain.TabIndex = 1
    '
    'tbpData
    '
    Me.tbpData.AutoScroll = True
    Me.tbpData.Controls.Add(Me.lnkHelp)
    Me.tbpData.Controls.Add(Me.dgr)
    Me.tbpData.Controls.Add(Me.lblDataSource)
    Me.tbpData.Controls.Add(Me.lblSource)
    Me.tbpData.Controls.Add(Me.txtDataSource)
    Me.tbpData.Controls.Add(Me.txtSource)
    Me.tbpData.Controls.Add(Me.grpDedup)
    Me.tbpData.Controls.Add(Me.cboSeparator)
    Me.tbpData.Controls.Add(Me.lblSeperator)
    Me.tbpData.Controls.Add(Me.cboTables)
    Me.tbpData.Controls.Add(Me.lblTableDesc)
    Me.tbpData.Controls.Add(Me.lblGroups)
    Me.tbpData.Controls.Add(Me.cboGroups)
    Me.tbpData.Controls.Add(Me.chkIgnore)
    Me.tbpData.Controls.Add(Me.cboAttrs)
    Me.tbpData.Controls.Add(Me.lblAttribute)
    Me.tbpData.Controls.Add(Me.lblColumn)
    Me.tbpData.Controls.Add(Me.cboType)
    Me.tbpData.Controls.Add(Me.lblDataImportType)
    Me.tbpData.Controls.Add(Me.lblKey)
    Me.tbpData.Controls.Add(Me.chkKey)
    Me.tbpData.Controls.Add(Me.lblMapValue)
    Me.tbpData.Controls.Add(Me.optMapValueLookup)
    Me.tbpData.Controls.Add(Me.optMapValueNull)
    Me.tbpData.Location = New System.Drawing.Point(4, 26)
    Me.tbpData.Name = "tbpData"
    Me.tbpData.Size = New System.Drawing.Size(921, 573)
    Me.tbpData.TabIndex = 0
    Me.tbpData.Text = "Data"
    Me.tbpData.UseVisualStyleBackColor = True
    '
    'lnkHelp
    '
    Me.lnkHelp.AutoSize = True
    Me.lnkHelp.Dock = System.Windows.Forms.DockStyle.Right
    Me.lnkHelp.Location = New System.Drawing.Point(853, 0)
    Me.lnkHelp.Name = "lnkHelp"
    Me.lnkHelp.Size = New System.Drawing.Size(70, 17)
    Me.lnkHelp.TabIndex = 23
    Me.lnkHelp.TabStop = True
    Me.lnkHelp.Text = "View Help"
    '
    'dgr
    '
    Me.dgr.AccessibleName = "Display Grid"
    Me.dgr.ActiveColumn = 0
    Me.dgr.AllowColumnResize = True
    Me.dgr.AllowSorting = False
    Me.dgr.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.dgr.AutoSetHeight = False
    Me.dgr.AutoSetRowHeight = False
    Me.dgr.DisplayTitle = Nothing
    Me.dgr.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!)
    Me.dgr.Location = New System.Drawing.Point(4, 215)
    Me.dgr.MaintenanceDesc = Nothing
    Me.dgr.MaxGridRows = 6
    Me.dgr.MultipleSelect = False
    Me.dgr.Name = "dgr"
    Me.dgr.RowCount = 10
    Me.dgr.ShowIfEmpty = False
    Me.dgr.Size = New System.Drawing.Size(917, 322)
    Me.dgr.SuppressHyperLinkFormat = False
    Me.dgr.TabIndex = 22
    '
    'lblDataSource
    '
    Me.lblDataSource.AutoSize = True
    Me.lblDataSource.Location = New System.Drawing.Point(8, 169)
    Me.lblDataSource.Name = "lblDataSource"
    Me.lblDataSource.Size = New System.Drawing.Size(91, 17)
    Me.lblDataSource.TabIndex = 14
    Me.lblDataSource.Text = "Data Source:"
    '
    'lblSource
    '
    Me.lblSource.AutoSize = True
    Me.lblSource.Location = New System.Drawing.Point(8, 139)
    Me.lblSource.Name = "lblSource"
    Me.lblSource.Size = New System.Drawing.Size(57, 17)
    Me.lblSource.TabIndex = 12
    Me.lblSource.Text = "Source:"
    '
    'txtDataSource
    '
    Me.txtDataSource.ActiveOnly = False
    Me.txtDataSource.BackColor = System.Drawing.Color.Transparent
    Me.txtDataSource.CustomFormNumber = 0
    Me.txtDataSource.Description = ""
    Me.txtDataSource.EnabledProperty = True
    Me.txtDataSource.ExamCentreId = 0
    Me.txtDataSource.ExamCentreUnitId = 0
    Me.txtDataSource.ExamUnitLinkId = 0
    Me.txtDataSource.HasDependancies = False
    Me.txtDataSource.IsDesign = False
    Me.txtDataSource.Location = New System.Drawing.Point(145, 166)
    Me.txtDataSource.MaxLength = 32767
    Me.txtDataSource.MultipleValuesSupported = False
    Me.txtDataSource.Name = "txtDataSource"
    Me.txtDataSource.OriginalText = Nothing
    Me.txtDataSource.PreventHistoricalSelection = False
    Me.txtDataSource.ReadOnlyProperty = False
    Me.txtDataSource.Size = New System.Drawing.Size(408, 24)
    Me.txtDataSource.TabIndex = 15
    Me.txtDataSource.TextReadOnly = False
    Me.txtDataSource.TotalWidth = 408
    Me.txtDataSource.ValidationRequired = True
    Me.txtDataSource.WarningMessage = Nothing
    '
    'txtSource
    '
    Me.txtSource.ActiveOnly = False
    Me.txtSource.BackColor = System.Drawing.Color.Transparent
    Me.txtSource.CustomFormNumber = 0
    Me.txtSource.Description = ""
    Me.txtSource.EnabledProperty = True
    Me.txtSource.ExamCentreId = 0
    Me.txtSource.ExamCentreUnitId = 0
    Me.txtSource.ExamUnitLinkId = 0
    Me.txtSource.HasDependancies = False
    Me.txtSource.IsDesign = False
    Me.txtSource.Location = New System.Drawing.Point(145, 136)
    Me.txtSource.MaxLength = 32767
    Me.txtSource.MultipleValuesSupported = False
    Me.txtSource.Name = "txtSource"
    Me.txtSource.OriginalText = Nothing
    Me.txtSource.PreventHistoricalSelection = False
    Me.txtSource.ReadOnlyProperty = False
    Me.txtSource.Size = New System.Drawing.Size(408, 24)
    Me.txtSource.TabIndex = 13
    Me.txtSource.TextReadOnly = False
    Me.txtSource.TotalWidth = 408
    Me.txtSource.ValidationRequired = True
    Me.txtSource.WarningMessage = Nothing
    '
    'grpDedup
    '
    Me.grpDedup.Controls.Add(Me.optDedupNone)
    Me.grpDedup.Controls.Add(Me.optDedupAddressOnly)
    Me.grpDedup.Controls.Add(Me.optDedupFull)
    Me.grpDedup.Location = New System.Drawing.Point(582, 109)
    Me.grpDedup.Name = "grpDedup"
    Me.grpDedup.Size = New System.Drawing.Size(271, 100)
    Me.grpDedup.TabIndex = 16
    Me.grpDedup.TabStop = False
    Me.grpDedup.Text = "De-Duplication"
    '
    'optDedupNone
    '
    Me.optDedupNone.AutoSize = True
    Me.optDedupNone.Location = New System.Drawing.Point(7, 74)
    Me.optDedupNone.Name = "optDedupNone"
    Me.optDedupNone.Size = New System.Drawing.Size(63, 21)
    Me.optDedupNone.TabIndex = 2
    Me.optDedupNone.Text = "None"
    Me.optDedupNone.UseVisualStyleBackColor = True
    '
    'optDedupAddressOnly
    '
    Me.optDedupAddressOnly.AutoSize = True
    Me.optDedupAddressOnly.Location = New System.Drawing.Point(7, 48)
    Me.optDedupAddressOnly.Name = "optDedupAddressOnly"
    Me.optDedupAddressOnly.Size = New System.Drawing.Size(114, 21)
    Me.optDedupAddressOnly.TabIndex = 1
    Me.optDedupAddressOnly.Text = "Address Only"
    Me.optDedupAddressOnly.UseVisualStyleBackColor = True
    '
    'optDedupFull
    '
    Me.optDedupFull.AutoSize = True
    Me.optDedupFull.Checked = True
    Me.optDedupFull.Location = New System.Drawing.Point(7, 21)
    Me.optDedupFull.Name = "optDedupFull"
    Me.optDedupFull.Size = New System.Drawing.Size(51, 21)
    Me.optDedupFull.TabIndex = 0
    Me.optDedupFull.TabStop = True
    Me.optDedupFull.Text = "Full"
    Me.optDedupFull.UseVisualStyleBackColor = True
    '
    'cboSeparator
    '
    Me.cboSeparator.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cboSeparator.FormattingEnabled = True
    Me.cboSeparator.Location = New System.Drawing.Point(726, 76)
    Me.cboSeparator.Name = "cboSeparator"
    Me.cboSeparator.Size = New System.Drawing.Size(127, 24)
    Me.cboSeparator.TabIndex = 9
    '
    'lblSeperator
    '
    Me.lblSeperator.AutoSize = True
    Me.lblSeperator.Location = New System.Drawing.Point(579, 79)
    Me.lblSeperator.Name = "lblSeperator"
    Me.lblSeperator.Size = New System.Drawing.Size(109, 17)
    Me.lblSeperator.TabIndex = 8
    Me.lblSeperator.Text = "Field Separator:"
    '
    'cboTables
    '
    Me.cboTables.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cboTables.FormattingEnabled = True
    Me.cboTables.Location = New System.Drawing.Point(145, 46)
    Me.cboTables.Name = "cboTables"
    Me.cboTables.Size = New System.Drawing.Size(321, 24)
    Me.cboTables.TabIndex = 4
    '
    'lblTableDesc
    '
    Me.lblTableDesc.AutoSize = True
    Me.lblTableDesc.Location = New System.Drawing.Point(8, 49)
    Me.lblTableDesc.Name = "lblTableDesc"
    Me.lblTableDesc.Size = New System.Drawing.Size(123, 17)
    Me.lblTableDesc.TabIndex = 3
    Me.lblTableDesc.Text = "Table Description:"
    '
    'lblGroups
    '
    Me.lblGroups.AutoSize = True
    Me.lblGroups.Location = New System.Drawing.Point(520, 49)
    Me.lblGroups.Name = "lblGroups"
    Me.lblGroups.Size = New System.Drawing.Size(52, 17)
    Me.lblGroups.TabIndex = 5
    Me.lblGroups.Text = "Group:"
    '
    'cboGroups
    '
    Me.cboGroups.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cboGroups.FormattingEnabled = True
    Me.cboGroups.Location = New System.Drawing.Point(582, 46)
    Me.cboGroups.Name = "cboGroups"
    Me.cboGroups.Size = New System.Drawing.Size(271, 24)
    Me.cboGroups.TabIndex = 6
    '
    'chkIgnore
    '
    Me.chkIgnore.AutoSize = True
    Me.chkIgnore.Location = New System.Drawing.Point(582, 18)
    Me.chkIgnore.Name = "chkIgnore"
    Me.chkIgnore.Size = New System.Drawing.Size(132, 21)
    Me.chkIgnore.TabIndex = 2
    Me.chkIgnore.Text = "Ignore First Row"
    Me.chkIgnore.UseVisualStyleBackColor = True
    '
    'cboAttrs
    '
    Me.cboAttrs.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cboAttrs.FormattingEnabled = True
    Me.cboAttrs.Location = New System.Drawing.Point(145, 106)
    Me.cboAttrs.Name = "cboAttrs"
    Me.cboAttrs.Size = New System.Drawing.Size(408, 24)
    Me.cboAttrs.TabIndex = 11
    '
    'lblAttribute
    '
    Me.lblAttribute.AutoSize = True
    Me.lblAttribute.Location = New System.Drawing.Point(8, 109)
    Me.lblAttribute.Name = "lblAttribute"
    Me.lblAttribute.Size = New System.Drawing.Size(65, 17)
    Me.lblAttribute.TabIndex = 10
    Me.lblAttribute.Text = "Attribute:"
    '
    'lblColumn
    '
    Me.lblColumn.AutoSize = True
    Me.lblColumn.Location = New System.Drawing.Point(8, 79)
    Me.lblColumn.Name = "lblColumn"
    Me.lblColumn.Size = New System.Drawing.Size(485, 17)
    Me.lblColumn.TabIndex = 7
    Me.lblColumn.Text = "Select a column then select the corresponding Attribute from the combo box"
    '
    'cboType
    '
    Me.cboType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cboType.FormattingEnabled = True
    Me.cboType.Location = New System.Drawing.Point(145, 15)
    Me.cboType.Name = "cboType"
    Me.cboType.Size = New System.Drawing.Size(321, 24)
    Me.cboType.TabIndex = 1
    '
    'lblDataImportType
    '
    Me.lblDataImportType.AutoSize = True
    Me.lblDataImportType.Location = New System.Drawing.Point(8, 18)
    Me.lblDataImportType.Name = "lblDataImportType"
    Me.lblDataImportType.Size = New System.Drawing.Size(121, 17)
    Me.lblDataImportType.TabIndex = 0
    Me.lblDataImportType.Text = "Data Import Type:"
    '
    'lblKey
    '
    Me.lblKey.AutoSize = True
    Me.lblKey.Location = New System.Drawing.Point(8, 139)
    Me.lblKey.Name = "lblKey"
    Me.lblKey.Size = New System.Drawing.Size(476, 17)
    Me.lblKey.TabIndex = 17
    Me.lblKey.Text = "Select a column then click on the checkbox to assign it as the key Attribute"
    '
    'chkKey
    '
    Me.chkKey.AutoSize = True
    Me.chkKey.Location = New System.Drawing.Point(11, 169)
    Me.chkKey.Name = "chkKey"
    Me.chkKey.Size = New System.Drawing.Size(155, 21)
    Me.chkKey.TabIndex = 18
    Me.chkKey.Text = "Set as Key Attribute"
    Me.chkKey.UseVisualStyleBackColor = True
    '
    'lblMapValue
    '
    Me.lblMapValue.AutoSize = True
    Me.lblMapValue.Location = New System.Drawing.Point(8, 149)
    Me.lblMapValue.Name = "lblMapValue"
    Me.lblMapValue.Size = New System.Drawing.Size(179, 17)
    Me.lblMapValue.TabIndex = 19
    Me.lblMapValue.Text = "For missing values, return: "
    '
    'optMapValueLookup
    '
    Me.optMapValueLookup.AutoSize = True
    Me.optMapValueLookup.Location = New System.Drawing.Point(91, 173)
    Me.optMapValueLookup.Name = "optMapValueLookup"
    Me.optMapValueLookup.Size = New System.Drawing.Size(114, 21)
    Me.optMapValueLookup.TabIndex = 21
    Me.optMapValueLookup.TabStop = True
    Me.optMapValueLookup.Text = "Lookup value"
    Me.optMapValueLookup.UseVisualStyleBackColor = True
    '
    'optMapValueNull
    '
    Me.optMapValueNull.AutoSize = True
    Me.optMapValueNull.Checked = True
    Me.optMapValueNull.Location = New System.Drawing.Point(11, 173)
    Me.optMapValueNull.Name = "optMapValueNull"
    Me.optMapValueNull.Size = New System.Drawing.Size(53, 21)
    Me.optMapValueNull.TabIndex = 20
    Me.optMapValueNull.TabStop = True
    Me.optMapValueNull.Text = "Null"
    Me.optMapValueNull.UseVisualStyleBackColor = True
    '
    'tbpOptions
    '
    Me.tbpOptions.Controls.Add(Me.tabSub)
    Me.tbpOptions.Location = New System.Drawing.Point(4, 26)
    Me.tbpOptions.Name = "tbpOptions"
    Me.tbpOptions.Size = New System.Drawing.Size(192, 70)
    Me.tbpOptions.TabIndex = 1
    Me.tbpOptions.Text = "Options"
    Me.tbpOptions.UseVisualStyleBackColor = True
    '
    'tabSub
    '
    Me.tabSub.Controls.Add(Me.tbpGeneralOpt)
    Me.tabSub.Controls.Add(Me.tbpMultImpRuns)
    Me.tabSub.Controls.Add(Me.tbpCustomOpt)
    Me.tabSub.Dock = System.Windows.Forms.DockStyle.Fill
    Me.tabSub.ItemSize = New System.Drawing.Size(139, 22)
    Me.tabSub.Location = New System.Drawing.Point(0, 0)
    Me.tabSub.Name = "tabSub"
    Me.tabSub.SelectedIndex = 0
    Me.tabSub.Size = New System.Drawing.Size(192, 70)
    Me.tabSub.SizeMode = System.Windows.Forms.TabSizeMode.Fixed
    Me.tabSub.TabIndex = 0
    '
    'tbpGeneralOpt
    '
    Me.tbpGeneralOpt.AutoScroll = True
    Me.tbpGeneralOpt.Controls.Add(Me.spltGeneralOpt)
    Me.tbpGeneralOpt.Location = New System.Drawing.Point(4, 26)
    Me.tbpGeneralOpt.Name = "tbpGeneralOpt"
    Me.tbpGeneralOpt.Size = New System.Drawing.Size(184, 40)
    Me.tbpGeneralOpt.TabIndex = 0
    Me.tbpGeneralOpt.Text = "General Options"
    Me.tbpGeneralOpt.UseVisualStyleBackColor = True
    '
    'spltGeneralOpt
    '
    Me.spltGeneralOpt.Dock = System.Windows.Forms.DockStyle.Fill
    Me.spltGeneralOpt.Location = New System.Drawing.Point(0, 0)
    Me.spltGeneralOpt.Name = "spltGeneralOpt"
    '
    'spltGeneralOpt.Panel1
    '
    Me.spltGeneralOpt.Panel1.Controls.Add(Me.grpProcessingOptions)
    '
    'spltGeneralOpt.Panel2
    '
    Me.spltGeneralOpt.Panel2.Controls.Add(Me.grpLogFileOptions)
    Me.spltGeneralOpt.Size = New System.Drawing.Size(184, 40)
    Me.spltGeneralOpt.SplitterDistance = 67
    Me.spltGeneralOpt.TabIndex = 0
    '
    'grpProcessingOptions
    '
    Me.grpProcessingOptions.Controls.Add(Me.chkAmendmentHistory)
    Me.grpProcessingOptions.Controls.Add(Me.lblControlNumberBlockSize)
    Me.grpProcessingOptions.Controls.Add(Me.txtControlNumberBlockSize)
    Me.grpProcessingOptions.Controls.Add(Me.txtReplaceQuestionMarkWith)
    Me.grpProcessingOptions.Controls.Add(Me.chkReplaceQuestionMark)
    Me.grpProcessingOptions.Controls.Add(Me.chkCMDSupp)
    Me.grpProcessingOptions.Controls.Add(Me.chkCreateCMD)
    Me.grpProcessingOptions.Controls.Add(Me.chkValCodes)
    Me.grpProcessingOptions.Controls.Add(Me.chkNoIndexes)
    Me.grpProcessingOptions.Controls.Add(Me.chkControlNumbers)
    Me.grpProcessingOptions.Dock = System.Windows.Forms.DockStyle.Fill
    Me.grpProcessingOptions.Location = New System.Drawing.Point(0, 0)
    Me.grpProcessingOptions.Name = "grpProcessingOptions"
    Me.grpProcessingOptions.Size = New System.Drawing.Size(67, 40)
    Me.grpProcessingOptions.TabIndex = 0
    Me.grpProcessingOptions.TabStop = False
    Me.grpProcessingOptions.Text = "Processing Options"
    '
    'chkAmendmentHistory
    '
    Me.chkAmendmentHistory.AutoSize = True
    Me.chkAmendmentHistory.Location = New System.Drawing.Point(10, 268)
    Me.chkAmendmentHistory.Name = "chkAmendmentHistory"
    Me.chkAmendmentHistory.Size = New System.Drawing.Size(199, 21)
    Me.chkAmendmentHistory.TabIndex = 11
    Me.chkAmendmentHistory.Text = "Create Amendment History"
    Me.chkAmendmentHistory.UseVisualStyleBackColor = True
    '
    'lblControlNumberBlockSize
    '
    Me.lblControlNumberBlockSize.AutoSize = True
    Me.lblControlNumberBlockSize.Location = New System.Drawing.Point(81, 232)
    Me.lblControlNumberBlockSize.Name = "lblControlNumberBlockSize"
    Me.lblControlNumberBlockSize.Size = New System.Drawing.Size(176, 17)
    Me.lblControlNumberBlockSize.TabIndex = 10
    Me.lblControlNumberBlockSize.Text = "Control Number Block Size"
    '
    'txtControlNumberBlockSize
    '
    Me.txtControlNumberBlockSize.Location = New System.Drawing.Point(9, 229)
    Me.txtControlNumberBlockSize.Name = "txtControlNumberBlockSize"
    Me.txtControlNumberBlockSize.Size = New System.Drawing.Size(51, 22)
    Me.txtControlNumberBlockSize.TabIndex = 9
    '
    'txtReplaceQuestionMarkWith
    '
    Me.txtReplaceQuestionMarkWith.Location = New System.Drawing.Point(278, 194)
    Me.txtReplaceQuestionMarkWith.Name = "txtReplaceQuestionMarkWith"
    Me.txtReplaceQuestionMarkWith.Size = New System.Drawing.Size(26, 22)
    Me.txtReplaceQuestionMarkWith.TabIndex = 8
    '
    'chkReplaceQuestionMark
    '
    Me.chkReplaceQuestionMark.AutoSize = True
    Me.chkReplaceQuestionMark.Location = New System.Drawing.Point(10, 192)
    Me.chkReplaceQuestionMark.Name = "chkReplaceQuestionMark"
    Me.chkReplaceQuestionMark.Size = New System.Drawing.Size(262, 21)
    Me.chkReplaceQuestionMark.TabIndex = 7
    Me.chkReplaceQuestionMark.Text = "For character fields Replace '?' with: "
    Me.chkReplaceQuestionMark.UseVisualStyleBackColor = True
    '
    'chkCMDSupp
    '
    Me.chkCMDSupp.Location = New System.Drawing.Point(10, 140)
    Me.chkCMDSupp.Name = "chkCMDSupp"
    Me.chkCMDSupp.Size = New System.Drawing.Size(332, 46)
    Me.chkCMDSupp.TabIndex = 6
    Me.chkCMDSupp.Text = "Create Contact Mailing Documents for Contacts with Warning Suppressions"
    Me.chkCMDSupp.UseVisualStyleBackColor = True
    '
    'chkCreateCMD
    '
    Me.chkCreateCMD.AutoSize = True
    Me.chkCreateCMD.Location = New System.Drawing.Point(10, 115)
    Me.chkCreateCMD.Name = "chkCreateCMD"
    Me.chkCreateCMD.Size = New System.Drawing.Size(247, 21)
    Me.chkCreateCMD.TabIndex = 5
    Me.chkCreateCMD.Text = "Create Contact Mailing Documents"
    Me.chkCreateCMD.UseVisualStyleBackColor = True
    '
    'chkValCodes
    '
    Me.chkValCodes.AutoSize = True
    Me.chkValCodes.Location = New System.Drawing.Point(10, 88)
    Me.chkValCodes.Name = "chkValCodes"
    Me.chkValCodes.Size = New System.Drawing.Size(125, 21)
    Me.chkValCodes.TabIndex = 4
    Me.chkValCodes.Text = "Validate Codes"
    Me.chkValCodes.UseVisualStyleBackColor = True
    '
    'chkNoIndexes
    '
    Me.chkNoIndexes.AutoSize = True
    Me.chkNoIndexes.Location = New System.Drawing.Point(10, 61)
    Me.chkNoIndexes.Name = "chkNoIndexes"
    Me.chkNoIndexes.Size = New System.Drawing.Size(134, 21)
    Me.chkNoIndexes.TabIndex = 3
    Me.chkNoIndexes.Text = "Remove Indexes"
    Me.chkNoIndexes.UseVisualStyleBackColor = True
    '
    'chkControlNumbers
    '
    Me.chkControlNumbers.AutoSize = True
    Me.chkControlNumbers.Checked = True
    Me.chkControlNumbers.CheckState = System.Windows.Forms.CheckState.Checked
    Me.chkControlNumbers.Location = New System.Drawing.Point(10, 33)
    Me.chkControlNumbers.Name = "chkControlNumbers"
    Me.chkControlNumbers.Size = New System.Drawing.Size(179, 21)
    Me.chkControlNumbers.TabIndex = 1
    Me.chkControlNumbers.Tag = ""
    Me.chkControlNumbers.Text = "Check Control Numbers"
    Me.chkControlNumbers.UseVisualStyleBackColor = True
    '
    'grpLogFileOptions
    '
    Me.grpLogFileOptions.Controls.Add(Me.optMIRecordsOriginalFile)
    Me.grpLogFileOptions.Controls.Add(Me.optMIRecordsSuccFromFirstImport)
    Me.grpLogFileOptions.Controls.Add(Me.optMIRecordsSuccFromPrevImport)
    Me.grpLogFileOptions.Controls.Add(Me.lblMultipleImport)
    Me.grpLogFileOptions.Controls.Add(Me.lblDefinitionFile)
    Me.grpLogFileOptions.Controls.Add(Me.chkNoFileName)
    Me.grpLogFileOptions.Controls.Add(Me.chkLogConversion)
    Me.grpLogFileOptions.Controls.Add(Me.chkLogDedupAudit)
    Me.grpLogFileOptions.Controls.Add(Me.chkLogDups)
    Me.grpLogFileOptions.Controls.Add(Me.chkLogWarn)
    Me.grpLogFileOptions.Controls.Add(Me.chkLogCreate)
    Me.grpLogFileOptions.Dock = System.Windows.Forms.DockStyle.Fill
    Me.grpLogFileOptions.Location = New System.Drawing.Point(0, 0)
    Me.grpLogFileOptions.Name = "grpLogFileOptions"
    Me.grpLogFileOptions.Size = New System.Drawing.Size(113, 40)
    Me.grpLogFileOptions.TabIndex = 1
    Me.grpLogFileOptions.TabStop = False
    Me.grpLogFileOptions.Text = "Log File Options"
    '
    'optMIRecordsOriginalFile
    '
    Me.optMIRecordsOriginalFile.AutoSize = True
    Me.optMIRecordsOriginalFile.Location = New System.Drawing.Point(11, 347)
    Me.optMIRecordsOriginalFile.Name = "optMIRecordsOriginalFile"
    Me.optMIRecordsOriginalFile.Size = New System.Drawing.Size(133, 21)
    Me.optMIRecordsOriginalFile.TabIndex = 15
    Me.optMIRecordsOriginalFile.Text = "Use Original File"
    Me.optMIRecordsOriginalFile.UseVisualStyleBackColor = True
    '
    'optMIRecordsSuccFromFirstImport
    '
    Me.optMIRecordsSuccFromFirstImport.AutoSize = True
    Me.optMIRecordsSuccFromFirstImport.Location = New System.Drawing.Point(11, 320)
    Me.optMIRecordsSuccFromFirstImport.Name = "optMIRecordsSuccFromFirstImport"
    Me.optMIRecordsSuccFromFirstImport.Size = New System.Drawing.Size(236, 21)
    Me.optMIRecordsSuccFromFirstImport.TabIndex = 14
    Me.optMIRecordsSuccFromFirstImport.Text = "Use Successes From First Import"
    Me.optMIRecordsSuccFromFirstImport.UseVisualStyleBackColor = True
    '
    'optMIRecordsSuccFromPrevImport
    '
    Me.optMIRecordsSuccFromPrevImport.AutoSize = True
    Me.optMIRecordsSuccFromPrevImport.Checked = True
    Me.optMIRecordsSuccFromPrevImport.Location = New System.Drawing.Point(11, 293)
    Me.optMIRecordsSuccFromPrevImport.Name = "optMIRecordsSuccFromPrevImport"
    Me.optMIRecordsSuccFromPrevImport.Size = New System.Drawing.Size(264, 21)
    Me.optMIRecordsSuccFromPrevImport.TabIndex = 13
    Me.optMIRecordsSuccFromPrevImport.TabStop = True
    Me.optMIRecordsSuccFromPrevImport.Text = "Use Successes From Previous Import"
    Me.optMIRecordsSuccFromPrevImport.UseVisualStyleBackColor = True
    '
    'lblMultipleImport
    '
    Me.lblMultipleImport.AutoSize = True
    Me.lblMultipleImport.Location = New System.Drawing.Point(8, 268)
    Me.lblMultipleImport.Name = "lblMultipleImport"
    Me.lblMultipleImport.Size = New System.Drawing.Size(203, 17)
    Me.lblMultipleImport.TabIndex = 12
    Me.lblMultipleImport.Text = "Multiple Import Records to use:"
    '
    'lblDefinitionFile
    '
    Me.lblDefinitionFile.AutoSize = True
    Me.lblDefinitionFile.Location = New System.Drawing.Point(8, 194)
    Me.lblDefinitionFile.Name = "lblDefinitionFile"
    Me.lblDefinitionFile.Size = New System.Drawing.Size(150, 17)
    Me.lblDefinitionFile.TabIndex = 11
    Me.lblDefinitionFile.Text = "Definition File Options:"
    '
    'chkNoFileName
    '
    Me.chkNoFileName.AutoSize = True
    Me.chkNoFileName.Location = New System.Drawing.Point(11, 219)
    Me.chkNoFileName.Name = "chkNoFileName"
    Me.chkNoFileName.Size = New System.Drawing.Size(321, 21)
    Me.chkNoFileName.TabIndex = 6
    Me.chkNoFileName.Text = "Save definition file without the import file name"
    Me.chkNoFileName.UseVisualStyleBackColor = True
    '
    'chkLogConversion
    '
    Me.chkLogConversion.AutoSize = True
    Me.chkLogConversion.Location = New System.Drawing.Point(11, 140)
    Me.chkLogConversion.Name = "chkLogConversion"
    Me.chkLogConversion.Size = New System.Drawing.Size(218, 21)
    Me.chkLogConversion.TabIndex = 5
    Me.chkLogConversion.Text = "Include Conversion Messages"
    Me.chkLogConversion.UseVisualStyleBackColor = True
    '
    'chkLogDedupAudit
    '
    Me.chkLogDedupAudit.AutoSize = True
    Me.chkLogDedupAudit.Location = New System.Drawing.Point(11, 113)
    Me.chkLogDedupAudit.Name = "chkLogDedupAudit"
    Me.chkLogDedupAudit.Size = New System.Drawing.Size(228, 21)
    Me.chkLogDedupAudit.TabIndex = 4
    Me.chkLogDedupAudit.Text = "Include Deduplication Audit Info"
    Me.chkLogDedupAudit.UseVisualStyleBackColor = True
    '
    'chkLogDups
    '
    Me.chkLogDups.AutoSize = True
    Me.chkLogDups.Location = New System.Drawing.Point(11, 86)
    Me.chkLogDups.Name = "chkLogDups"
    Me.chkLogDups.Size = New System.Drawing.Size(333, 21)
    Me.chkLogDups.TabIndex = 3
    Me.chkLogDups.Text = "Include Duplicate Contact, Org and Address info"
    Me.chkLogDups.UseVisualStyleBackColor = True
    '
    'chkLogWarn
    '
    Me.chkLogWarn.AutoSize = True
    Me.chkLogWarn.Location = New System.Drawing.Point(11, 59)
    Me.chkLogWarn.Name = "chkLogWarn"
    Me.chkLogWarn.Size = New System.Drawing.Size(200, 21)
    Me.chkLogWarn.TabIndex = 2
    Me.chkLogWarn.Text = "Include Warning Messages"
    Me.chkLogWarn.UseVisualStyleBackColor = True
    '
    'chkLogCreate
    '
    Me.chkLogCreate.AutoSize = True
    Me.chkLogCreate.Location = New System.Drawing.Point(11, 32)
    Me.chkLogCreate.Name = "chkLogCreate"
    Me.chkLogCreate.Size = New System.Drawing.Size(189, 21)
    Me.chkLogCreate.TabIndex = 1
    Me.chkLogCreate.Text = "Include Create Messages"
    Me.chkLogCreate.UseVisualStyleBackColor = True
    '
    'tbpMultImpRuns
    '
    Me.tbpMultImpRuns.Controls.Add(Me.lblSelectedImportType)
    Me.tbpMultImpRuns.Controls.Add(Me.chkDupAsError)
    Me.tbpMultImpRuns.Controls.Add(Me.lstMultiImportSelected)
    Me.tbpMultImpRuns.Controls.Add(Me.cmdRemoveDefinition)
    Me.tbpMultImpRuns.Controls.Add(Me.cmdAddDefinition)
    Me.tbpMultImpRuns.Controls.Add(Me.lblAvailableImportTypes)
    Me.tbpMultImpRuns.Controls.Add(Me.lstMultiImportAvailable)
    Me.tbpMultImpRuns.Location = New System.Drawing.Point(4, 26)
    Me.tbpMultImpRuns.Name = "tbpMultImpRuns"
    Me.tbpMultImpRuns.Size = New System.Drawing.Size(192, 70)
    Me.tbpMultImpRuns.TabIndex = 1
    Me.tbpMultImpRuns.Text = "Multiple Import Runs"
    Me.tbpMultImpRuns.UseVisualStyleBackColor = True
    '
    'lblSelectedImportType
    '
    Me.lblSelectedImportType.AutoSize = True
    Me.lblSelectedImportType.Location = New System.Drawing.Point(441, 31)
    Me.lblSelectedImportType.Name = "lblSelectedImportType"
    Me.lblSelectedImportType.Size = New System.Drawing.Size(149, 17)
    Me.lblSelectedImportType.TabIndex = 7
    Me.lblSelectedImportType.Text = "Selected Import Types"
    '
    'chkDupAsError
    '
    Me.chkDupAsError.AutoSize = True
    Me.chkDupAsError.Location = New System.Drawing.Point(15, 295)
    Me.chkDupAsError.Name = "chkDupAsError"
    Me.chkDupAsError.Size = New System.Drawing.Size(336, 21)
    Me.chkDupAsError.TabIndex = 6
    Me.chkDupAsError.Text = "Process subsequent loads for duplicate contacts"
    Me.chkDupAsError.UseVisualStyleBackColor = True
    '
    'lstMultiImportSelected
    '
    Me.lstMultiImportSelected.FormattingEnabled = True
    Me.lstMultiImportSelected.ItemHeight = 16
    Me.lstMultiImportSelected.Location = New System.Drawing.Point(441, 62)
    Me.lstMultiImportSelected.Name = "lstMultiImportSelected"
    Me.lstMultiImportSelected.Size = New System.Drawing.Size(275, 212)
    Me.lstMultiImportSelected.TabIndex = 5
    '
    'cmdRemoveDefinition
    '
    Me.cmdRemoveDefinition.Location = New System.Drawing.Point(314, 174)
    Me.cmdRemoveDefinition.Name = "cmdRemoveDefinition"
    Me.cmdRemoveDefinition.Size = New System.Drawing.Size(96, 27)
    Me.cmdRemoveDefinition.TabIndex = 4
    Me.cmdRemoveDefinition.Text = "<< Remove"
    Me.cmdRemoveDefinition.UseVisualStyleBackColor = True
    '
    'cmdAddDefinition
    '
    Me.cmdAddDefinition.Enabled = False
    Me.cmdAddDefinition.Location = New System.Drawing.Point(314, 131)
    Me.cmdAddDefinition.Name = "cmdAddDefinition"
    Me.cmdAddDefinition.Size = New System.Drawing.Size(96, 27)
    Me.cmdAddDefinition.TabIndex = 3
    Me.cmdAddDefinition.Text = "Add >>"
    Me.cmdAddDefinition.UseVisualStyleBackColor = True
    '
    'lblAvailableImportTypes
    '
    Me.lblAvailableImportTypes.AutoSize = True
    Me.lblAvailableImportTypes.Location = New System.Drawing.Point(12, 31)
    Me.lblAvailableImportTypes.Name = "lblAvailableImportTypes"
    Me.lblAvailableImportTypes.Size = New System.Drawing.Size(151, 17)
    Me.lblAvailableImportTypes.TabIndex = 1
    Me.lblAvailableImportTypes.Text = "Available Import Types"
    '
    'lstMultiImportAvailable
    '
    Me.lstMultiImportAvailable.FormattingEnabled = True
    Me.lstMultiImportAvailable.ItemHeight = 16
    Me.lstMultiImportAvailable.Location = New System.Drawing.Point(15, 62)
    Me.lstMultiImportAvailable.Name = "lstMultiImportAvailable"
    Me.lstMultiImportAvailable.Size = New System.Drawing.Size(275, 212)
    Me.lstMultiImportAvailable.TabIndex = 0
    '
    'tbpCustomOpt
    '
    Me.tbpCustomOpt.Controls.Add(Me.pnlConAndOrg)
    Me.tbpCustomOpt.Controls.Add(Me.pnlTableImport)
    Me.tbpCustomOpt.Controls.Add(Me.pnlPayment)
    Me.tbpCustomOpt.Controls.Add(Me.pnlDocument)
    Me.tbpCustomOpt.Controls.Add(Me.pnlStock)
    Me.tbpCustomOpt.Controls.Add(Me.pnlBankTransactions)
    Me.tbpCustomOpt.Controls.Add(Me.pnlAddrUpdate)
    Me.tbpCustomOpt.Location = New System.Drawing.Point(4, 26)
    Me.tbpCustomOpt.Name = "tbpCustomOpt"
    Me.tbpCustomOpt.Size = New System.Drawing.Size(192, 70)
    Me.tbpCustomOpt.TabIndex = 2
    Me.tbpCustomOpt.UseVisualStyleBackColor = True
    '
    'pnlConAndOrg
    '
    Me.pnlConAndOrg.Controls.Add(Me.grpDataOption)
    Me.pnlConAndOrg.Controls.Add(Me.grpDupUpdate)
    Me.pnlConAndOrg.Controls.Add(Me.grpDeDuplication)
    Me.pnlConAndOrg.Location = New System.Drawing.Point(0, 0)
    Me.pnlConAndOrg.Name = "pnlConAndOrg"
    Me.pnlConAndOrg.Size = New System.Drawing.Size(906, 537)
    Me.pnlConAndOrg.TabIndex = 18
    '
    'grpDataOption
    '
    Me.grpDataOption.Controls.Add(Me.chkAllowBlankForOrgName)
    Me.grpDataOption.Controls.Add(Me.txtOrgNumber)
    Me.grpDataOption.Controls.Add(Me.lblOrganisation)
    Me.grpDataOption.Controls.Add(Me.chkAddPosition)
    Me.grpDataOption.Controls.Add(Me.chkEmployee)
    Me.grpDataOption.Controls.Add(Me.chkCacheMailsort)
    Me.grpDataOption.Controls.Add(Me.chkDefAddrFromUnknown)
    Me.grpDataOption.Controls.Add(Me.chkDefSupp)
    Me.grpDataOption.Controls.Add(Me.chkCreateGridRefs)
    Me.grpDataOption.Controls.Add(Me.chkRePostcode)
    Me.grpDataOption.Controls.Add(Me.chkPAFAddress)
    Me.grpDataOption.Controls.Add(Me.chkSurnameFirst)
    Me.grpDataOption.Controls.Add(Me.chkCaps)
    Me.grpDataOption.Controls.Add(Me.chkDear)
    Me.grpDataOption.Location = New System.Drawing.Point(2, 310)
    Me.grpDataOption.Name = "grpDataOption"
    Me.grpDataOption.Size = New System.Drawing.Size(896, 227)
    Me.grpDataOption.TabIndex = 2
    Me.grpDataOption.TabStop = False
    Me.grpDataOption.Text = "Data Options"
    '
    'chkAllowBlankForOrgName
    '
    Me.chkAllowBlankForOrgName.AutoSize = True
    Me.chkAllowBlankForOrgName.Location = New System.Drawing.Point(15, 190)
    Me.chkAllowBlankForOrgName.Name = "chkAllowBlankForOrgName"
    Me.chkAllowBlankForOrgName.Size = New System.Drawing.Size(226, 21)
    Me.chkAllowBlankForOrgName.TabIndex = 17
    Me.chkAllowBlankForOrgName.Text = "Allow blank Organisation Name"
    Me.chkAllowBlankForOrgName.UseVisualStyleBackColor = True
    '
    'txtOrgNumber
    '
    Me.txtOrgNumber.ActiveOnly = False
    Me.txtOrgNumber.BackColor = System.Drawing.Color.Transparent
    Me.txtOrgNumber.CustomFormNumber = 0
    Me.txtOrgNumber.Description = ""
    Me.txtOrgNumber.Enabled = False
    Me.txtOrgNumber.EnabledProperty = True
    Me.txtOrgNumber.ExamCentreId = 0
    Me.txtOrgNumber.ExamCentreUnitId = 0
    Me.txtOrgNumber.ExamUnitLinkId = 0
    Me.txtOrgNumber.HasDependancies = False
    Me.txtOrgNumber.IsDesign = False
    Me.txtOrgNumber.Location = New System.Drawing.Point(318, 167)
    Me.txtOrgNumber.MaxLength = 32767
    Me.txtOrgNumber.MultipleValuesSupported = False
    Me.txtOrgNumber.Name = "txtOrgNumber"
    Me.txtOrgNumber.OriginalText = Nothing
    Me.txtOrgNumber.PreventHistoricalSelection = False
    Me.txtOrgNumber.ReadOnlyProperty = False
    Me.txtOrgNumber.Size = New System.Drawing.Size(408, 24)
    Me.txtOrgNumber.TabIndex = 24
    Me.txtOrgNumber.TextReadOnly = False
    Me.txtOrgNumber.TotalWidth = 408
    Me.txtOrgNumber.ValidationRequired = True
    Me.txtOrgNumber.WarningMessage = Nothing
    '
    'lblOrganisation
    '
    Me.lblOrganisation.AutoSize = True
    Me.lblOrganisation.Location = New System.Drawing.Point(335, 197)
    Me.lblOrganisation.Name = "lblOrganisation"
    Me.lblOrganisation.Size = New System.Drawing.Size(0, 17)
    Me.lblOrganisation.TabIndex = 23
    '
    'chkAddPosition
    '
    Me.chkAddPosition.AutoSize = True
    Me.chkAddPosition.Location = New System.Drawing.Point(317, 112)
    Me.chkAddPosition.Name = "chkAddPosition"
    Me.chkAddPosition.Size = New System.Drawing.Size(350, 21)
    Me.chkAddPosition.TabIndex = 21
    Me.chkAddPosition.Text = "Add position if it does not exist(Change of Position)"
    Me.chkAddPosition.UseVisualStyleBackColor = True
    '
    'chkEmployee
    '
    Me.chkEmployee.AutoSize = True
    Me.chkEmployee.Location = New System.Drawing.Point(317, 139)
    Me.chkEmployee.Name = "chkEmployee"
    Me.chkEmployee.Size = New System.Drawing.Size(352, 21)
    Me.chkEmployee.TabIndex = 22
    Me.chkEmployee.Text = "Employee load or update for Organisation Number:"
    Me.chkEmployee.UseVisualStyleBackColor = True
    '
    'chkCacheMailsort
    '
    Me.chkCacheMailsort.AutoSize = True
    Me.chkCacheMailsort.Checked = True
    Me.chkCacheMailsort.CheckState = System.Windows.Forms.CheckState.Checked
    Me.chkCacheMailsort.Location = New System.Drawing.Point(317, 85)
    Me.chkCacheMailsort.Name = "chkCacheMailsort"
    Me.chkCacheMailsort.Size = New System.Drawing.Size(155, 21)
    Me.chkCacheMailsort.TabIndex = 20
    Me.chkCacheMailsort.Text = "Cache Mailsort data"
    Me.chkCacheMailsort.UseVisualStyleBackColor = True
    '
    'chkDefAddrFromUnknown
    '
    Me.chkDefAddrFromUnknown.AutoSize = True
    Me.chkDefAddrFromUnknown.Location = New System.Drawing.Point(317, 58)
    Me.chkDefAddrFromUnknown.Name = "chkDefAddrFromUnknown"
    Me.chkDefAddrFromUnknown.Size = New System.Drawing.Size(398, 21)
    Me.chkDefAddrFromUnknown.TabIndex = 19
    Me.chkDefAddrFromUnknown.Text = "Default Address from Unknown Address (for new contacts)"
    Me.chkDefAddrFromUnknown.UseVisualStyleBackColor = True
    '
    'chkDefSupp
    '
    Me.chkDefSupp.AutoSize = True
    Me.chkDefSupp.Checked = True
    Me.chkDefSupp.CheckState = System.Windows.Forms.CheckState.Checked
    Me.chkDefSupp.Location = New System.Drawing.Point(317, 31)
    Me.chkDefSupp.Name = "chkDefSupp"
    Me.chkDefSupp.Size = New System.Drawing.Size(290, 21)
    Me.chkDefSupp.TabIndex = 18
    Me.chkDefSupp.Text = "Add default suppression for new contacts"
    Me.chkDefSupp.UseVisualStyleBackColor = True
    '
    'chkCreateGridRefs
    '
    Me.chkCreateGridRefs.AutoSize = True
    Me.chkCreateGridRefs.Enabled = False
    Me.chkCreateGridRefs.Location = New System.Drawing.Point(15, 166)
    Me.chkCreateGridRefs.Name = "chkCreateGridRefs"
    Me.chkCreateGridRefs.Size = New System.Drawing.Size(180, 21)
    Me.chkCreateGridRefs.TabIndex = 16
    Me.chkCreateGridRefs.Text = "Create Grid References"
    Me.chkCreateGridRefs.UseVisualStyleBackColor = True
    '
    'chkRePostcode
    '
    Me.chkRePostcode.AutoSize = True
    Me.chkRePostcode.Location = New System.Drawing.Point(15, 139)
    Me.chkRePostcode.Name = "chkRePostcode"
    Me.chkRePostcode.Size = New System.Drawing.Size(137, 21)
    Me.chkRePostcode.TabIndex = 15
    Me.chkRePostcode.Text = "Allow repostcode"
    Me.chkRePostcode.UseVisualStyleBackColor = True
    '
    'chkPAFAddress
    '
    Me.chkPAFAddress.AutoSize = True
    Me.chkPAFAddress.Location = New System.Drawing.Point(15, 112)
    Me.chkPAFAddress.Name = "chkPAFAddress"
    Me.chkPAFAddress.Size = New System.Drawing.Size(126, 21)
    Me.chkPAFAddress.TabIndex = 14
    Me.chkPAFAddress.Text = "PAF addresses"
    Me.chkPAFAddress.UseVisualStyleBackColor = True
    '
    'chkSurnameFirst
    '
    Me.chkSurnameFirst.AutoSize = True
    Me.chkSurnameFirst.Location = New System.Drawing.Point(15, 85)
    Me.chkSurnameFirst.Name = "chkSurnameFirst"
    Me.chkSurnameFirst.Size = New System.Drawing.Size(114, 21)
    Me.chkSurnameFirst.TabIndex = 13
    Me.chkSurnameFirst.Text = "Surname first"
    Me.chkSurnameFirst.UseVisualStyleBackColor = True
    '
    'chkCaps
    '
    Me.chkCaps.AutoSize = True
    Me.chkCaps.Location = New System.Drawing.Point(15, 58)
    Me.chkCaps.Name = "chkCaps"
    Me.chkCaps.Size = New System.Drawing.Size(241, 21)
    Me.chkCaps.TabIndex = 12
    Me.chkCaps.Text = "No capitalisation of imported data"
    Me.chkCaps.UseVisualStyleBackColor = True
    '
    'chkDear
    '
    Me.chkDear.AutoSize = True
    Me.chkDear.Checked = True
    Me.chkDear.CheckState = System.Windows.Forms.CheckState.Checked
    Me.chkDear.Location = New System.Drawing.Point(15, 31)
    Me.chkDear.Name = "chkDear"
    Me.chkDear.Size = New System.Drawing.Size(235, 21)
    Me.chkDear.TabIndex = 11
    Me.chkDear.Text = "Ensure Salutation prefix is 'Dear'"
    Me.chkDear.UseVisualStyleBackColor = True
    '
    'grpDupUpdate
    '
    Me.grpDupUpdate.Controls.Add(Me.lblUpdateSub)
    Me.grpDupUpdate.Controls.Add(Me.chkNameGatheringIncentives)
    Me.grpDupUpdate.Controls.Add(Me.chkActivity)
    Me.grpDupUpdate.Controls.Add(Me.chkUpdateWithNull)
    Me.grpDupUpdate.Controls.Add(Me.chkUpdateAll)
    Me.grpDupUpdate.Controls.Add(Me.chkUpdate)
    Me.grpDupUpdate.Location = New System.Drawing.Point(437, 3)
    Me.grpDupUpdate.Name = "grpDupUpdate"
    Me.grpDupUpdate.Size = New System.Drawing.Size(463, 301)
    Me.grpDupUpdate.TabIndex = 1
    Me.grpDupUpdate.TabStop = False
    Me.grpDupUpdate.Text = "Duplicate Update Options"
    '
    'lblUpdateSub
    '
    Me.lblUpdateSub.Location = New System.Drawing.Point(15, 29)
    Me.lblUpdateSub.Name = "lblUpdateSub"
    Me.lblUpdateSub.Size = New System.Drawing.Size(371, 35)
    Me.lblUpdateSub.TabIndex = 15
    Me.lblUpdateSub.Text = "Update existing records with additional information from import file:"
    '
    'chkNameGatheringIncentives
    '
    Me.chkNameGatheringIncentives.AutoSize = True
    Me.chkNameGatheringIncentives.Location = New System.Drawing.Point(18, 191)
    Me.chkNameGatheringIncentives.Name = "chkNameGatheringIncentives"
    Me.chkNameGatheringIncentives.Size = New System.Drawing.Size(265, 21)
    Me.chkNameGatheringIncentives.TabIndex = 14
    Me.chkNameGatheringIncentives.Text = "Generate Name Gathering Incentives"
    Me.chkNameGatheringIncentives.UseVisualStyleBackColor = True
    '
    'chkActivity
    '
    Me.chkActivity.AutoSize = True
    Me.chkActivity.Location = New System.Drawing.Point(18, 164)
    Me.chkActivity.Name = "chkActivity"
    Me.chkActivity.Size = New System.Drawing.Size(247, 21)
    Me.chkActivity.TabIndex = 13
    Me.chkActivity.Text = "Only add activity if it does not exist"
    Me.chkActivity.UseVisualStyleBackColor = True
    '
    'chkUpdateWithNull
    '
    Me.chkUpdateWithNull.AutoSize = True
    Me.chkUpdateWithNull.Location = New System.Drawing.Point(18, 137)
    Me.chkUpdateWithNull.Name = "chkUpdateWithNull"
    Me.chkUpdateWithNull.Size = New System.Drawing.Size(179, 21)
    Me.chkUpdateWithNull.TabIndex = 12
    Me.chkUpdateWithNull.Text = "Update with Null Values"
    Me.chkUpdateWithNull.UseVisualStyleBackColor = True
    '
    'chkUpdateAll
    '
    Me.chkUpdateAll.AutoSize = True
    Me.chkUpdateAll.Location = New System.Drawing.Point(18, 110)
    Me.chkUpdateAll.Name = "chkUpdateAll"
    Me.chkUpdateAll.Size = New System.Drawing.Size(106, 21)
    Me.chkUpdateAll.TabIndex = 11
    Me.chkUpdateAll.Text = "For all fields"
    Me.chkUpdateAll.UseVisualStyleBackColor = True
    '
    'chkUpdate
    '
    Me.chkUpdate.AutoSize = True
    Me.chkUpdate.Enabled = False
    Me.chkUpdate.Location = New System.Drawing.Point(18, 83)
    Me.chkUpdate.Name = "chkUpdate"
    Me.chkUpdate.Size = New System.Drawing.Size(156, 21)
    Me.chkUpdate.TabIndex = 10
    Me.chkUpdate.Text = "For blank fields only"
    Me.chkUpdate.UseVisualStyleBackColor = True
    '
    'grpDeDuplication
    '
    Me.grpDeDuplication.Controls.Add(Me.chkOrgNamePostCodeAddress)
    Me.grpDeDuplication.Controls.Add(Me.chkBankDetailsDedup)
    Me.grpDeDuplication.Controls.Add(Me.chkOrgAddressPotDup)
    Me.grpDeDuplication.Controls.Add(Me.chkSoundexDedup)
    Me.grpDeDuplication.Controls.Add(Me.chkAddressDedup)
    Me.grpDeDuplication.Controls.Add(Me.chkForeInitDeDup)
    Me.grpDeDuplication.Controls.Add(Me.chkTitleDeDup)
    Me.grpDeDuplication.Controls.Add(Me.chkEmailDedup)
    Me.grpDeDuplication.Controls.Add(Me.chkNumberDeDup)
    Me.grpDeDuplication.Controls.Add(Me.chkExtRefDeDup)
    Me.grpDeDuplication.Controls.Add(Me.chkExclUnkAdd)
    Me.grpDeDuplication.Location = New System.Drawing.Point(8, 3)
    Me.grpDeDuplication.Name = "grpDeDuplication"
    Me.grpDeDuplication.Size = New System.Drawing.Size(423, 301)
    Me.grpDeDuplication.TabIndex = 0
    Me.grpDeDuplication.TabStop = False
    Me.grpDeDuplication.Text = "De-Duplication"
    '
    'chkOrgNamePostCodeAddress
    '
    Me.chkOrgNamePostCodeAddress.Location = New System.Drawing.Point(15, 215)
    Me.chkOrgNamePostCodeAddress.Name = "chkOrgNamePostCodeAddress"
    Me.chkOrgNamePostCodeAddress.Size = New System.Drawing.Size(402, 32)
    Me.chkOrgNamePostCodeAddress.TabIndex = 16
    Me.chkOrgNamePostCodeAddress.Text = "Mark organisation as duplicate only if Name and Postcode and/or Address Line 1 ex" &
    "ists in the database"
    Me.chkOrgNamePostCodeAddress.UseVisualStyleBackColor = True
    '
    'chkBankDetailsDedup
    '
    Me.chkBankDetailsDedup.AutoSize = True
    Me.chkBankDetailsDedup.Checked = True
    Me.chkBankDetailsDedup.CheckState = System.Windows.Forms.CheckState.Checked
    Me.chkBankDetailsDedup.Location = New System.Drawing.Point(15, 81)
    Me.chkBankDetailsDedup.Name = "chkBankDetailsDedup"
    Me.chkBankDetailsDedup.Size = New System.Drawing.Size(149, 21)
    Me.chkBankDetailsDedup.TabIndex = 11
    Me.chkBankDetailsDedup.Text = "Using Bank Details"
    Me.chkBankDetailsDedup.UseVisualStyleBackColor = True
    '
    'chkOrgAddressPotDup
    '
    Me.chkOrgAddressPotDup.Location = New System.Drawing.Point(15, 188)
    Me.chkOrgAddressPotDup.Name = "chkOrgAddressPotDup"
    Me.chkOrgAddressPotDup.Size = New System.Drawing.Size(363, 22)
    Me.chkOrgAddressPotDup.TabIndex = 15
    Me.chkOrgAddressPotDup.Text = "Mark organisation as duplicate if address exists in the database"
    Me.chkOrgAddressPotDup.UseVisualStyleBackColor = True
    '
    'chkSoundexDedup
    '
    Me.chkSoundexDedup.AutoSize = True
    Me.chkSoundexDedup.Checked = True
    Me.chkSoundexDedup.CheckState = System.Windows.Forms.CheckState.Checked
    Me.chkSoundexDedup.Location = New System.Drawing.Point(275, 81)
    Me.chkSoundexDedup.Name = "chkSoundexDedup"
    Me.chkSoundexDedup.Size = New System.Drawing.Size(125, 21)
    Me.chkSoundexDedup.TabIndex = 19
    Me.chkSoundexDedup.Text = "Using Soundex"
    Me.chkSoundexDedup.UseVisualStyleBackColor = True
    '
    'chkAddressDedup
    '
    Me.chkAddressDedup.AutoSize = True
    Me.chkAddressDedup.Checked = True
    Me.chkAddressDedup.CheckState = System.Windows.Forms.CheckState.Checked
    Me.chkAddressDedup.Location = New System.Drawing.Point(275, 56)
    Me.chkAddressDedup.Name = "chkAddressDedup"
    Me.chkAddressDedup.Size = New System.Drawing.Size(122, 21)
    Me.chkAddressDedup.TabIndex = 18
    Me.chkAddressDedup.Text = "Using Address"
    Me.chkAddressDedup.UseVisualStyleBackColor = True
    '
    'chkForeInitDeDup
    '
    Me.chkForeInitDeDup.AutoSize = True
    Me.chkForeInitDeDup.Checked = True
    Me.chkForeInitDeDup.CheckState = System.Windows.Forms.CheckState.Checked
    Me.chkForeInitDeDup.Location = New System.Drawing.Point(15, 134)
    Me.chkForeInitDeDup.Name = "chkForeInitDeDup"
    Me.chkForeInitDeDup.Size = New System.Drawing.Size(229, 21)
    Me.chkForeInitDeDup.TabIndex = 13
    Me.chkForeInitDeDup.Text = "Using Forenames and/or Initials"
    Me.chkForeInitDeDup.UseVisualStyleBackColor = True
    '
    'chkTitleDeDup
    '
    Me.chkTitleDeDup.AutoSize = True
    Me.chkTitleDeDup.Checked = True
    Me.chkTitleDeDup.CheckState = System.Windows.Forms.CheckState.Checked
    Me.chkTitleDeDup.Location = New System.Drawing.Point(275, 29)
    Me.chkTitleDeDup.Name = "chkTitleDeDup"
    Me.chkTitleDeDup.Size = New System.Drawing.Size(97, 21)
    Me.chkTitleDeDup.TabIndex = 17
    Me.chkTitleDeDup.Text = "Using Title"
    Me.chkTitleDeDup.UseVisualStyleBackColor = True
    '
    'chkEmailDedup
    '
    Me.chkEmailDedup.AutoSize = True
    Me.chkEmailDedup.Checked = True
    Me.chkEmailDedup.CheckState = System.Windows.Forms.CheckState.Checked
    Me.chkEmailDedup.Location = New System.Drawing.Point(15, 107)
    Me.chkEmailDedup.Name = "chkEmailDedup"
    Me.chkEmailDedup.Size = New System.Drawing.Size(175, 21)
    Me.chkEmailDedup.TabIndex = 12
    Me.chkEmailDedup.Text = "Using Email Addresses"
    Me.chkEmailDedup.UseVisualStyleBackColor = True
    '
    'chkNumberDeDup
    '
    Me.chkNumberDeDup.AutoSize = True
    Me.chkNumberDeDup.Checked = True
    Me.chkNumberDeDup.CheckState = System.Windows.Forms.CheckState.Checked
    Me.chkNumberDeDup.Location = New System.Drawing.Point(15, 56)
    Me.chkNumberDeDup.Name = "chkNumberDeDup"
    Me.chkNumberDeDup.Size = New System.Drawing.Size(257, 21)
    Me.chkNumberDeDup.TabIndex = 10
    Me.chkNumberDeDup.Text = "Using Contact/Organisation Number"
    Me.chkNumberDeDup.UseVisualStyleBackColor = True
    '
    'chkExtRefDeDup
    '
    Me.chkExtRefDeDup.AutoSize = True
    Me.chkExtRefDeDup.Checked = True
    Me.chkExtRefDeDup.CheckState = System.Windows.Forms.CheckState.Checked
    Me.chkExtRefDeDup.Location = New System.Drawing.Point(15, 29)
    Me.chkExtRefDeDup.Name = "chkExtRefDeDup"
    Me.chkExtRefDeDup.Size = New System.Drawing.Size(191, 21)
    Me.chkExtRefDeDup.TabIndex = 9
    Me.chkExtRefDeDup.Text = "Using External Reference"
    Me.chkExtRefDeDup.UseVisualStyleBackColor = True
    '
    'chkExclUnkAdd
    '
    Me.chkExclUnkAdd.AutoSize = True
    Me.chkExclUnkAdd.Checked = True
    Me.chkExclUnkAdd.CheckState = System.Windows.Forms.CheckState.Checked
    Me.chkExclUnkAdd.Location = New System.Drawing.Point(15, 161)
    Me.chkExclUnkAdd.Name = "chkExclUnkAdd"
    Me.chkExclUnkAdd.Size = New System.Drawing.Size(212, 21)
    Me.chkExclUnkAdd.TabIndex = 14
    Me.chkExclUnkAdd.Text = "Exclude Unknown Addresses"
    Me.chkExclUnkAdd.UseVisualStyleBackColor = True
    '
    'pnlTableImport
    '
    Me.pnlTableImport.Controls.Add(Me.chkEmptyBeforeImport)
    Me.pnlTableImport.Controls.Add(Me.grp)
    Me.pnlTableImport.Location = New System.Drawing.Point(0, 0)
    Me.pnlTableImport.Name = "pnlTableImport"
    Me.pnlTableImport.Size = New System.Drawing.Size(905, 540)
    Me.pnlTableImport.TabIndex = 1
    '
    'chkEmptyBeforeImport
    '
    Me.chkEmptyBeforeImport.AutoSize = True
    Me.chkEmptyBeforeImport.Location = New System.Drawing.Point(18, 56)
    Me.chkEmptyBeforeImport.Name = "chkEmptyBeforeImport"
    Me.chkEmptyBeforeImport.Size = New System.Drawing.Size(192, 21)
    Me.chkEmptyBeforeImport.TabIndex = 1
    Me.chkEmptyBeforeImport.Text = "Empty table before import"
    Me.chkEmptyBeforeImport.UseVisualStyleBackColor = True
    '
    'grp
    '
    Me.grp.Controls.Add(Me.lblTableImport)
    Me.grp.Controls.Add(Me.chkUpdateAllTableImport)
    Me.grp.Controls.Add(Me.chkUpdateTableImport)
    Me.grp.Location = New System.Drawing.Point(431, 15)
    Me.grp.Name = "grp"
    Me.grp.Size = New System.Drawing.Size(471, 492)
    Me.grp.TabIndex = 0
    Me.grp.TabStop = False
    Me.grp.Text = "Duplicate Update Options"
    '
    'lblTableImport
    '
    Me.lblTableImport.AutoSize = True
    Me.lblTableImport.Location = New System.Drawing.Point(28, 31)
    Me.lblTableImport.Name = "lblTableImport"
    Me.lblTableImport.Size = New System.Drawing.Size(425, 17)
    Me.lblTableImport.TabIndex = 2
    Me.lblTableImport.Text = "Update existing records with additional information from import file:"
    '
    'chkUpdateAllTableImport
    '
    Me.chkUpdateAllTableImport.AutoSize = True
    Me.chkUpdateAllTableImport.Location = New System.Drawing.Point(31, 106)
    Me.chkUpdateAllTableImport.Name = "chkUpdateAllTableImport"
    Me.chkUpdateAllTableImport.Size = New System.Drawing.Size(106, 21)
    Me.chkUpdateAllTableImport.TabIndex = 1
    Me.chkUpdateAllTableImport.Text = "For all fields"
    Me.chkUpdateAllTableImport.UseVisualStyleBackColor = True
    '
    'chkUpdateTableImport
    '
    Me.chkUpdateTableImport.AutoSize = True
    Me.chkUpdateTableImport.Location = New System.Drawing.Point(31, 79)
    Me.chkUpdateTableImport.Name = "chkUpdateTableImport"
    Me.chkUpdateTableImport.Size = New System.Drawing.Size(156, 21)
    Me.chkUpdateTableImport.TabIndex = 0
    Me.chkUpdateTableImport.Text = "For blank fields only"
    Me.chkUpdateTableImport.UseVisualStyleBackColor = True
    '
    'pnlPayment
    '
    Me.pnlPayment.Controls.Add(Me.lblNumberOfDays)
    Me.pnlPayment.Controls.Add(Me.txtNumberOfDays)
    Me.pnlPayment.Controls.Add(Me.chkMatchSchPayment)
    Me.pnlPayment.Controls.Add(Me.chkCreateAct)
    Me.pnlPayment.Controls.Add(Me.chkSkipZeroAmt)
    Me.pnlPayment.Controls.Add(Me.chkProcessIncentives)
    Me.pnlPayment.Controls.Add(Me.chkReference)
    Me.pnlPayment.Controls.Add(Me.chkAddTransactions)
    Me.pnlPayment.Controls.Add(Me.chkNoFromFile)
    Me.pnlPayment.Controls.Add(Me.chkGiftAidRecords)
    Me.pnlPayment.Controls.Add(Me.grpTypeOfPayment)
    Me.pnlPayment.Location = New System.Drawing.Point(0, 0)
    Me.pnlPayment.Name = "pnlPayment"
    Me.pnlPayment.Size = New System.Drawing.Size(900, 514)
    Me.pnlPayment.TabIndex = 2
    '
    'lblNumberOfDays
    '
    Me.lblNumberOfDays.AutoSize = True
    Me.lblNumberOfDays.Location = New System.Drawing.Point(473, 276)
    Me.lblNumberOfDays.Name = "lblNumberOfDays"
    Me.lblNumberOfDays.Size = New System.Drawing.Size(305, 17)
    Me.lblNumberOfDays.TabIndex = 10
    Me.lblNumberOfDays.Text = "Number of days either side of Transaction date"
    '
    'txtNumberOfDays
    '
    Me.txtNumberOfDays.Enabled = False
    Me.txtNumberOfDays.Location = New System.Drawing.Point(409, 272)
    Me.txtNumberOfDays.Name = "txtNumberOfDays"
    Me.txtNumberOfDays.Size = New System.Drawing.Size(45, 22)
    Me.txtNumberOfDays.TabIndex = 9
    '
    'chkMatchSchPayment
    '
    Me.chkMatchSchPayment.AutoSize = True
    Me.chkMatchSchPayment.Location = New System.Drawing.Point(409, 235)
    Me.chkMatchSchPayment.Name = "chkMatchSchPayment"
    Me.chkMatchSchPayment.Size = New System.Drawing.Size(474, 21)
    Me.chkMatchSchPayment.TabIndex = 8
    Me.chkMatchSchPayment.Text = "Match Pay Plan Payments to Scheduled Payment on Amount and Date"
    Me.chkMatchSchPayment.UseVisualStyleBackColor = True
    '
    'chkCreateAct
    '
    Me.chkCreateAct.AutoSize = True
    Me.chkCreateAct.Location = New System.Drawing.Point(409, 207)
    Me.chkCreateAct.Name = "chkCreateAct"
    Me.chkCreateAct.Size = New System.Drawing.Size(257, 21)
    Me.chkCreateAct.TabIndex = 7
    Me.chkCreateAct.Text = "Create activity for product payments"
    Me.chkCreateAct.UseVisualStyleBackColor = True
    '
    'chkSkipZeroAmt
    '
    Me.chkSkipZeroAmt.AutoSize = True
    Me.chkSkipZeroAmt.Location = New System.Drawing.Point(409, 179)
    Me.chkSkipZeroAmt.Name = "chkSkipZeroAmt"
    Me.chkSkipZeroAmt.Size = New System.Drawing.Size(208, 21)
    Me.chkSkipZeroAmt.TabIndex = 6
    Me.chkSkipZeroAmt.Text = "Skip lines with zero amounts"
    Me.chkSkipZeroAmt.UseVisualStyleBackColor = True
    '
    'chkProcessIncentives
    '
    Me.chkProcessIncentives.AutoSize = True
    Me.chkProcessIncentives.Location = New System.Drawing.Point(409, 151)
    Me.chkProcessIncentives.Name = "chkProcessIncentives"
    Me.chkProcessIncentives.Size = New System.Drawing.Size(148, 21)
    Me.chkProcessIncentives.TabIndex = 5
    Me.chkProcessIncentives.Text = "Process incentives"
    Me.chkProcessIncentives.UseVisualStyleBackColor = True
    '
    'chkReference
    '
    Me.chkReference.AutoSize = True
    Me.chkReference.Location = New System.Drawing.Point(409, 122)
    Me.chkReference.Name = "chkReference"
    Me.chkReference.Size = New System.Drawing.Size(388, 21)
    Me.chkReference.TabIndex = 4
    Me.chkReference.Text = "Default Reference to Batch Number/Transaction Number"
    Me.chkReference.UseVisualStyleBackColor = True
    '
    'chkAddTransactions
    '
    Me.chkAddTransactions.AutoSize = True
    Me.chkAddTransactions.Enabled = False
    Me.chkAddTransactions.Location = New System.Drawing.Point(409, 94)
    Me.chkAddTransactions.Name = "chkAddTransactions"
    Me.chkAddTransactions.Size = New System.Drawing.Size(257, 21)
    Me.chkAddTransactions.TabIndex = 3
    Me.chkAddTransactions.Text = "Add transactions to existing batches"
    Me.chkAddTransactions.UseVisualStyleBackColor = True
    '
    'chkNoFromFile
    '
    Me.chkNoFromFile.AutoSize = True
    Me.chkNoFromFile.Location = New System.Drawing.Point(409, 67)
    Me.chkNoFromFile.Name = "chkNoFromFile"
    Me.chkNoFromFile.Size = New System.Drawing.Size(352, 21)
    Me.chkNoFromFile.TabIndex = 2
    Me.chkNoFromFile.Text = "Use Batch, Transaction and Line Numbers from file"
    Me.chkNoFromFile.UseVisualStyleBackColor = True
    '
    'chkGiftAidRecords
    '
    Me.chkGiftAidRecords.AutoSize = True
    Me.chkGiftAidRecords.Location = New System.Drawing.Point(409, 40)
    Me.chkGiftAidRecords.Name = "chkGiftAidRecords"
    Me.chkGiftAidRecords.Size = New System.Drawing.Size(258, 21)
    Me.chkGiftAidRecords.TabIndex = 1
    Me.chkGiftAidRecords.Text = "Create unclaimed Gift Aid Donations"
    Me.chkGiftAidRecords.UseVisualStyleBackColor = True
    '
    'grpTypeOfPayment
    '
    Me.grpTypeOfPayment.Controls.Add(Me.optPaymentsUnposted)
    Me.grpTypeOfPayment.Controls.Add(Me.optPaymentsPostedToNominal)
    Me.grpTypeOfPayment.Controls.Add(Me.optPaymentsFinHistory)
    Me.grpTypeOfPayment.Controls.Add(Me.optPaymentsPostedToCB)
    Me.grpTypeOfPayment.Location = New System.Drawing.Point(6, 17)
    Me.grpTypeOfPayment.Name = "grpTypeOfPayment"
    Me.grpTypeOfPayment.Size = New System.Drawing.Size(386, 486)
    Me.grpTypeOfPayment.TabIndex = 0
    Me.grpTypeOfPayment.TabStop = False
    Me.grpTypeOfPayment.Text = "Type of Payments"
    '
    'optPaymentsUnposted
    '
    Me.optPaymentsUnposted.AutoSize = True
    Me.optPaymentsUnposted.Location = New System.Drawing.Point(20, 133)
    Me.optPaymentsUnposted.Name = "optPaymentsUnposted"
    Me.optPaymentsUnposted.Size = New System.Drawing.Size(176, 21)
    Me.optPaymentsUnposted.TabIndex = 4
    Me.optPaymentsUnposted.TabStop = True
    Me.optPaymentsUnposted.Text = "Unposted Transactions"
    Me.optPaymentsUnposted.UseVisualStyleBackColor = True
    '
    'optPaymentsPostedToNominal
    '
    Me.optPaymentsPostedToNominal.AutoSize = True
    Me.optPaymentsPostedToNominal.Location = New System.Drawing.Point(20, 106)
    Me.optPaymentsPostedToNominal.Name = "optPaymentsPostedToNominal"
    Me.optPaymentsPostedToNominal.Size = New System.Drawing.Size(230, 21)
    Me.optPaymentsPostedToNominal.TabIndex = 3
    Me.optPaymentsPostedToNominal.TabStop = True
    Me.optPaymentsPostedToNominal.Text = "Posted to Nominal Transactions"
    Me.optPaymentsPostedToNominal.UseVisualStyleBackColor = True
    '
    'optPaymentsFinHistory
    '
    Me.optPaymentsFinHistory.AutoSize = True
    Me.optPaymentsFinHistory.Checked = True
    Me.optPaymentsFinHistory.Location = New System.Drawing.Point(20, 52)
    Me.optPaymentsFinHistory.Name = "optPaymentsFinHistory"
    Me.optPaymentsFinHistory.Size = New System.Drawing.Size(166, 21)
    Me.optPaymentsFinHistory.TabIndex = 2
    Me.optPaymentsFinHistory.TabStop = True
    Me.optPaymentsFinHistory.Text = "Financial History Only"
    Me.optPaymentsFinHistory.UseVisualStyleBackColor = True
    '
    'optPaymentsPostedToCB
    '
    Me.optPaymentsPostedToCB.AutoSize = True
    Me.optPaymentsPostedToCB.Location = New System.Drawing.Point(21, 79)
    Me.optPaymentsPostedToCB.Name = "optPaymentsPostedToCB"
    Me.optPaymentsPostedToCB.Size = New System.Drawing.Size(243, 21)
    Me.optPaymentsPostedToCB.TabIndex = 1
    Me.optPaymentsPostedToCB.TabStop = True
    Me.optPaymentsPostedToCB.Text = "Posted to Cash BookTransactions"
    Me.optPaymentsPostedToCB.UseVisualStyleBackColor = True
    '
    'pnlDocument
    '
    Me.pnlDocument.Controls.Add(Me.grpDupUpdateOpt)
    Me.pnlDocument.Location = New System.Drawing.Point(0, 0)
    Me.pnlDocument.Name = "pnlDocument"
    Me.pnlDocument.Size = New System.Drawing.Size(904, 521)
    Me.pnlDocument.TabIndex = 5
    '
    'grpDupUpdateOpt
    '
    Me.grpDupUpdateOpt.Controls.Add(Me.chkUpdateAllDoc)
    Me.grpDupUpdateOpt.Controls.Add(Me.chkUpdateDoc)
    Me.grpDupUpdateOpt.Controls.Add(Me.lblUpdateExisting)
    Me.grpDupUpdateOpt.Location = New System.Drawing.Point(12, 17)
    Me.grpDupUpdateOpt.Name = "grpDupUpdateOpt"
    Me.grpDupUpdateOpt.Size = New System.Drawing.Size(885, 490)
    Me.grpDupUpdateOpt.TabIndex = 0
    Me.grpDupUpdateOpt.TabStop = False
    Me.grpDupUpdateOpt.Text = "Duplicate Update Options"
    '
    'chkUpdateAllDoc
    '
    Me.chkUpdateAllDoc.AutoSize = True
    Me.chkUpdateAllDoc.Location = New System.Drawing.Point(20, 105)
    Me.chkUpdateAllDoc.Name = "chkUpdateAllDoc"
    Me.chkUpdateAllDoc.Size = New System.Drawing.Size(106, 21)
    Me.chkUpdateAllDoc.TabIndex = 2
    Me.chkUpdateAllDoc.Text = "For all fields"
    Me.chkUpdateAllDoc.UseVisualStyleBackColor = True
    '
    'chkUpdateDoc
    '
    Me.chkUpdateDoc.AutoSize = True
    Me.chkUpdateDoc.Location = New System.Drawing.Point(20, 66)
    Me.chkUpdateDoc.Name = "chkUpdateDoc"
    Me.chkUpdateDoc.Size = New System.Drawing.Size(156, 21)
    Me.chkUpdateDoc.TabIndex = 1
    Me.chkUpdateDoc.Text = "For blank fields only"
    Me.chkUpdateDoc.UseVisualStyleBackColor = True
    '
    'lblUpdateExisting
    '
    Me.lblUpdateExisting.AutoSize = True
    Me.lblUpdateExisting.Location = New System.Drawing.Point(17, 36)
    Me.lblUpdateExisting.Name = "lblUpdateExisting"
    Me.lblUpdateExisting.Size = New System.Drawing.Size(425, 17)
    Me.lblUpdateExisting.TabIndex = 0
    Me.lblUpdateExisting.Tag = "p"
    Me.lblUpdateExisting.Text = "Update existing records with additional information from import file:"
    '
    'pnlStock
    '
    Me.pnlStock.Controls.Add(Me.optStockSet)
    Me.pnlStock.Controls.Add(Me.optStockUpdate)
    Me.pnlStock.Location = New System.Drawing.Point(-1, 0)
    Me.pnlStock.Name = "pnlStock"
    Me.pnlStock.Size = New System.Drawing.Size(901, 513)
    Me.pnlStock.TabIndex = 5
    '
    'optStockSet
    '
    Me.optStockSet.AutoSize = True
    Me.optStockSet.Location = New System.Drawing.Point(38, 56)
    Me.optStockSet.Name = "optStockSet"
    Me.optStockSet.Size = New System.Drawing.Size(134, 21)
    Me.optStockSet.TabIndex = 1
    Me.optStockSet.TabStop = True
    Me.optStockSet.Text = "Set Stock Levels"
    Me.optStockSet.UseVisualStyleBackColor = True
    '
    'optStockUpdate
    '
    Me.optStockUpdate.AutoSize = True
    Me.optStockUpdate.Location = New System.Drawing.Point(38, 29)
    Me.optStockUpdate.Name = "optStockUpdate"
    Me.optStockUpdate.Size = New System.Drawing.Size(159, 21)
    Me.optStockUpdate.TabIndex = 0
    Me.optStockUpdate.TabStop = True
    Me.optStockUpdate.Text = "Update Stock Levels"
    Me.optStockUpdate.UseVisualStyleBackColor = True
    '
    'pnlBankTransactions
    '
    Me.pnlBankTransactions.Controls.Add(Me.chkDASImport)
    Me.pnlBankTransactions.Location = New System.Drawing.Point(-1, 0)
    Me.pnlBankTransactions.Name = "pnlBankTransactions"
    Me.pnlBankTransactions.Size = New System.Drawing.Size(901, 513)
    Me.pnlBankTransactions.TabIndex = 5
    '
    'chkDASImport
    '
    Me.chkDASImport.AutoSize = True
    Me.chkDASImport.Location = New System.Drawing.Point(38, 29)
    Me.chkDASImport.Name = "chkDASImport"
    Me.chkDASImport.Size = New System.Drawing.Size(101, 21)
    Me.chkDASImport.TabIndex = 0
    Me.chkDASImport.Text = "DAS Import"
    Me.chkDASImport.UseVisualStyleBackColor = True
    '
    'pnlAddrUpdate
    '
    Me.pnlAddrUpdate.Controls.Add(Me.chkCacheMailsortAddr)
    Me.pnlAddrUpdate.Controls.Add(Me.chkExtractAddr)
    Me.pnlAddrUpdate.Location = New System.Drawing.Point(-4, 0)
    Me.pnlAddrUpdate.Name = "pnlAddrUpdate"
    Me.pnlAddrUpdate.Size = New System.Drawing.Size(904, 517)
    Me.pnlAddrUpdate.TabIndex = 18
    '
    'chkCacheMailsortAddr
    '
    Me.chkCacheMailsortAddr.AutoSize = True
    Me.chkCacheMailsortAddr.Checked = True
    Me.chkCacheMailsortAddr.CheckState = System.Windows.Forms.CheckState.Checked
    Me.chkCacheMailsortAddr.Location = New System.Drawing.Point(41, 56)
    Me.chkCacheMailsortAddr.Name = "chkCacheMailsortAddr"
    Me.chkCacheMailsortAddr.Size = New System.Drawing.Size(155, 21)
    Me.chkCacheMailsortAddr.TabIndex = 1
    Me.chkCacheMailsortAddr.Text = "Cache Mailsort data"
    Me.chkCacheMailsortAddr.UseVisualStyleBackColor = True
    '
    'chkExtractAddr
    '
    Me.chkExtractAddr.AutoSize = True
    Me.chkExtractAddr.Location = New System.Drawing.Point(41, 29)
    Me.chkExtractAddr.Name = "chkExtractAddr"
    Me.chkExtractAddr.Size = New System.Drawing.Size(268, 21)
    Me.chkExtractAddr.TabIndex = 0
    Me.chkExtractAddr.Text = "Extract original address records to file"
    Me.chkExtractAddr.UseVisualStyleBackColor = True
    '
    'tbpDefaults
    '
    Me.tbpDefaults.Controls.Add(Me.dgrDefaults)
    Me.tbpDefaults.Controls.Add(Me.pnlDefaults)
    Me.tbpDefaults.Location = New System.Drawing.Point(4, 26)
    Me.tbpDefaults.Name = "tbpDefaults"
    Me.tbpDefaults.Size = New System.Drawing.Size(192, 70)
    Me.tbpDefaults.TabIndex = 2
    Me.tbpDefaults.Text = "Defaults"
    Me.tbpDefaults.UseVisualStyleBackColor = True
    '
    'dgrDefaults
    '
    Me.dgrDefaults.AccessibleName = "Display Grid"
    Me.dgrDefaults.ActiveColumn = 0
    Me.dgrDefaults.AllowColumnResize = True
    Me.dgrDefaults.AllowSorting = True
    Me.dgrDefaults.AutoSetHeight = False
    Me.dgrDefaults.AutoSetRowHeight = False
    Me.dgrDefaults.DisplayTitle = Nothing
    Me.dgrDefaults.Dock = System.Windows.Forms.DockStyle.Fill
    Me.dgrDefaults.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!)
    Me.dgrDefaults.Location = New System.Drawing.Point(0, 0)
    Me.dgrDefaults.MaintenanceDesc = Nothing
    Me.dgrDefaults.MaxGridRows = 6
    Me.dgrDefaults.MultipleSelect = False
    Me.dgrDefaults.Name = "dgrDefaults"
    Me.dgrDefaults.RowCount = 10
    Me.dgrDefaults.ShowIfEmpty = False
    Me.dgrDefaults.Size = New System.Drawing.Size(192, 0)
    Me.dgrDefaults.SuppressHyperLinkFormat = False
    Me.dgrDefaults.TabIndex = 0
    '
    'pnlDefaults
    '
    Me.pnlDefaults.Controls.Add(Me.txtLookupDefValue)
    Me.pnlDefaults.Controls.Add(Me.lblElse)
    Me.pnlDefaults.Controls.Add(Me.dtpckValue)
    Me.pnlDefaults.Controls.Add(Me.cmdDefaultAdd)
    Me.pnlDefaults.Controls.Add(Me.cboPatternValue)
    Me.pnlDefaults.Controls.Add(Me.cboDefAttrs)
    Me.pnlDefaults.Controls.Add(Me.chkCtrlNo)
    Me.pnlDefaults.Controls.Add(Me.chkIncPerLine)
    Me.pnlDefaults.Controls.Add(Me.lblValue)
    Me.pnlDefaults.Controls.Add(Me.lblDefAttribute)
    Me.pnlDefaults.Controls.Add(Me.txtDefValue)
    Me.pnlDefaults.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.pnlDefaults.Location = New System.Drawing.Point(0, -48)
    Me.pnlDefaults.Name = "pnlDefaults"
    Me.pnlDefaults.Size = New System.Drawing.Size(192, 118)
    Me.pnlDefaults.TabIndex = 1
    '
    'txtLookupDefValue
    '
    Me.txtLookupDefValue.ActiveOnly = False
    Me.txtLookupDefValue.BackColor = System.Drawing.Color.Transparent
    Me.txtLookupDefValue.CustomFormNumber = 0
    Me.txtLookupDefValue.Description = ""
    Me.txtLookupDefValue.EnabledProperty = True
    Me.txtLookupDefValue.ExamCentreId = 0
    Me.txtLookupDefValue.ExamCentreUnitId = 0
    Me.txtLookupDefValue.ExamUnitLinkId = 0
    Me.txtLookupDefValue.HasDependancies = False
    Me.txtLookupDefValue.IsDesign = False
    Me.txtLookupDefValue.Location = New System.Drawing.Point(343, 36)
    Me.txtLookupDefValue.MaxLength = 32767
    Me.txtLookupDefValue.MultipleValuesSupported = False
    Me.txtLookupDefValue.Name = "txtLookupDefValue"
    Me.txtLookupDefValue.OriginalText = Nothing
    Me.txtLookupDefValue.PreventHistoricalSelection = False
    Me.txtLookupDefValue.ReadOnlyProperty = False
    Me.txtLookupDefValue.Size = New System.Drawing.Size(408, 24)
    Me.txtLookupDefValue.TabIndex = 12
    Me.txtLookupDefValue.TextReadOnly = False
    Me.txtLookupDefValue.TotalWidth = 408
    Me.txtLookupDefValue.ValidationRequired = True
    Me.txtLookupDefValue.WarningMessage = Nothing
    '
    'lblElse
    '
    Me.lblElse.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
    Me.lblElse.Location = New System.Drawing.Point(343, 36)
    Me.lblElse.Margin = New System.Windows.Forms.Padding(0)
    Me.lblElse.Name = "lblElse"
    Me.lblElse.Size = New System.Drawing.Size(301, 25)
    Me.lblElse.TabIndex = 10
    Me.lblElse.Text = "Cannot Be Defaulted"
    '
    'dtpckValue
    '
    Me.dtpckValue.CustomFormat = "dd/MM/yyyy"
    Me.dtpckValue.Location = New System.Drawing.Point(344, 37)
    Me.dtpckValue.Name = "dtpckValue"
    Me.dtpckValue.Size = New System.Drawing.Size(300, 22)
    Me.dtpckValue.TabIndex = 8
    '
    'cmdDefaultAdd
    '
    Me.cmdDefaultAdd.Location = New System.Drawing.Point(764, 35)
    Me.cmdDefaultAdd.Name = "cmdDefaultAdd"
    Me.cmdDefaultAdd.Size = New System.Drawing.Size(65, 24)
    Me.cmdDefaultAdd.TabIndex = 7
    Me.cmdDefaultAdd.Text = "Add"
    Me.cmdDefaultAdd.UseVisualStyleBackColor = True
    '
    'cboPatternValue
    '
    Me.cboPatternValue.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cboPatternValue.FormattingEnabled = True
    Me.cboPatternValue.Location = New System.Drawing.Point(344, 36)
    Me.cboPatternValue.Name = "cboPatternValue"
    Me.cboPatternValue.Size = New System.Drawing.Size(300, 24)
    Me.cboPatternValue.TabIndex = 5
    '
    'cboDefAttrs
    '
    Me.cboDefAttrs.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cboDefAttrs.FormattingEnabled = True
    Me.cboDefAttrs.Location = New System.Drawing.Point(3, 36)
    Me.cboDefAttrs.Name = "cboDefAttrs"
    Me.cboDefAttrs.Size = New System.Drawing.Size(262, 24)
    Me.cboDefAttrs.TabIndex = 4
    '
    'chkCtrlNo
    '
    Me.chkCtrlNo.AutoSize = True
    Me.chkCtrlNo.Location = New System.Drawing.Point(173, 66)
    Me.chkCtrlNo.Name = "chkCtrlNo"
    Me.chkCtrlNo.Size = New System.Drawing.Size(129, 21)
    Me.chkCtrlNo.TabIndex = 3
    Me.chkCtrlNo.Text = "Control Number"
    Me.chkCtrlNo.UseVisualStyleBackColor = True
    '
    'chkIncPerLine
    '
    Me.chkIncPerLine.AutoSize = True
    Me.chkIncPerLine.Location = New System.Drawing.Point(0, 66)
    Me.chkIncPerLine.Name = "chkIncPerLine"
    Me.chkIncPerLine.Size = New System.Drawing.Size(149, 21)
    Me.chkIncPerLine.TabIndex = 2
    Me.chkIncPerLine.Text = "Increment Per Line"
    Me.chkIncPerLine.UseVisualStyleBackColor = True
    '
    'lblValue
    '
    Me.lblValue.AutoSize = True
    Me.lblValue.Location = New System.Drawing.Point(343, 10)
    Me.lblValue.Name = "lblValue"
    Me.lblValue.Size = New System.Drawing.Size(48, 17)
    Me.lblValue.TabIndex = 1
    Me.lblValue.Text = "Value:"
    '
    'lblDefAttribute
    '
    Me.lblDefAttribute.AutoSize = True
    Me.lblDefAttribute.Location = New System.Drawing.Point(2, 10)
    Me.lblDefAttribute.Name = "lblDefAttribute"
    Me.lblDefAttribute.Size = New System.Drawing.Size(72, 17)
    Me.lblDefAttribute.TabIndex = 0
    Me.lblDefAttribute.Text = "Attributes:"
    '
    'txtDefValue
    '
    Me.txtDefValue.Location = New System.Drawing.Point(343, 36)
    Me.txtDefValue.Name = "txtDefValue"
    Me.txtDefValue.Size = New System.Drawing.Size(376, 22)
    Me.txtDefValue.TabIndex = 13
    '
    'SplitControl
    '
    Me.SplitControl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.SplitControl.Location = New System.Drawing.Point(0, -4)
    Me.SplitControl.Name = "SplitControl"
    Me.SplitControl.Orientation = System.Windows.Forms.Orientation.Horizontal
    '
    'SplitControl.Panel1
    '
    Me.SplitControl.Panel1.Controls.Add(Me.bpl)
    '
    'SplitControl.Panel2
    '
    Me.SplitControl.Panel2.Controls.Add(Me.lblJobNumber)
    Me.SplitControl.Panel2.Controls.Add(Me.lblStatus)
    Me.SplitControl.Size = New System.Drawing.Size(929, 72)
    Me.SplitControl.SplitterDistance = 43
    Me.SplitControl.TabIndex = 19
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdStop)
    Me.bpl.Controls.Add(Me.cmdTest)
    Me.bpl.Controls.Add(Me.cmdOk)
    Me.bpl.Controls.Add(Me.cmdCancel)
    Me.bpl.Controls.Add(Me.cmdSave)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bpl.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bpl.Location = New System.Drawing.Point(0, 4)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(929, 39)
    Me.bpl.TabIndex = 18
    '
    'cmdStop
    '
    Me.cmdStop.Location = New System.Drawing.Point(194, 6)
    Me.cmdStop.Name = "cmdStop"
    Me.cmdStop.Size = New System.Drawing.Size(96, 27)
    Me.cmdStop.TabIndex = 6
    Me.cmdStop.Text = "Stop Process"
    Me.cmdStop.UseVisualStyleBackColor = True
    Me.cmdStop.Visible = False
    '
    'cmdTest
    '
    Me.cmdTest.Location = New System.Drawing.Point(305, 6)
    Me.cmdTest.Name = "cmdTest"
    Me.cmdTest.Size = New System.Drawing.Size(96, 27)
    Me.cmdTest.TabIndex = 2
    Me.cmdTest.Text = "Test"
    Me.cmdTest.UseVisualStyleBackColor = True
    '
    'cmdOk
    '
    Me.cmdOk.Location = New System.Drawing.Point(416, 6)
    Me.cmdOk.Name = "cmdOk"
    Me.cmdOk.Size = New System.Drawing.Size(96, 27)
    Me.cmdOk.TabIndex = 4
    Me.cmdOk.Text = "OK"
    Me.cmdOk.UseVisualStyleBackColor = True
    '
    'cmdCancel
    '
    Me.cmdCancel.Location = New System.Drawing.Point(527, 6)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(96, 27)
    Me.cmdCancel.TabIndex = 3
    Me.cmdCancel.Text = "Cancel"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'cmdSave
    '
    Me.cmdSave.Location = New System.Drawing.Point(638, 6)
    Me.cmdSave.Name = "cmdSave"
    Me.cmdSave.Size = New System.Drawing.Size(96, 27)
    Me.cmdSave.TabIndex = 1
    Me.cmdSave.Text = "Save Settings"
    Me.cmdSave.UseVisualStyleBackColor = True
    Me.cmdSave.Visible = False
    '
    'lblJobNumber
    '
    Me.lblJobNumber.AutoSize = True
    Me.lblJobNumber.Location = New System.Drawing.Point(3, 9)
    Me.lblJobNumber.Name = "lblJobNumber"
    Me.lblJobNumber.Size = New System.Drawing.Size(0, 17)
    Me.lblJobNumber.TabIndex = 1
    '
    'lblStatus
    '
    Me.lblStatus.AutoSize = True
    Me.lblStatus.Location = New System.Drawing.Point(48, 9)
    Me.lblStatus.Name = "lblStatus"
    Me.lblStatus.Size = New System.Drawing.Size(0, 17)
    Me.lblStatus.TabIndex = 0
    '
    'frmImport
    '
    Me.ClientSize = New System.Drawing.Size(929, 675)
    Me.Controls.Add(Me.SplitContainer1)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmImport"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    Me.Text = "Import File"
    Me.dgrMenuStrip.ResumeLayout(False)
    Me.SplitContainer1.Panel1.ResumeLayout(False)
    Me.SplitContainer1.Panel2.ResumeLayout(False)
    CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
    Me.SplitContainer1.ResumeLayout(False)
    Me.tabMain.ResumeLayout(False)
    Me.tbpData.ResumeLayout(False)
    Me.tbpData.PerformLayout()
    Me.grpDedup.ResumeLayout(False)
    Me.grpDedup.PerformLayout()
    Me.tbpOptions.ResumeLayout(False)
    Me.tabSub.ResumeLayout(False)
    Me.tbpGeneralOpt.ResumeLayout(False)
    Me.spltGeneralOpt.Panel1.ResumeLayout(False)
    Me.spltGeneralOpt.Panel2.ResumeLayout(False)
    CType(Me.spltGeneralOpt, System.ComponentModel.ISupportInitialize).EndInit()
    Me.spltGeneralOpt.ResumeLayout(False)
    Me.grpProcessingOptions.ResumeLayout(False)
    Me.grpProcessingOptions.PerformLayout()
    Me.grpLogFileOptions.ResumeLayout(False)
    Me.grpLogFileOptions.PerformLayout()
    Me.tbpMultImpRuns.ResumeLayout(False)
    Me.tbpMultImpRuns.PerformLayout()
    Me.tbpCustomOpt.ResumeLayout(False)
    Me.pnlConAndOrg.ResumeLayout(False)
    Me.grpDataOption.ResumeLayout(False)
    Me.grpDataOption.PerformLayout()
    Me.grpDupUpdate.ResumeLayout(False)
    Me.grpDupUpdate.PerformLayout()
    Me.grpDeDuplication.ResumeLayout(False)
    Me.grpDeDuplication.PerformLayout()
    Me.pnlTableImport.ResumeLayout(False)
    Me.pnlTableImport.PerformLayout()
    Me.grp.ResumeLayout(False)
    Me.grp.PerformLayout()
    Me.pnlPayment.ResumeLayout(False)
    Me.pnlPayment.PerformLayout()
    Me.grpTypeOfPayment.ResumeLayout(False)
    Me.grpTypeOfPayment.PerformLayout()
    Me.pnlDocument.ResumeLayout(False)
    Me.grpDupUpdateOpt.ResumeLayout(False)
    Me.grpDupUpdateOpt.PerformLayout()
    Me.pnlStock.ResumeLayout(False)
    Me.pnlStock.PerformLayout()
    Me.pnlBankTransactions.ResumeLayout(False)
    Me.pnlBankTransactions.PerformLayout()
    Me.pnlAddrUpdate.ResumeLayout(False)
    Me.pnlAddrUpdate.PerformLayout()
    Me.tbpDefaults.ResumeLayout(False)
    Me.pnlDefaults.ResumeLayout(False)
    Me.pnlDefaults.PerformLayout()
    Me.SplitControl.Panel1.ResumeLayout(False)
    Me.SplitControl.Panel2.ResumeLayout(False)
    Me.SplitControl.Panel2.PerformLayout()
    CType(Me.SplitControl, System.ComponentModel.ISupportInitialize).EndInit()
    Me.SplitControl.ResumeLayout(False)
    Me.bpl.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub

  Friend WithEvents pnlAddressUpdate As System.Windows.Forms.Panel

  'Friend WithEvents chk As System.Windows.Forms.CheckBox
  'Friend WithEvents Panel1 As System.Windows.Forms.Panel
  Friend WithEvents dgrMenuStrip As System.Windows.Forms.ContextMenuStrip
  Friend WithEvents dgrMenuDateFormat As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents dgrMenuMapAttribute As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
  Friend WithEvents tabMain As CDBNETCL.TabControl
  Friend WithEvents tbpData As System.Windows.Forms.TabPage
  Friend WithEvents dgr As CDBNETCL.DisplayGrid
  Friend WithEvents lblDataSource As System.Windows.Forms.Label
  Friend WithEvents lblSource As System.Windows.Forms.Label
  Friend WithEvents txtDataSource As CDBNETCL.TextLookupBox
  Friend WithEvents txtSource As CDBNETCL.TextLookupBox
  Friend WithEvents grpDedup As System.Windows.Forms.GroupBox
  Friend WithEvents optDedupNone As System.Windows.Forms.RadioButton
  Friend WithEvents optDedupAddressOnly As System.Windows.Forms.RadioButton
  Friend WithEvents optDedupFull As System.Windows.Forms.RadioButton
  Friend WithEvents cboSeparator As System.Windows.Forms.ComboBox
  Friend WithEvents lblSeperator As System.Windows.Forms.Label
  Friend WithEvents cboTables As System.Windows.Forms.ComboBox
  Friend WithEvents lblTableDesc As System.Windows.Forms.Label
  Friend WithEvents lblGroups As System.Windows.Forms.Label
  Friend WithEvents cboGroups As System.Windows.Forms.ComboBox
  Friend WithEvents chkIgnore As System.Windows.Forms.CheckBox
  Friend WithEvents cboAttrs As System.Windows.Forms.ComboBox
  Friend WithEvents lblAttribute As System.Windows.Forms.Label
  Friend WithEvents lblColumn As System.Windows.Forms.Label
  Friend WithEvents cboType As System.Windows.Forms.ComboBox
  Friend WithEvents lblDataImportType As System.Windows.Forms.Label
  Friend WithEvents lblKey As System.Windows.Forms.Label
  Friend WithEvents chkKey As System.Windows.Forms.CheckBox
  Friend WithEvents lblMapValue As System.Windows.Forms.Label
  Friend WithEvents optMapValueLookup As System.Windows.Forms.RadioButton
  Friend WithEvents optMapValueNull As System.Windows.Forms.RadioButton
  Friend WithEvents tbpOptions As System.Windows.Forms.TabPage
  Friend WithEvents tabSub As CDBNETCL.TabControl
  Friend WithEvents tbpGeneralOpt As System.Windows.Forms.TabPage
  Friend WithEvents spltGeneralOpt As System.Windows.Forms.SplitContainer
  Friend WithEvents grpProcessingOptions As System.Windows.Forms.GroupBox
  Friend WithEvents lblControlNumberBlockSize As System.Windows.Forms.Label
  Friend WithEvents txtControlNumberBlockSize As System.Windows.Forms.TextBox
  Friend WithEvents txtReplaceQuestionMarkWith As System.Windows.Forms.TextBox
  Friend WithEvents chkReplaceQuestionMark As System.Windows.Forms.CheckBox
  Friend WithEvents chkCMDSupp As System.Windows.Forms.CheckBox
  Friend WithEvents chkCreateCMD As System.Windows.Forms.CheckBox
  Friend WithEvents chkValCodes As System.Windows.Forms.CheckBox
  Friend WithEvents chkNoIndexes As System.Windows.Forms.CheckBox
  Friend WithEvents chkControlNumbers As System.Windows.Forms.CheckBox
  Friend WithEvents grpLogFileOptions As System.Windows.Forms.GroupBox
  Friend WithEvents optMIRecordsOriginalFile As System.Windows.Forms.RadioButton
  Friend WithEvents optMIRecordsSuccFromFirstImport As System.Windows.Forms.RadioButton
  Friend WithEvents optMIRecordsSuccFromPrevImport As System.Windows.Forms.RadioButton
  Friend WithEvents lblMultipleImport As System.Windows.Forms.Label
  Friend WithEvents lblDefinitionFile As System.Windows.Forms.Label
  Friend WithEvents chkNoFileName As System.Windows.Forms.CheckBox
  Friend WithEvents chkLogConversion As System.Windows.Forms.CheckBox
  Friend WithEvents chkLogDedupAudit As System.Windows.Forms.CheckBox
  Friend WithEvents chkLogDups As System.Windows.Forms.CheckBox
  Friend WithEvents chkLogWarn As System.Windows.Forms.CheckBox
  Friend WithEvents chkLogCreate As System.Windows.Forms.CheckBox
  Friend WithEvents tbpMultImpRuns As System.Windows.Forms.TabPage
  Friend WithEvents lblSelectedImportType As System.Windows.Forms.Label
  Friend WithEvents chkDupAsError As System.Windows.Forms.CheckBox
  Friend WithEvents lstMultiImportSelected As System.Windows.Forms.ListBox
  Friend WithEvents cmdRemoveDefinition As System.Windows.Forms.Button
  Friend WithEvents cmdAddDefinition As System.Windows.Forms.Button
  Friend WithEvents lblAvailableImportTypes As System.Windows.Forms.Label
  Friend WithEvents lstMultiImportAvailable As System.Windows.Forms.ListBox
  Friend WithEvents tbpCustomOpt As System.Windows.Forms.TabPage
  Friend WithEvents pnlConAndOrg As System.Windows.Forms.Panel
  Friend WithEvents grpDataOption As System.Windows.Forms.GroupBox
  Friend WithEvents txtOrgNumber As CDBNETCL.TextLookupBox
  Friend WithEvents lblOrganisation As System.Windows.Forms.Label
  Friend WithEvents chkAddPosition As System.Windows.Forms.CheckBox
  Friend WithEvents chkEmployee As System.Windows.Forms.CheckBox
  Friend WithEvents chkCacheMailsort As System.Windows.Forms.CheckBox
  Friend WithEvents chkDefAddrFromUnknown As System.Windows.Forms.CheckBox
  Friend WithEvents chkDefSupp As System.Windows.Forms.CheckBox
  Friend WithEvents chkCreateGridRefs As System.Windows.Forms.CheckBox
  Friend WithEvents chkRePostcode As System.Windows.Forms.CheckBox
  Friend WithEvents chkPAFAddress As System.Windows.Forms.CheckBox
  Friend WithEvents chkSurnameFirst As System.Windows.Forms.CheckBox
  Friend WithEvents chkCaps As System.Windows.Forms.CheckBox
  Friend WithEvents chkDear As System.Windows.Forms.CheckBox
  Friend WithEvents grpDupUpdate As System.Windows.Forms.GroupBox
  Friend WithEvents lblUpdateSub As System.Windows.Forms.Label
  Friend WithEvents chkNameGatheringIncentives As System.Windows.Forms.CheckBox
  Friend WithEvents chkActivity As System.Windows.Forms.CheckBox
  Friend WithEvents chkUpdateWithNull As System.Windows.Forms.CheckBox
  Friend WithEvents chkUpdateAll As System.Windows.Forms.CheckBox
  Friend WithEvents chkUpdate As System.Windows.Forms.CheckBox
  Friend WithEvents grpDeDuplication As System.Windows.Forms.GroupBox
  Friend WithEvents chkBankDetailsDedup As System.Windows.Forms.CheckBox
  Friend WithEvents chkOrgAddressPotDup As System.Windows.Forms.CheckBox
  Friend WithEvents chkSoundexDedup As System.Windows.Forms.CheckBox
  Friend WithEvents chkAddressDedup As System.Windows.Forms.CheckBox
  Friend WithEvents chkForeInitDeDup As System.Windows.Forms.CheckBox
  Friend WithEvents chkTitleDeDup As System.Windows.Forms.CheckBox
  Friend WithEvents chkExclUnkAdd As System.Windows.Forms.CheckBox
  Friend WithEvents chkEmailDedup As System.Windows.Forms.CheckBox
  Friend WithEvents chkNumberDeDup As System.Windows.Forms.CheckBox
  Friend WithEvents chkExtRefDeDup As System.Windows.Forms.CheckBox
  Friend WithEvents pnlTableImport As System.Windows.Forms.Panel
  Friend WithEvents chkEmptyBeforeImport As System.Windows.Forms.CheckBox
  Friend WithEvents grp As System.Windows.Forms.GroupBox
  Friend WithEvents lblTableImport As System.Windows.Forms.Label
  Friend WithEvents chkUpdateAllTableImport As System.Windows.Forms.CheckBox
  Friend WithEvents chkUpdateTableImport As System.Windows.Forms.CheckBox
  Friend WithEvents pnlPayment As System.Windows.Forms.Panel
  Friend WithEvents lblNumberOfDays As System.Windows.Forms.Label
  Friend WithEvents txtNumberOfDays As System.Windows.Forms.TextBox
  Friend WithEvents chkMatchSchPayment As System.Windows.Forms.CheckBox
  Friend WithEvents chkCreateAct As System.Windows.Forms.CheckBox
  Friend WithEvents chkSkipZeroAmt As System.Windows.Forms.CheckBox
  Friend WithEvents chkProcessIncentives As System.Windows.Forms.CheckBox
  Friend WithEvents chkReference As System.Windows.Forms.CheckBox
  Friend WithEvents chkAddTransactions As System.Windows.Forms.CheckBox
  Friend WithEvents chkNoFromFile As System.Windows.Forms.CheckBox
  Friend WithEvents chkGiftAidRecords As System.Windows.Forms.CheckBox
  Friend WithEvents grpTypeOfPayment As System.Windows.Forms.GroupBox
  Friend WithEvents optPaymentsUnposted As System.Windows.Forms.RadioButton
  Friend WithEvents optPaymentsPostedToNominal As System.Windows.Forms.RadioButton
  Friend WithEvents optPaymentsFinHistory As System.Windows.Forms.RadioButton
  Friend WithEvents optPaymentsPostedToCB As System.Windows.Forms.RadioButton
  Friend WithEvents pnlDocument As System.Windows.Forms.Panel
  Friend WithEvents grpDupUpdateOpt As System.Windows.Forms.GroupBox
  Friend WithEvents chkUpdateAllDoc As System.Windows.Forms.CheckBox
  Friend WithEvents chkUpdateDoc As System.Windows.Forms.CheckBox
  Friend WithEvents lblUpdateExisting As System.Windows.Forms.Label
  Friend WithEvents pnlStock As System.Windows.Forms.Panel
  Friend WithEvents optStockSet As System.Windows.Forms.RadioButton
  Friend WithEvents optStockUpdate As System.Windows.Forms.RadioButton
  Friend WithEvents pnlBankTransactions As System.Windows.Forms.Panel
  Friend WithEvents chkDASImport As System.Windows.Forms.CheckBox
  Friend WithEvents pnlAddrUpdate As System.Windows.Forms.Panel
  Friend WithEvents chkCacheMailsortAddr As System.Windows.Forms.CheckBox
  Friend WithEvents chkExtractAddr As System.Windows.Forms.CheckBox
  Friend WithEvents tbpDefaults As System.Windows.Forms.TabPage
  Friend WithEvents pnlDefaults As System.Windows.Forms.Panel
  Friend WithEvents txtLookupDefValue As CDBNETCL.TextLookupBox
  Friend WithEvents lblElse As System.Windows.Forms.Label
  Friend WithEvents dtpckValue As System.Windows.Forms.DateTimePicker
  Friend WithEvents cmdDefaultAdd As System.Windows.Forms.Button
  Friend WithEvents cboPatternValue As System.Windows.Forms.ComboBox
  Friend WithEvents cboDefAttrs As System.Windows.Forms.ComboBox
  Friend WithEvents chkCtrlNo As System.Windows.Forms.CheckBox
  Friend WithEvents chkIncPerLine As System.Windows.Forms.CheckBox
  Friend WithEvents lblValue As System.Windows.Forms.Label
  Friend WithEvents lblDefAttribute As System.Windows.Forms.Label
  Friend WithEvents txtDefValue As System.Windows.Forms.TextBox
  Friend WithEvents dgrDefaults As CDBNETCL.DisplayGrid
  Friend WithEvents SplitControl As System.Windows.Forms.SplitContainer
  Friend WithEvents bpl As CDBNETCL.ButtonPanel
  Friend WithEvents cmdStop As System.Windows.Forms.Button
  Friend WithEvents cmdTest As System.Windows.Forms.Button
  Friend WithEvents cmdOk As System.Windows.Forms.Button
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents cmdSave As System.Windows.Forms.Button
  Friend WithEvents lblJobNumber As System.Windows.Forms.Label
  Friend WithEvents lblStatus As System.Windows.Forms.Label
  Friend WithEvents chkAmendmentHistory As System.Windows.Forms.CheckBox
  Friend WithEvents chkOrgNamePostCodeAddress As System.Windows.Forms.CheckBox
  Friend WithEvents chkAllowBlankForOrgName As System.Windows.Forms.CheckBox
  Friend WithEvents lnkHelp As LinkLabel

End Class
