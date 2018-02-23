<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAddCriteria
  Inherits ThemedForm

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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAddCriteria))
    Me.lblSearchArea = New System.Windows.Forms.Label()
    Me.txtLookupArea = New CDBNETCL.TextLookupBox()
    Me.cmdOK = New System.Windows.Forms.Button()
    Me.cmdCancel = New System.Windows.Forms.Button()
    Me.Panel1 = New System.Windows.Forms.Panel()
    Me.optPerson = New System.Windows.Forms.RadioButton()
    Me.optOrganisation = New System.Windows.Forms.RadioButton()
    Me.Panel2 = New System.Windows.Forms.Panel()
    Me.optExclude = New System.Windows.Forms.RadioButton()
    Me.optInclude = New System.Windows.Forms.RadioButton()
    Me.chkYN = New System.Windows.Forms.CheckBox()
    Me.dtpValue = New System.Windows.Forms.DateTimePicker()
    Me.lstValue = New System.Windows.Forms.ListBox()
    Me.cmdAddValue = New System.Windows.Forms.Button()
    Me.cmdDeleteValue = New System.Windows.Forms.Button()
    Me.dtpEndValue = New System.Windows.Forms.DateTimePicker()
    Me.chkValueRange = New System.Windows.Forms.CheckBox()
    Me.cmdDeleteSubValue = New System.Windows.Forms.Button()
    Me.cmdAddSubValue = New System.Windows.Forms.Button()
    Me.lstSubValue = New System.Windows.Forms.ListBox()
    Me.dtpSubValue = New System.Windows.Forms.DateTimePicker()
    Me.dtpEndSubValue = New System.Windows.Forms.DateTimePicker()
    Me.chkSubValueRange = New System.Windows.Forms.CheckBox()
    Me.vse7 = New System.Windows.Forms.Panel()
    Me.ButtonPanel1 = New CDBNETCL.ButtonPanel()
    Me.vse2 = New System.Windows.Forms.Panel()
    Me.cmdValueVar = New System.Windows.Forms.Button()
    Me.lblValue = New System.Windows.Forms.Label()
    Me.txtEndValue = New System.Windows.Forms.TextBox()
    Me.txtLookupValue = New CDBNETCL.TextLookupBox()
    Me.txtValue = New System.Windows.Forms.TextBox()
    Me.Panel4 = New System.Windows.Forms.Panel()
    Me.cmdPeriodVar = New System.Windows.Forms.Button()
    Me.cmdSubValueVar = New System.Windows.Forms.Button()
    Me.lblSubValue = New System.Windows.Forms.Label()
    Me.txtPeriodVar = New System.Windows.Forms.TextBox()
    Me.lblPeriodVar = New System.Windows.Forms.Label()
    Me.txtLookupEndSubValue = New CDBNETCL.TextLookupBox()
    Me.txtEndSubValue = New System.Windows.Forms.TextBox()
    Me.txtLookupSubValue = New CDBNETCL.TextLookupBox()
    Me.txtSubValue = New System.Windows.Forms.TextBox()
    Me.dtpTo = New System.Windows.Forms.DateTimePicker()
    Me.lblPeriodTo = New System.Windows.Forms.Label()
    Me.cmdAddPeriod = New System.Windows.Forms.Button()
    Me.cmdDeletePeriod = New System.Windows.Forms.Button()
    Me.dtpFrom = New System.Windows.Forms.DateTimePicker()
    Me.lstPeriod = New System.Windows.Forms.ListBox()
    Me.lblPeriodFrom = New System.Windows.Forms.Label()
    Me.Panel3 = New System.Windows.Forms.Panel()
    Me.txtLookupEndValue = New CDBNETCL.TextLookupBox()
    Me.Panel1.SuspendLayout()
    Me.Panel2.SuspendLayout()
    Me.vse7.SuspendLayout()
    Me.ButtonPanel1.SuspendLayout()
    Me.vse2.SuspendLayout()
    Me.Panel4.SuspendLayout()
    Me.Panel3.SuspendLayout()
    Me.SuspendLayout()
    '
    'lblSearchArea
    '
    Me.lblSearchArea.AutoSize = True
    Me.lblSearchArea.Cursor = System.Windows.Forms.Cursors.Default
    Me.lblSearchArea.Location = New System.Drawing.Point(7, 9)
    Me.lblSearchArea.Name = "lblSearchArea"
    Me.lblSearchArea.Size = New System.Drawing.Size(69, 13)
    Me.lblSearchArea.TabIndex = 0
    Me.lblSearchArea.Text = "Search Area:"
    '
    'txtLookupArea
    '
    Me.txtLookupArea.ActiveOnly = False
    Me.txtLookupArea.BackColor = System.Drawing.SystemColors.Control
    Me.txtLookupArea.Cursor = System.Windows.Forms.Cursors.Default
    Me.txtLookupArea.CustomFormNumber = 0
    Me.txtLookupArea.Description = ""
    Me.txtLookupArea.EnabledProperty = True
    Me.txtLookupArea.HasDependancies = False
    Me.txtLookupArea.IsDesign = False
    Me.txtLookupArea.Location = New System.Drawing.Point(107, 6)
    Me.txtLookupArea.MaxLength = 32767
    Me.txtLookupArea.MultipleValuesSupported = False
    Me.txtLookupArea.Name = "txtLookupArea"
    Me.txtLookupArea.OriginalText = Nothing
    Me.txtLookupArea.ReadOnlyProperty = False
    Me.txtLookupArea.Size = New System.Drawing.Size(407, 31)
    Me.txtLookupArea.TabIndex = 1
    Me.txtLookupArea.TextReadOnly = False
    Me.txtLookupArea.TotalWidth = 408
    Me.txtLookupArea.ValidationRequired = True
    '
    'cmdOK
    '
    Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
    Me.cmdOK.Location = New System.Drawing.Point(216, 6)
    Me.cmdOK.Name = "cmdOK"
    Me.cmdOK.Size = New System.Drawing.Size(96, 27)
    Me.cmdOK.TabIndex = 0
    Me.cmdOK.Text = "OK"
    Me.cmdOK.UseVisualStyleBackColor = True
    '
    'cmdCancel
    '
    Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
    Me.cmdCancel.Location = New System.Drawing.Point(327, 6)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(96, 27)
    Me.cmdCancel.TabIndex = 1
    Me.cmdCancel.Text = "Cancel"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'Panel1
    '
    Me.Panel1.Controls.Add(Me.optPerson)
    Me.Panel1.Controls.Add(Me.optOrganisation)
    Me.Panel1.Cursor = System.Windows.Forms.Cursors.Default
    Me.Panel1.Location = New System.Drawing.Point(9, 43)
    Me.Panel1.Name = "Panel1"
    Me.Panel1.Size = New System.Drawing.Size(228, 28)
    Me.Panel1.TabIndex = 2
    '
    'optPerson
    '
    Me.optPerson.AutoSize = True
    Me.optPerson.Checked = True
    Me.optPerson.Cursor = System.Windows.Forms.Cursors.Default
    Me.optPerson.Location = New System.Drawing.Point(119, 3)
    Me.optPerson.Name = "optPerson"
    Me.optPerson.Size = New System.Drawing.Size(62, 17)
    Me.optPerson.TabIndex = 1
    Me.optPerson.TabStop = True
    Me.optPerson.Text = "Contact"
    Me.optPerson.UseVisualStyleBackColor = True
    '
    'optOrganisation
    '
    Me.optOrganisation.AutoSize = True
    Me.optOrganisation.Cursor = System.Windows.Forms.Cursors.Default
    Me.optOrganisation.Location = New System.Drawing.Point(3, 3)
    Me.optOrganisation.Name = "optOrganisation"
    Me.optOrganisation.Size = New System.Drawing.Size(84, 17)
    Me.optOrganisation.TabIndex = 0
    Me.optOrganisation.Text = "Organisation"
    Me.optOrganisation.UseVisualStyleBackColor = True
    '
    'Panel2
    '
    Me.Panel2.Controls.Add(Me.optExclude)
    Me.Panel2.Controls.Add(Me.optInclude)
    Me.Panel2.Cursor = System.Windows.Forms.Cursors.Default
    Me.Panel2.Location = New System.Drawing.Point(289, 43)
    Me.Panel2.Name = "Panel2"
    Me.Panel2.Size = New System.Drawing.Size(225, 28)
    Me.Panel2.TabIndex = 3
    '
    'optExclude
    '
    Me.optExclude.AutoSize = True
    Me.optExclude.Cursor = System.Windows.Forms.Cursors.Default
    Me.optExclude.Location = New System.Drawing.Point(116, 3)
    Me.optExclude.Name = "optExclude"
    Me.optExclude.Size = New System.Drawing.Size(63, 17)
    Me.optExclude.TabIndex = 1
    Me.optExclude.Text = "Exclude"
    Me.optExclude.UseVisualStyleBackColor = True
    '
    'optInclude
    '
    Me.optInclude.AutoSize = True
    Me.optInclude.Checked = True
    Me.optInclude.Cursor = System.Windows.Forms.Cursors.Default
    Me.optInclude.Location = New System.Drawing.Point(8, 3)
    Me.optInclude.Name = "optInclude"
    Me.optInclude.Size = New System.Drawing.Size(60, 17)
    Me.optInclude.TabIndex = 0
    Me.optInclude.TabStop = True
    Me.optInclude.Text = "Include"
    Me.optInclude.UseVisualStyleBackColor = True
    '
    'chkYN
    '
    Me.chkYN.AutoSize = True
    Me.chkYN.Cursor = System.Windows.Forms.Cursors.Default
    Me.chkYN.Location = New System.Drawing.Point(11, 80)
    Me.chkYN.Name = "chkYN"
    Me.chkYN.Size = New System.Drawing.Size(53, 17)
    Me.chkYN.TabIndex = 6
    Me.chkYN.Text = "Value"
    Me.chkYN.UseVisualStyleBackColor = True
    '
    'dtpValue
    '
    Me.dtpValue.Cursor = System.Windows.Forms.Cursors.Default
    Me.dtpValue.CustomFormat = "dd/MM/yyyy"
    Me.dtpValue.Format = System.Windows.Forms.DateTimePickerFormat.Custom
    Me.dtpValue.Location = New System.Drawing.Point(247, 79)
    Me.dtpValue.Name = "dtpValue"
    Me.dtpValue.ShowCheckBox = True
    Me.dtpValue.Size = New System.Drawing.Size(127, 20)
    Me.dtpValue.TabIndex = 8
    Me.dtpValue.Visible = False
    '
    'lstValue
    '
    Me.lstValue.Cursor = System.Windows.Forms.Cursors.Default
    Me.lstValue.FormattingEnabled = True
    Me.lstValue.Location = New System.Drawing.Point(403, 76)
    Me.lstValue.Name = "lstValue"
    Me.lstValue.Size = New System.Drawing.Size(120, 43)
    Me.lstValue.TabIndex = 8
    '
    'cmdAddValue
    '
    Me.cmdAddValue.Cursor = System.Windows.Forms.Cursors.Default
    Me.cmdAddValue.Location = New System.Drawing.Point(3, 10)
    Me.cmdAddValue.Name = "cmdAddValue"
    Me.cmdAddValue.Size = New System.Drawing.Size(75, 23)
    Me.cmdAddValue.TabIndex = 0
    Me.cmdAddValue.Text = "Add"
    Me.cmdAddValue.UseVisualStyleBackColor = True
    '
    'cmdDeleteValue
    '
    Me.cmdDeleteValue.Cursor = System.Windows.Forms.Cursors.Default
    Me.cmdDeleteValue.Location = New System.Drawing.Point(3, 37)
    Me.cmdDeleteValue.Name = "cmdDeleteValue"
    Me.cmdDeleteValue.Size = New System.Drawing.Size(75, 23)
    Me.cmdDeleteValue.TabIndex = 1
    Me.cmdDeleteValue.Text = "Delete"
    Me.cmdDeleteValue.UseVisualStyleBackColor = True
    '
    'dtpEndValue
    '
    Me.dtpEndValue.Cursor = System.Windows.Forms.Cursors.Default
    Me.dtpEndValue.CustomFormat = "dd/MM/yyyy"
    Me.dtpEndValue.Format = System.Windows.Forms.DateTimePickerFormat.Custom
    Me.dtpEndValue.Location = New System.Drawing.Point(249, 149)
    Me.dtpEndValue.Name = "dtpEndValue"
    Me.dtpEndValue.ShowCheckBox = True
    Me.dtpEndValue.Size = New System.Drawing.Size(127, 20)
    Me.dtpEndValue.TabIndex = 14
    Me.dtpEndValue.Visible = False
    '
    'chkValueRange
    '
    Me.chkValueRange.AutoSize = True
    Me.chkValueRange.Cursor = System.Windows.Forms.Cursors.Default
    Me.chkValueRange.Location = New System.Drawing.Point(11, 149)
    Me.chkValueRange.Name = "chkValueRange"
    Me.chkValueRange.Size = New System.Drawing.Size(58, 17)
    Me.chkValueRange.TabIndex = 10
    Me.chkValueRange.Text = "Range"
    Me.chkValueRange.UseVisualStyleBackColor = True
    '
    'cmdDeleteSubValue
    '
    Me.cmdDeleteSubValue.Cursor = System.Windows.Forms.Cursors.Default
    Me.cmdDeleteSubValue.Location = New System.Drawing.Point(524, 35)
    Me.cmdDeleteSubValue.Name = "cmdDeleteSubValue"
    Me.cmdDeleteSubValue.Size = New System.Drawing.Size(75, 23)
    Me.cmdDeleteSubValue.TabIndex = 7
    Me.cmdDeleteSubValue.Text = "Delete"
    Me.cmdDeleteSubValue.UseVisualStyleBackColor = True
    '
    'cmdAddSubValue
    '
    Me.cmdAddSubValue.Cursor = System.Windows.Forms.Cursors.Default
    Me.cmdAddSubValue.Location = New System.Drawing.Point(524, 8)
    Me.cmdAddSubValue.Name = "cmdAddSubValue"
    Me.cmdAddSubValue.Size = New System.Drawing.Size(75, 23)
    Me.cmdAddSubValue.TabIndex = 6
    Me.cmdAddSubValue.Text = "Add"
    Me.cmdAddSubValue.UseVisualStyleBackColor = True
    '
    'lstSubValue
    '
    Me.lstSubValue.Cursor = System.Windows.Forms.Cursors.Default
    Me.lstSubValue.FormattingEnabled = True
    Me.lstSubValue.Location = New System.Drawing.Point(395, 8)
    Me.lstSubValue.Name = "lstSubValue"
    Me.lstSubValue.Size = New System.Drawing.Size(120, 43)
    Me.lstSubValue.TabIndex = 5
    '
    'dtpSubValue
    '
    Me.dtpSubValue.Cursor = System.Windows.Forms.Cursors.Default
    Me.dtpSubValue.CustomFormat = "dd/MM/yyyy"
    Me.dtpSubValue.Format = System.Windows.Forms.DateTimePickerFormat.Custom
    Me.dtpSubValue.Location = New System.Drawing.Point(240, 6)
    Me.dtpSubValue.Name = "dtpSubValue"
    Me.dtpSubValue.ShowCheckBox = True
    Me.dtpSubValue.Size = New System.Drawing.Size(127, 20)
    Me.dtpSubValue.TabIndex = 3
    Me.dtpSubValue.Visible = False
    '
    'dtpEndSubValue
    '
    Me.dtpEndSubValue.Cursor = System.Windows.Forms.Cursors.Default
    Me.dtpEndSubValue.CustomFormat = "dd/MM/yyyy"
    Me.dtpEndSubValue.Format = System.Windows.Forms.DateTimePickerFormat.Custom
    Me.dtpEndSubValue.Location = New System.Drawing.Point(239, 72)
    Me.dtpEndSubValue.Name = "dtpEndSubValue"
    Me.dtpEndSubValue.ShowCheckBox = True
    Me.dtpEndSubValue.Size = New System.Drawing.Size(127, 20)
    Me.dtpEndSubValue.TabIndex = 10
    Me.dtpEndSubValue.Visible = False
    '
    'chkSubValueRange
    '
    Me.chkSubValueRange.AutoSize = True
    Me.chkSubValueRange.Cursor = System.Windows.Forms.Cursors.Default
    Me.chkSubValueRange.Location = New System.Drawing.Point(6, 72)
    Me.chkSubValueRange.Name = "chkSubValueRange"
    Me.chkSubValueRange.Size = New System.Drawing.Size(58, 17)
    Me.chkSubValueRange.TabIndex = 8
    Me.chkSubValueRange.Text = "Range"
    Me.chkSubValueRange.UseVisualStyleBackColor = True
    '
    'vse7
    '
    Me.vse7.Controls.Add(Me.ButtonPanel1)
    Me.vse7.Cursor = System.Windows.Forms.Cursors.Default
    Me.vse7.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.vse7.Location = New System.Drawing.Point(0, 450)
    Me.vse7.Name = "vse7"
    Me.vse7.Size = New System.Drawing.Size(639, 46)
    Me.vse7.TabIndex = 1
    '
    'ButtonPanel1
    '
    Me.ButtonPanel1.Controls.Add(Me.cmdOK)
    Me.ButtonPanel1.Controls.Add(Me.cmdCancel)
    Me.ButtonPanel1.Cursor = System.Windows.Forms.Cursors.Default
    Me.ButtonPanel1.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.ButtonPanel1.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.ButtonPanel1.Location = New System.Drawing.Point(0, 7)
    Me.ButtonPanel1.Name = "ButtonPanel1"
    Me.ButtonPanel1.Size = New System.Drawing.Size(639, 39)
    Me.ButtonPanel1.TabIndex = 0
    '
    'vse2
    '
    Me.vse2.Controls.Add(Me.cmdValueVar)
    Me.vse2.Controls.Add(Me.lblValue)
    Me.vse2.Controls.Add(Me.txtEndValue)
    Me.vse2.Controls.Add(Me.txtLookupValue)
    Me.vse2.Controls.Add(Me.txtValue)
    Me.vse2.Controls.Add(Me.Panel4)
    Me.vse2.Controls.Add(Me.Panel3)
    Me.vse2.Controls.Add(Me.lblSearchArea)
    Me.vse2.Controls.Add(Me.txtLookupArea)
    Me.vse2.Controls.Add(Me.Panel1)
    Me.vse2.Controls.Add(Me.Panel2)
    Me.vse2.Controls.Add(Me.chkYN)
    Me.vse2.Controls.Add(Me.dtpValue)
    Me.vse2.Controls.Add(Me.lstValue)
    Me.vse2.Controls.Add(Me.dtpEndValue)
    Me.vse2.Controls.Add(Me.chkValueRange)
    Me.vse2.Controls.Add(Me.txtLookupEndValue)
    Me.vse2.Cursor = System.Windows.Forms.Cursors.Default
    Me.vse2.Dock = System.Windows.Forms.DockStyle.Top
    Me.vse2.Location = New System.Drawing.Point(0, 0)
    Me.vse2.Name = "vse2"
    Me.vse2.Size = New System.Drawing.Size(639, 430)
    Me.vse2.TabIndex = 0
    '
    'cmdValueVar
    '
    Me.cmdValueVar.Cursor = System.Windows.Forms.Cursors.Default
    Me.cmdValueVar.Location = New System.Drawing.Point(217, 76)
    Me.cmdValueVar.Name = "cmdValueVar"
    Me.cmdValueVar.Size = New System.Drawing.Size(24, 26)
    Me.cmdValueVar.TabIndex = 5
    Me.cmdValueVar.Text = "$"
    Me.cmdValueVar.UseVisualStyleBackColor = True
    '
    'lblValue
    '
    Me.lblValue.AutoSize = True
    Me.lblValue.Cursor = System.Windows.Forms.Cursors.Default
    Me.lblValue.Location = New System.Drawing.Point(11, 81)
    Me.lblValue.Name = "lblValue"
    Me.lblValue.Size = New System.Drawing.Size(34, 13)
    Me.lblValue.TabIndex = 4
    Me.lblValue.Text = "Value"
    Me.lblValue.Visible = False
    '
    'txtEndValue
    '
    Me.txtEndValue.Cursor = System.Windows.Forms.Cursors.Default
    Me.txtEndValue.Location = New System.Drawing.Point(249, 149)
    Me.txtEndValue.Name = "txtEndValue"
    Me.txtEndValue.Size = New System.Drawing.Size(100, 20)
    Me.txtEndValue.TabIndex = 11
    '
    'txtLookupValue
    '
    Me.txtLookupValue.ActiveOnly = False
    Me.txtLookupValue.BackColor = System.Drawing.SystemColors.Control
    Me.txtLookupValue.Cursor = System.Windows.Forms.Cursors.Default
    Me.txtLookupValue.CustomFormNumber = 0
    Me.txtLookupValue.Description = ""
    Me.txtLookupValue.EnabledProperty = True
    Me.txtLookupValue.HasDependancies = False
    Me.txtLookupValue.IsDesign = False
    Me.txtLookupValue.Location = New System.Drawing.Point(87, 107)
    Me.txtLookupValue.MaxLength = 32767
    Me.txtLookupValue.MultipleValuesSupported = False
    Me.txtLookupValue.Name = "txtLookupValue"
    Me.txtLookupValue.OriginalText = Nothing
    Me.txtLookupValue.ReadOnlyProperty = False
    Me.txtLookupValue.Size = New System.Drawing.Size(288, 24)
    Me.txtLookupValue.TabIndex = 7
    Me.txtLookupValue.TextReadOnly = False
    Me.txtLookupValue.TotalWidth = 408
    Me.txtLookupValue.ValidationRequired = True
    '
    'txtValue
    '
    Me.txtValue.Cursor = System.Windows.Forms.Cursors.Default
    Me.txtValue.Location = New System.Drawing.Point(247, 79)
    Me.txtValue.Name = "txtValue"
    Me.txtValue.Size = New System.Drawing.Size(100, 20)
    Me.txtValue.TabIndex = 6
    '
    'Panel4
    '
    Me.Panel4.Controls.Add(Me.cmdPeriodVar)
    Me.Panel4.Controls.Add(Me.cmdSubValueVar)
    Me.Panel4.Controls.Add(Me.lblSubValue)
    Me.Panel4.Controls.Add(Me.txtPeriodVar)
    Me.Panel4.Controls.Add(Me.lblPeriodVar)
    Me.Panel4.Controls.Add(Me.txtLookupEndSubValue)
    Me.Panel4.Controls.Add(Me.txtEndSubValue)
    Me.Panel4.Controls.Add(Me.txtLookupSubValue)
    Me.Panel4.Controls.Add(Me.txtSubValue)
    Me.Panel4.Controls.Add(Me.dtpTo)
    Me.Panel4.Controls.Add(Me.lblPeriodTo)
    Me.Panel4.Controls.Add(Me.cmdAddPeriod)
    Me.Panel4.Controls.Add(Me.cmdDeletePeriod)
    Me.Panel4.Controls.Add(Me.dtpFrom)
    Me.Panel4.Controls.Add(Me.lstPeriod)
    Me.Panel4.Controls.Add(Me.lblPeriodFrom)
    Me.Panel4.Controls.Add(Me.cmdAddSubValue)
    Me.Panel4.Controls.Add(Me.cmdDeleteSubValue)
    Me.Panel4.Controls.Add(Me.dtpEndSubValue)
    Me.Panel4.Controls.Add(Me.dtpSubValue)
    Me.Panel4.Controls.Add(Me.lstSubValue)
    Me.Panel4.Controls.Add(Me.chkSubValueRange)
    Me.Panel4.Cursor = System.Windows.Forms.Cursors.Default
    Me.Panel4.Location = New System.Drawing.Point(9, 211)
    Me.Panel4.Name = "Panel4"
    Me.Panel4.Size = New System.Drawing.Size(608, 205)
    Me.Panel4.TabIndex = 13
    '
    'cmdPeriodVar
    '
    Me.cmdPeriodVar.Cursor = System.Windows.Forms.Cursors.Default
    Me.cmdPeriodVar.Location = New System.Drawing.Point(116, 142)
    Me.cmdPeriodVar.Name = "cmdPeriodVar"
    Me.cmdPeriodVar.Size = New System.Drawing.Size(24, 23)
    Me.cmdPeriodVar.TabIndex = 13
    Me.cmdPeriodVar.Text = "$"
    Me.cmdPeriodVar.UseVisualStyleBackColor = True
    '
    'cmdSubValueVar
    '
    Me.cmdSubValueVar.Cursor = System.Windows.Forms.Cursors.Default
    Me.cmdSubValueVar.Location = New System.Drawing.Point(210, 3)
    Me.cmdSubValueVar.Name = "cmdSubValueVar"
    Me.cmdSubValueVar.Size = New System.Drawing.Size(24, 25)
    Me.cmdSubValueVar.TabIndex = 1
    Me.cmdSubValueVar.Text = "$"
    Me.cmdSubValueVar.UseVisualStyleBackColor = True
    '
    'lblSubValue
    '
    Me.lblSubValue.AutoSize = True
    Me.lblSubValue.Cursor = System.Windows.Forms.Cursors.Default
    Me.lblSubValue.Location = New System.Drawing.Point(0, 6)
    Me.lblSubValue.Name = "lblSubValue"
    Me.lblSubValue.Size = New System.Drawing.Size(56, 13)
    Me.lblSubValue.TabIndex = 0
    Me.lblSubValue.Text = "Sub Value"
    Me.lblSubValue.Visible = False
    '
    'txtPeriodVar
    '
    Me.txtPeriodVar.Cursor = System.Windows.Forms.Cursors.Default
    Me.txtPeriodVar.Location = New System.Drawing.Point(146, 142)
    Me.txtPeriodVar.Name = "txtPeriodVar"
    Me.txtPeriodVar.Size = New System.Drawing.Size(100, 20)
    Me.txtPeriodVar.TabIndex = 14
    '
    'lblPeriodVar
    '
    Me.lblPeriodVar.AutoSize = True
    Me.lblPeriodVar.Cursor = System.Windows.Forms.Cursors.Default
    Me.lblPeriodVar.Location = New System.Drawing.Point(4, 142)
    Me.lblPeriodVar.Name = "lblPeriodVar"
    Me.lblPeriodVar.Size = New System.Drawing.Size(40, 13)
    Me.lblPeriodVar.TabIndex = 12
    Me.lblPeriodVar.Text = "Period:"
    '
    'txtLookupEndSubValue
    '
    Me.txtLookupEndSubValue.ActiveOnly = False
    Me.txtLookupEndSubValue.BackColor = System.Drawing.SystemColors.Control
    Me.txtLookupEndSubValue.Cursor = System.Windows.Forms.Cursors.Default
    Me.txtLookupEndSubValue.CustomFormNumber = 0
    Me.txtLookupEndSubValue.Description = ""
    Me.txtLookupEndSubValue.EnabledProperty = True
    Me.txtLookupEndSubValue.HasDependancies = False
    Me.txtLookupEndSubValue.IsDesign = False
    Me.txtLookupEndSubValue.Location = New System.Drawing.Point(81, 98)
    Me.txtLookupEndSubValue.MaxLength = 32767
    Me.txtLookupEndSubValue.MultipleValuesSupported = False
    Me.txtLookupEndSubValue.Name = "txtLookupEndSubValue"
    Me.txtLookupEndSubValue.OriginalText = Nothing
    Me.txtLookupEndSubValue.ReadOnlyProperty = False
    Me.txtLookupEndSubValue.Size = New System.Drawing.Size(286, 24)
    Me.txtLookupEndSubValue.TabIndex = 11
    Me.txtLookupEndSubValue.TextReadOnly = False
    Me.txtLookupEndSubValue.TotalWidth = 408
    Me.txtLookupEndSubValue.ValidationRequired = True
    '
    'txtEndSubValue
    '
    Me.txtEndSubValue.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
    Me.txtEndSubValue.Cursor = System.Windows.Forms.Cursors.Default
    Me.txtEndSubValue.Location = New System.Drawing.Point(239, 72)
    Me.txtEndSubValue.Name = "txtEndSubValue"
    Me.txtEndSubValue.Size = New System.Drawing.Size(100, 20)
    Me.txtEndSubValue.TabIndex = 9
    '
    'txtLookupSubValue
    '
    Me.txtLookupSubValue.ActiveOnly = False
    Me.txtLookupSubValue.BackColor = System.Drawing.SystemColors.Control
    Me.txtLookupSubValue.Cursor = System.Windows.Forms.Cursors.Default
    Me.txtLookupSubValue.CustomFormNumber = 0
    Me.txtLookupSubValue.Description = ""
    Me.txtLookupSubValue.EnabledProperty = True
    Me.txtLookupSubValue.HasDependancies = False
    Me.txtLookupSubValue.IsDesign = False
    Me.txtLookupSubValue.Location = New System.Drawing.Point(81, 32)
    Me.txtLookupSubValue.MaxLength = 32767
    Me.txtLookupSubValue.MultipleValuesSupported = False
    Me.txtLookupSubValue.Name = "txtLookupSubValue"
    Me.txtLookupSubValue.OriginalText = Nothing
    Me.txtLookupSubValue.ReadOnlyProperty = False
    Me.txtLookupSubValue.Size = New System.Drawing.Size(286, 24)
    Me.txtLookupSubValue.TabIndex = 4
    Me.txtLookupSubValue.TextReadOnly = False
    Me.txtLookupSubValue.TotalWidth = 408
    Me.txtLookupSubValue.ValidationRequired = True
    '
    'txtSubValue
    '
    Me.txtSubValue.Cursor = System.Windows.Forms.Cursors.Default
    Me.txtSubValue.Location = New System.Drawing.Point(240, 6)
    Me.txtSubValue.Name = "txtSubValue"
    Me.txtSubValue.Size = New System.Drawing.Size(100, 20)
    Me.txtSubValue.TabIndex = 2
    '
    'dtpTo
    '
    Me.dtpTo.Cursor = System.Windows.Forms.Cursors.Default
    Me.dtpTo.CustomFormat = "dd/MM/yyyy"
    Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
    Me.dtpTo.Location = New System.Drawing.Point(155, 178)
    Me.dtpTo.Name = "dtpTo"
    Me.dtpTo.ShowCheckBox = True
    Me.dtpTo.Size = New System.Drawing.Size(127, 20)
    Me.dtpTo.TabIndex = 17
    Me.dtpTo.Visible = False
    '
    'lblPeriodTo
    '
    Me.lblPeriodTo.AutoSize = True
    Me.lblPeriodTo.Cursor = System.Windows.Forms.Cursors.Default
    Me.lblPeriodTo.Location = New System.Drawing.Point(4, 178)
    Me.lblPeriodTo.Name = "lblPeriodTo"
    Me.lblPeriodTo.Size = New System.Drawing.Size(62, 13)
    Me.lblPeriodTo.TabIndex = 16
    Me.lblPeriodTo.Text = "Period (To):"
    '
    'cmdAddPeriod
    '
    Me.cmdAddPeriod.Cursor = System.Windows.Forms.Cursors.Default
    Me.cmdAddPeriod.Location = New System.Drawing.Point(524, 142)
    Me.cmdAddPeriod.Name = "cmdAddPeriod"
    Me.cmdAddPeriod.Size = New System.Drawing.Size(75, 23)
    Me.cmdAddPeriod.TabIndex = 19
    Me.cmdAddPeriod.Text = "Add"
    Me.cmdAddPeriod.UseVisualStyleBackColor = True
    '
    'cmdDeletePeriod
    '
    Me.cmdDeletePeriod.Cursor = System.Windows.Forms.Cursors.Default
    Me.cmdDeletePeriod.Location = New System.Drawing.Point(524, 171)
    Me.cmdDeletePeriod.Name = "cmdDeletePeriod"
    Me.cmdDeletePeriod.Size = New System.Drawing.Size(75, 23)
    Me.cmdDeletePeriod.TabIndex = 20
    Me.cmdDeletePeriod.Text = "Delete"
    Me.cmdDeletePeriod.UseVisualStyleBackColor = True
    '
    'dtpFrom
    '
    Me.dtpFrom.Cursor = System.Windows.Forms.Cursors.Default
    Me.dtpFrom.CustomFormat = "dd/MM/yyyy"
    Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
    Me.dtpFrom.Location = New System.Drawing.Point(155, 142)
    Me.dtpFrom.Name = "dtpFrom"
    Me.dtpFrom.ShowCheckBox = True
    Me.dtpFrom.Size = New System.Drawing.Size(127, 20)
    Me.dtpFrom.TabIndex = 15
    Me.dtpFrom.Visible = False
    '
    'lstPeriod
    '
    Me.lstPeriod.Cursor = System.Windows.Forms.Cursors.Default
    Me.lstPeriod.FormattingEnabled = True
    Me.lstPeriod.Location = New System.Drawing.Point(343, 142)
    Me.lstPeriod.Name = "lstPeriod"
    Me.lstPeriod.Size = New System.Drawing.Size(172, 43)
    Me.lstPeriod.TabIndex = 18
    '
    'lblPeriodFrom
    '
    Me.lblPeriodFrom.AutoSize = True
    Me.lblPeriodFrom.Cursor = System.Windows.Forms.Cursors.Default
    Me.lblPeriodFrom.Location = New System.Drawing.Point(3, 142)
    Me.lblPeriodFrom.Name = "lblPeriodFrom"
    Me.lblPeriodFrom.Size = New System.Drawing.Size(72, 13)
    Me.lblPeriodFrom.TabIndex = 24
    Me.lblPeriodFrom.Text = "Period (From):"
    '
    'Panel3
    '
    Me.Panel3.Controls.Add(Me.cmdAddValue)
    Me.Panel3.Controls.Add(Me.cmdDeleteValue)
    Me.Panel3.Cursor = System.Windows.Forms.Cursors.Default
    Me.Panel3.Location = New System.Drawing.Point(529, 65)
    Me.Panel3.Name = "Panel3"
    Me.Panel3.Size = New System.Drawing.Size(87, 80)
    Me.Panel3.TabIndex = 9
    '
    'txtLookupEndValue
    '
    Me.txtLookupEndValue.ActiveOnly = False
    Me.txtLookupEndValue.BackColor = System.Drawing.SystemColors.Control
    Me.txtLookupEndValue.Cursor = System.Windows.Forms.Cursors.Default
    Me.txtLookupEndValue.CustomFormNumber = 0
    Me.txtLookupEndValue.Description = ""
    Me.txtLookupEndValue.EnabledProperty = True
    Me.txtLookupEndValue.HasDependancies = False
    Me.txtLookupEndValue.IsDesign = False
    Me.txtLookupEndValue.Location = New System.Drawing.Point(89, 177)
    Me.txtLookupEndValue.MaxLength = 32767
    Me.txtLookupEndValue.MultipleValuesSupported = False
    Me.txtLookupEndValue.Name = "txtLookupEndValue"
    Me.txtLookupEndValue.OriginalText = Nothing
    Me.txtLookupEndValue.ReadOnlyProperty = False
    Me.txtLookupEndValue.Size = New System.Drawing.Size(288, 24)
    Me.txtLookupEndValue.TabIndex = 12
    Me.txtLookupEndValue.TextReadOnly = False
    Me.txtLookupEndValue.TotalWidth = 408
    Me.txtLookupEndValue.ValidationRequired = True
    '
    'frmAddCriteria
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(639, 496)
    Me.Controls.Add(Me.vse2)
    Me.Controls.Add(Me.vse7)
    Me.Cursor = System.Windows.Forms.Cursors.Default
    Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmAddCriteria"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    Me.Text = "Selection Manager - Edit Criteria"
    Me.Panel1.ResumeLayout(False)
    Me.Panel1.PerformLayout()
    Me.Panel2.ResumeLayout(False)
    Me.Panel2.PerformLayout()
    Me.vse7.ResumeLayout(False)
    Me.ButtonPanel1.ResumeLayout(False)
    Me.vse2.ResumeLayout(False)
    Me.vse2.PerformLayout()
    Me.Panel4.ResumeLayout(False)
    Me.Panel4.PerformLayout()
    Me.Panel3.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents lblSearchArea As System.Windows.Forms.Label
  Friend WithEvents txtLookupArea As CDBNETCL.TextLookupBox
  Friend WithEvents cmdOK As System.Windows.Forms.Button
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents Panel1 As System.Windows.Forms.Panel
  Friend WithEvents optOrganisation As System.Windows.Forms.RadioButton
  Friend WithEvents optPerson As System.Windows.Forms.RadioButton
  Friend WithEvents Panel2 As System.Windows.Forms.Panel
  Friend WithEvents optExclude As System.Windows.Forms.RadioButton
  Friend WithEvents optInclude As System.Windows.Forms.RadioButton
  Friend WithEvents chkYN As System.Windows.Forms.CheckBox
  Friend WithEvents dtpValue As System.Windows.Forms.DateTimePicker
  Friend WithEvents lstValue As System.Windows.Forms.ListBox
  Friend WithEvents cmdAddValue As System.Windows.Forms.Button
  Friend WithEvents cmdDeleteValue As System.Windows.Forms.Button
  Friend WithEvents dtpEndValue As System.Windows.Forms.DateTimePicker
  Friend WithEvents chkValueRange As System.Windows.Forms.CheckBox
  Friend WithEvents cmdDeleteSubValue As System.Windows.Forms.Button
  Friend WithEvents cmdAddSubValue As System.Windows.Forms.Button
  Friend WithEvents lstSubValue As System.Windows.Forms.ListBox
  Friend WithEvents dtpSubValue As System.Windows.Forms.DateTimePicker
  Friend WithEvents dtpEndSubValue As System.Windows.Forms.DateTimePicker
  Friend WithEvents chkSubValueRange As System.Windows.Forms.CheckBox
  Friend WithEvents vse7 As System.Windows.Forms.Panel
  Friend WithEvents vse2 As System.Windows.Forms.Panel
  Friend WithEvents Panel3 As System.Windows.Forms.Panel
  Friend WithEvents Panel4 As System.Windows.Forms.Panel
  Friend WithEvents lblPeriodFrom As System.Windows.Forms.Label
  Friend WithEvents cmdAddPeriod As System.Windows.Forms.Button
  Friend WithEvents cmdDeletePeriod As System.Windows.Forms.Button
  Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
  Friend WithEvents lstPeriod As System.Windows.Forms.ListBox
  Friend WithEvents lblPeriodTo As System.Windows.Forms.Label
  Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
  Friend WithEvents ButtonPanel1 As CDBNETCL.ButtonPanel
  Friend WithEvents txtValue As System.Windows.Forms.TextBox
  Friend WithEvents txtLookupValue As CDBNETCL.TextLookupBox
  Friend WithEvents txtLookupEndValue As CDBNETCL.TextLookupBox
  Friend WithEvents txtSubValue As System.Windows.Forms.TextBox
  Friend WithEvents txtLookupSubValue As CDBNETCL.TextLookupBox
  Friend WithEvents txtEndSubValue As System.Windows.Forms.TextBox
  Friend WithEvents txtLookupEndSubValue As CDBNETCL.TextLookupBox
  Friend WithEvents lblPeriodVar As System.Windows.Forms.Label
  Friend WithEvents txtPeriodVar As System.Windows.Forms.TextBox
  Friend WithEvents txtEndValue As System.Windows.Forms.TextBox
  Friend WithEvents lblValue As System.Windows.Forms.Label
  Friend WithEvents lblSubValue As System.Windows.Forms.Label
  Friend WithEvents cmdValueVar As System.Windows.Forms.Button
  Friend WithEvents cmdSubValueVar As System.Windows.Forms.Button
  Friend WithEvents cmdPeriodVar As System.Windows.Forms.Button
End Class
