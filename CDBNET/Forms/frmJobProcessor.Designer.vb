<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmJobProcessor
  Inherits CDBNETCL.ThemedForm

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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmJobProcessor))
    Me.bpl = New CDBNETCL.ButtonPanel()
    Me.cmdDelete = New System.Windows.Forms.Button()
    Me.cmdReSubmit = New System.Windows.Forms.Button()
    Me.cmdRefresh = New System.Windows.Forms.Button()
    Me.cmdClose = New System.Windows.Forms.Button()
    Me.tab = New CDBNETCL.TabControl()
    Me.tbpJobHistory = New System.Windows.Forms.TabPage()
    Me.dgrJobSchedule = New CDBNETCL.DisplayGrid()
    Me.tbpOptions = New System.Windows.Forms.TabPage()
    Me.lblMaxRows = New System.Windows.Forms.Label()
    Me.txtMaxRows = New System.Windows.Forms.TextBox()
    Me.txtTimerInterval = New System.Windows.Forms.TextBox()
    Me.lblRefreshInterval = New System.Windows.Forms.Label()
    Me.grpOptions = New System.Windows.Forms.GroupBox()
    Me.lblTo = New System.Windows.Forms.Label()
    Me.lblFrom = New System.Windows.Forms.Label()
    Me.dtpTo = New System.Windows.Forms.DateTimePicker()
    Me.dtpFrom = New System.Windows.Forms.DateTimePicker()
    Me.optCustomPeriod = New System.Windows.Forms.RadioButton()
    Me.optPrevQuater = New System.Windows.Forms.RadioButton()
    Me.optPrevMonth = New System.Windows.Forms.RadioButton()
    Me.optPrevWeek = New System.Windows.Forms.RadioButton()
    Me.tbpJobProcessors = New System.Windows.Forms.TabPage()
    Me.dgrJobProcessors = New CDBNETCL.DisplayGrid()
    Me.tim = New System.Windows.Forms.Timer(Me.components)
    Me.bpl.SuspendLayout()
    Me.tab.SuspendLayout()
    Me.tbpJobHistory.SuspendLayout()
    Me.tbpOptions.SuspendLayout()
    Me.grpOptions.SuspendLayout()
    Me.tbpJobProcessors.SuspendLayout()
    Me.SuspendLayout()
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdDelete)
    Me.bpl.Controls.Add(Me.cmdReSubmit)
    Me.bpl.Controls.Add(Me.cmdRefresh)
    Me.bpl.Controls.Add(Me.cmdClose)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bpl.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bpl.Location = New System.Drawing.Point(0, 383)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(724, 39)
    Me.bpl.TabIndex = 0
    '
    'cmdDelete
    '
    Me.cmdDelete.Location = New System.Drawing.Point(147, 6)
    Me.cmdDelete.Name = "cmdDelete"
    Me.cmdDelete.Size = New System.Drawing.Size(96, 27)
    Me.cmdDelete.TabIndex = 2
    Me.cmdDelete.Text = "&Delete"
    Me.cmdDelete.UseVisualStyleBackColor = True
    '
    'cmdReSubmit
    '
    Me.cmdReSubmit.Location = New System.Drawing.Point(258, 6)
    Me.cmdReSubmit.Name = "cmdReSubmit"
    Me.cmdReSubmit.Size = New System.Drawing.Size(96, 27)
    Me.cmdReSubmit.TabIndex = 1
    Me.cmdReSubmit.Text = "Re&Submit"
    Me.cmdReSubmit.UseVisualStyleBackColor = True
    '
    'cmdRefresh
    '
    Me.cmdRefresh.Location = New System.Drawing.Point(369, 6)
    Me.cmdRefresh.Name = "cmdRefresh"
    Me.cmdRefresh.Size = New System.Drawing.Size(96, 27)
    Me.cmdRefresh.TabIndex = 0
    Me.cmdRefresh.Text = "&Refresh"
    Me.cmdRefresh.UseVisualStyleBackColor = True
    '
    'cmdClose
    '
    Me.cmdClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdClose.Location = New System.Drawing.Point(480, 6)
    Me.cmdClose.Name = "cmdClose"
    Me.cmdClose.Size = New System.Drawing.Size(96, 27)
    Me.cmdClose.TabIndex = 3
    Me.cmdClose.Text = "Close"
    Me.cmdClose.UseVisualStyleBackColor = True
    '
    'tab
    '
    Me.tab.Controls.Add(Me.tbpJobHistory)
    Me.tab.Controls.Add(Me.tbpOptions)
    Me.tab.Controls.Add(Me.tbpJobProcessors)
    Me.tab.Dock = System.Windows.Forms.DockStyle.Fill
    Me.tab.ItemSize = New System.Drawing.Size(112, 20)
    Me.tab.Location = New System.Drawing.Point(0, 0)
    Me.tab.Name = "tab"
    Me.tab.SelectedIndex = 0
    Me.tab.Size = New System.Drawing.Size(724, 383)
    Me.tab.SizeMode = System.Windows.Forms.TabSizeMode.Fixed
    Me.tab.TabIndex = 1
    '
    'tbpJobHistory
    '
    Me.tbpJobHistory.Controls.Add(Me.dgrJobSchedule)
    Me.tbpJobHistory.Location = New System.Drawing.Point(4, 24)
    Me.tbpJobHistory.Name = "tbpJobHistory"
    Me.tbpJobHistory.Padding = New System.Windows.Forms.Padding(3)
    Me.tbpJobHistory.Size = New System.Drawing.Size(716, 355)
    Me.tbpJobHistory.TabIndex = 0
    Me.tbpJobHistory.Text = "Job History"
    Me.tbpJobHistory.UseVisualStyleBackColor = True
    '
    'dgrJobSchedule
    '
    Me.dgrJobSchedule.AccessibleName = "Display Grid"
    Me.dgrJobSchedule.ActiveColumn = 0
    Me.dgrJobSchedule.AllowSorting = True
    Me.dgrJobSchedule.AutoSetHeight = False
    Me.dgrJobSchedule.AutoSetRowHeight = False
    Me.dgrJobSchedule.DisplayTitle = Nothing
    Me.dgrJobSchedule.Dock = System.Windows.Forms.DockStyle.Fill
    Me.dgrJobSchedule.Location = New System.Drawing.Point(3, 3)
    Me.dgrJobSchedule.MaintenanceDesc = Nothing
    Me.dgrJobSchedule.MaxGridRows = 6
    Me.dgrJobSchedule.MultipleSelect = False
    Me.dgrJobSchedule.Name = "dgrJobSchedule"
    Me.dgrJobSchedule.RowCount = 10
    Me.dgrJobSchedule.ShowIfEmpty = False
    Me.dgrJobSchedule.Size = New System.Drawing.Size(710, 349)
    Me.dgrJobSchedule.SuppressHyperLinkFormat = False
    Me.dgrJobSchedule.TabIndex = 0
    '
    'tbpOptions
    '
    Me.tbpOptions.Controls.Add(Me.lblMaxRows)
    Me.tbpOptions.Controls.Add(Me.txtMaxRows)
    Me.tbpOptions.Controls.Add(Me.txtTimerInterval)
    Me.tbpOptions.Controls.Add(Me.lblRefreshInterval)
    Me.tbpOptions.Controls.Add(Me.grpOptions)
    Me.tbpOptions.Location = New System.Drawing.Point(4, 24)
    Me.tbpOptions.Name = "tbpOptions"
    Me.tbpOptions.Padding = New System.Windows.Forms.Padding(3)
    Me.tbpOptions.Size = New System.Drawing.Size(192, 72)
    Me.tbpOptions.TabIndex = 1
    Me.tbpOptions.Text = "Options"
    Me.tbpOptions.UseVisualStyleBackColor = True
    '
    'lblMaxRows
    '
    Me.lblMaxRows.AutoSize = True
    Me.lblMaxRows.Location = New System.Drawing.Point(8, 120)
    Me.lblMaxRows.Name = "lblMaxRows"
    Me.lblMaxRows.Size = New System.Drawing.Size(63, 13)
    Me.lblMaxRows.TabIndex = 5
    Me.lblMaxRows.Text = "Max Rows: "
    '
    'txtMaxRows
    '
    Me.txtMaxRows.Location = New System.Drawing.Point(93, 115)
    Me.txtMaxRows.Name = "txtMaxRows"
    Me.txtMaxRows.Size = New System.Drawing.Size(107, 20)
    Me.txtMaxRows.TabIndex = 4
    '
    'txtTimerInterval
    '
    Me.txtTimerInterval.Location = New System.Drawing.Point(569, 24)
    Me.txtTimerInterval.MaxLength = 3
    Me.txtTimerInterval.Name = "txtTimerInterval"
    Me.txtTimerInterval.Size = New System.Drawing.Size(52, 20)
    Me.txtTimerInterval.TabIndex = 3
    '
    'lblRefreshInterval
    '
    Me.lblRefreshInterval.AutoSize = True
    Me.lblRefreshInterval.Location = New System.Drawing.Point(408, 29)
    Me.lblRefreshInterval.Name = "lblRefreshInterval"
    Me.lblRefreshInterval.Size = New System.Drawing.Size(115, 13)
    Me.lblRefreshInterval.TabIndex = 1
    Me.lblRefreshInterval.Text = "Refresh Interval (mins):"
    '
    'grpOptions
    '
    Me.grpOptions.Controls.Add(Me.lblTo)
    Me.grpOptions.Controls.Add(Me.lblFrom)
    Me.grpOptions.Controls.Add(Me.dtpTo)
    Me.grpOptions.Controls.Add(Me.dtpFrom)
    Me.grpOptions.Controls.Add(Me.optCustomPeriod)
    Me.grpOptions.Controls.Add(Me.optPrevQuater)
    Me.grpOptions.Controls.Add(Me.optPrevMonth)
    Me.grpOptions.Controls.Add(Me.optPrevWeek)
    Me.grpOptions.Location = New System.Drawing.Point(8, 6)
    Me.grpOptions.Name = "grpOptions"
    Me.grpOptions.Size = New System.Drawing.Size(377, 103)
    Me.grpOptions.TabIndex = 0
    Me.grpOptions.TabStop = False
    Me.grpOptions.Text = "Job History For Jobs Due in Period"
    '
    'lblTo
    '
    Me.lblTo.AutoSize = True
    Me.lblTo.Location = New System.Drawing.Point(205, 71)
    Me.lblTo.Name = "lblTo"
    Me.lblTo.Size = New System.Drawing.Size(23, 13)
    Me.lblTo.TabIndex = 7
    Me.lblTo.Text = "To:"
    '
    'lblFrom
    '
    Me.lblFrom.AutoSize = True
    Me.lblFrom.Location = New System.Drawing.Point(204, 47)
    Me.lblFrom.Name = "lblFrom"
    Me.lblFrom.Size = New System.Drawing.Size(33, 13)
    Me.lblFrom.TabIndex = 6
    Me.lblFrom.Text = "From:"
    '
    'dtpTo
    '
    Me.dtpTo.CustomFormat = "dd/MM/yyyy"
    Me.dtpTo.Enabled = False
    Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
    Me.dtpTo.Location = New System.Drawing.Point(248, 69)
    Me.dtpTo.Name = "dtpTo"
    Me.dtpTo.ShowCheckBox = True
    Me.dtpTo.Size = New System.Drawing.Size(112, 20)
    Me.dtpTo.TabIndex = 5
    '
    'dtpFrom
    '
    Me.dtpFrom.CustomFormat = "dd/MM/yyyy"
    Me.dtpFrom.Enabled = False
    Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
    Me.dtpFrom.Location = New System.Drawing.Point(248, 44)
    Me.dtpFrom.Name = "dtpFrom"
    Me.dtpFrom.ShowCheckBox = True
    Me.dtpFrom.Size = New System.Drawing.Size(112, 20)
    Me.dtpFrom.TabIndex = 4
    '
    'optCustomPeriod
    '
    Me.optCustomPeriod.AutoSize = True
    Me.optCustomPeriod.Location = New System.Drawing.Point(207, 22)
    Me.optCustomPeriod.Name = "optCustomPeriod"
    Me.optCustomPeriod.Size = New System.Drawing.Size(93, 17)
    Me.optCustomPeriod.TabIndex = 3
    Me.optCustomPeriod.TabStop = True
    Me.optCustomPeriod.Text = "Custom Period"
    Me.optCustomPeriod.UseVisualStyleBackColor = True
    '
    'optPrevQuater
    '
    Me.optPrevQuater.AutoSize = True
    Me.optPrevQuater.Location = New System.Drawing.Point(11, 67)
    Me.optPrevQuater.Name = "optPrevQuater"
    Me.optPrevQuater.Size = New System.Drawing.Size(104, 17)
    Me.optPrevQuater.TabIndex = 2
    Me.optPrevQuater.TabStop = True
    Me.optPrevQuater.Text = "Previous Quarter"
    Me.optPrevQuater.UseVisualStyleBackColor = True
    '
    'optPrevMonth
    '
    Me.optPrevMonth.AutoSize = True
    Me.optPrevMonth.Location = New System.Drawing.Point(11, 45)
    Me.optPrevMonth.Name = "optPrevMonth"
    Me.optPrevMonth.Size = New System.Drawing.Size(99, 17)
    Me.optPrevMonth.TabIndex = 1
    Me.optPrevMonth.TabStop = True
    Me.optPrevMonth.Text = "Previous Month"
    Me.optPrevMonth.UseVisualStyleBackColor = True
    '
    'optPrevWeek
    '
    Me.optPrevWeek.AutoSize = True
    Me.optPrevWeek.Location = New System.Drawing.Point(11, 22)
    Me.optPrevWeek.Name = "optPrevWeek"
    Me.optPrevWeek.Size = New System.Drawing.Size(98, 17)
    Me.optPrevWeek.TabIndex = 0
    Me.optPrevWeek.TabStop = True
    Me.optPrevWeek.Text = "Previous Week"
    Me.optPrevWeek.UseVisualStyleBackColor = True
    '
    'tbpJobProcessors
    '
    Me.tbpJobProcessors.Controls.Add(Me.dgrJobProcessors)
    Me.tbpJobProcessors.Location = New System.Drawing.Point(4, 24)
    Me.tbpJobProcessors.Name = "tbpJobProcessors"
    Me.tbpJobProcessors.Size = New System.Drawing.Size(192, 72)
    Me.tbpJobProcessors.TabIndex = 2
    Me.tbpJobProcessors.Text = "Job Processors"
    Me.tbpJobProcessors.UseVisualStyleBackColor = True
    '
    'dgrJobProcessors
    '
    Me.dgrJobProcessors.AccessibleName = "Display Grid"
    Me.dgrJobProcessors.ActiveColumn = 0
    Me.dgrJobProcessors.AllowSorting = True
    Me.dgrJobProcessors.AutoSetHeight = False
    Me.dgrJobProcessors.AutoSetRowHeight = False
    Me.dgrJobProcessors.DisplayTitle = Nothing
    Me.dgrJobProcessors.Dock = System.Windows.Forms.DockStyle.Fill
    Me.dgrJobProcessors.Location = New System.Drawing.Point(0, 0)
    Me.dgrJobProcessors.MaintenanceDesc = Nothing
    Me.dgrJobProcessors.MaxGridRows = 6
    Me.dgrJobProcessors.MultipleSelect = False
    Me.dgrJobProcessors.Name = "dgrJobProcessors"
    Me.dgrJobProcessors.RowCount = 10
    Me.dgrJobProcessors.ShowIfEmpty = False
    Me.dgrJobProcessors.Size = New System.Drawing.Size(192, 72)
    Me.dgrJobProcessors.SuppressHyperLinkFormat = False
    Me.dgrJobProcessors.TabIndex = 0
    '
    'tim
    '
    '
    'frmJobProcessor
    '
    Me.CancelButton = Me.cmdClose
    Me.ClientSize = New System.Drawing.Size(724, 422)
    Me.Controls.Add(Me.tab)
    Me.Controls.Add(Me.bpl)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmJobProcessor"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    Me.Text = "Job Schedule"
    Me.bpl.ResumeLayout(False)
    Me.tab.ResumeLayout(False)
    Me.tbpJobHistory.ResumeLayout(False)
    Me.tbpOptions.ResumeLayout(False)
    Me.tbpOptions.PerformLayout()
    Me.grpOptions.ResumeLayout(False)
    Me.grpOptions.PerformLayout()
    Me.tbpJobProcessors.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents bpl As CDBNETCL.ButtonPanel
  Friend WithEvents cmdDelete As System.Windows.Forms.Button
  Friend WithEvents cmdReSubmit As System.Windows.Forms.Button
  Friend WithEvents cmdRefresh As System.Windows.Forms.Button
  Friend WithEvents tab As CDBNETCL.TabControl
  Friend WithEvents tbpJobHistory As System.Windows.Forms.TabPage
  Friend WithEvents tbpOptions As System.Windows.Forms.TabPage
  Friend WithEvents tbpJobProcessors As System.Windows.Forms.TabPage
  Friend WithEvents dgrJobSchedule As CDBNETCL.DisplayGrid
  Friend WithEvents dgrJobProcessors As CDBNETCL.DisplayGrid
  Friend WithEvents grpOptions As System.Windows.Forms.GroupBox
  Friend WithEvents optCustomPeriod As System.Windows.Forms.RadioButton
  Friend WithEvents optPrevQuater As System.Windows.Forms.RadioButton
  Friend WithEvents optPrevMonth As System.Windows.Forms.RadioButton
  Friend WithEvents optPrevWeek As System.Windows.Forms.RadioButton
  Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
  Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
  Friend WithEvents lblTo As System.Windows.Forms.Label
  Friend WithEvents lblFrom As System.Windows.Forms.Label
  Friend WithEvents txtTimerInterval As System.Windows.Forms.TextBox
  Friend WithEvents lblRefreshInterval As System.Windows.Forms.Label
  Friend WithEvents tim As System.Windows.Forms.Timer
  Friend WithEvents cmdClose As System.Windows.Forms.Button
  Friend WithEvents lblMaxRows As System.Windows.Forms.Label
  Friend WithEvents txtMaxRows As System.Windows.Forms.TextBox

End Class
