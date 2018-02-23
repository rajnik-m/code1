<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmEditCriteria
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
    Me.components = New System.ComponentModel.Container()
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmEditCriteria))
    Me.cmdOK = New System.Windows.Forms.Button()
    Me.cmdUpdate = New System.Windows.Forms.Button()
    Me.vseInfo = New System.Windows.Forms.GroupBox()
    Me.lblCurrentCriteria = New System.Windows.Forms.Label()
    Me.lblCurrentList = New System.Windows.Forms.Label()
    Me.cmdCancel = New System.Windows.Forms.Button()
    Me.cmdCount = New System.Windows.Forms.Button()
    Me.cmdCriteria = New System.Windows.Forms.Button()
    Me.cmdClear = New System.Windows.Forms.Button()
    Me.cmdLists = New System.Windows.Forms.Button()
    Me.cmdSaveCriteria = New System.Windows.Forms.Button()
    Me.cmdAdd = New System.Windows.Forms.Button()
    Me.cmdDelete = New System.Windows.Forms.Button()
    Me.pnlButton = New System.Windows.Forms.Panel()
    Me.ButtonPanel1 = New CDBNETCL.ButtonPanel()
    Me.timPolling = New System.Windows.Forms.Timer(Me.components)
    Me.SplitContainerForGrid = New System.Windows.Forms.SplitContainer()
    Me.grpCurrentCriteria = New System.Windows.Forms.GroupBox()
    Me.dgrCriteria = New CDBNETCL.DisplayGrid()
    Me.grpStandardExclusions = New System.Windows.Forms.GroupBox()
    Me.dgrDefault = New CDBNETCL.DisplayGrid()
    Me.pnlBottom = New System.Windows.Forms.Panel()
    Me.vseOptions = New CDBNETCL.PanelEx()
    Me.vseStdExclOptions = New System.Windows.Forms.GroupBox()
    Me.optStandardExclusions2 = New System.Windows.Forms.RadioButton()
    Me.optStandardExclusions = New System.Windows.Forms.RadioButton()
    Me.chkSkipCriteriaCount = New System.Windows.Forms.CheckBox()
    Me.pnlParent = New System.Windows.Forms.Panel()
    Me.pnlCenterMain = New System.Windows.Forms.Panel()
    Me.pnlTop = New System.Windows.Forms.Panel()
    Me.vseInfo.SuspendLayout()
    Me.pnlButton.SuspendLayout()
    Me.ButtonPanel1.SuspendLayout()
    CType(Me.SplitContainerForGrid, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SplitContainerForGrid.Panel1.SuspendLayout()
    Me.SplitContainerForGrid.Panel2.SuspendLayout()
    Me.SplitContainerForGrid.SuspendLayout()
    Me.grpCurrentCriteria.SuspendLayout()
    Me.grpStandardExclusions.SuspendLayout()
    Me.pnlBottom.SuspendLayout()
    Me.vseOptions.SuspendLayout()
    Me.vseStdExclOptions.SuspendLayout()
    Me.pnlParent.SuspendLayout()
    Me.pnlCenterMain.SuspendLayout()
    Me.pnlTop.SuspendLayout()
    Me.SuspendLayout()
    '
    'cmdOK
    '
    Me.cmdOK.Location = New System.Drawing.Point(8, 10)
    Me.cmdOK.Name = "cmdOK"
    Me.cmdOK.Size = New System.Drawing.Size(96, 27)
    Me.cmdOK.TabIndex = 0
    Me.cmdOK.Text = "OK"
    Me.cmdOK.UseVisualStyleBackColor = True
    '
    'cmdUpdate
    '
    Me.cmdUpdate.Location = New System.Drawing.Point(8, 195)
    Me.cmdUpdate.Name = "cmdUpdate"
    Me.cmdUpdate.Size = New System.Drawing.Size(96, 27)
    Me.cmdUpdate.TabIndex = 11
    Me.cmdUpdate.Text = "&Update..."
    Me.cmdUpdate.UseVisualStyleBackColor = True
    '
    'vseInfo
    '
    Me.vseInfo.Controls.Add(Me.lblCurrentCriteria)
    Me.vseInfo.Controls.Add(Me.lblCurrentList)
    Me.vseInfo.Dock = System.Windows.Forms.DockStyle.Top
    Me.vseInfo.Location = New System.Drawing.Point(0, 0)
    Me.vseInfo.Name = "vseInfo"
    Me.vseInfo.Size = New System.Drawing.Size(480, 66)
    Me.vseInfo.TabIndex = 1
    Me.vseInfo.TabStop = False
    '
    'lblCurrentCriteria
    '
    Me.lblCurrentCriteria.AutoSize = True
    Me.lblCurrentCriteria.Location = New System.Drawing.Point(10, 42)
    Me.lblCurrentCriteria.Name = "lblCurrentCriteria"
    Me.lblCurrentCriteria.Size = New System.Drawing.Size(101, 13)
    Me.lblCurrentCriteria.TabIndex = 1
    Me.lblCurrentCriteria.Text = "No Criteria Selected"
    '
    'lblCurrentList
    '
    Me.lblCurrentList.AutoSize = True
    Me.lblCurrentList.Location = New System.Drawing.Point(10, 16)
    Me.lblCurrentList.Name = "lblCurrentList"
    Me.lblCurrentList.Size = New System.Drawing.Size(85, 13)
    Me.lblCurrentList.TabIndex = 0
    Me.lblCurrentList.Text = "No List Selected"
    '
    'cmdCancel
    '
    Me.cmdCancel.Location = New System.Drawing.Point(8, 47)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(96, 27)
    Me.cmdCancel.TabIndex = 3
    Me.cmdCancel.Text = "Cancel"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'cmdCount
    '
    Me.cmdCount.Location = New System.Drawing.Point(8, 343)
    Me.cmdCount.Name = "cmdCount"
    Me.cmdCount.Size = New System.Drawing.Size(96, 27)
    Me.cmdCount.TabIndex = 10
    Me.cmdCount.Text = "&Count"
    Me.cmdCount.UseVisualStyleBackColor = True
    '
    'cmdCriteria
    '
    Me.cmdCriteria.Location = New System.Drawing.Point(8, 84)
    Me.cmdCriteria.Name = "cmdCriteria"
    Me.cmdCriteria.Size = New System.Drawing.Size(96, 27)
    Me.cmdCriteria.TabIndex = 4
    Me.cmdCriteria.Text = "Cri&teria..."
    Me.cmdCriteria.UseVisualStyleBackColor = True
    '
    'cmdClear
    '
    Me.cmdClear.Location = New System.Drawing.Point(8, 306)
    Me.cmdClear.Name = "cmdClear"
    Me.cmdClear.Size = New System.Drawing.Size(96, 27)
    Me.cmdClear.TabIndex = 9
    Me.cmdClear.Text = "Cl&ear"
    Me.cmdClear.UseVisualStyleBackColor = True
    '
    'cmdLists
    '
    Me.cmdLists.Location = New System.Drawing.Point(8, 121)
    Me.cmdLists.Name = "cmdLists"
    Me.cmdLists.Size = New System.Drawing.Size(96, 27)
    Me.cmdLists.TabIndex = 5
    Me.cmdLists.Text = "&Lists..."
    Me.cmdLists.UseVisualStyleBackColor = True
    '
    'cmdSaveCriteria
    '
    Me.cmdSaveCriteria.Location = New System.Drawing.Point(8, 269)
    Me.cmdSaveCriteria.Name = "cmdSaveCriteria"
    Me.cmdSaveCriteria.Size = New System.Drawing.Size(96, 27)
    Me.cmdSaveCriteria.TabIndex = 8
    Me.cmdSaveCriteria.Text = "Sa&ve Criteria"
    Me.cmdSaveCriteria.UseVisualStyleBackColor = True
    '
    'cmdAdd
    '
    Me.cmdAdd.Location = New System.Drawing.Point(8, 158)
    Me.cmdAdd.Name = "cmdAdd"
    Me.cmdAdd.Size = New System.Drawing.Size(96, 27)
    Me.cmdAdd.TabIndex = 6
    Me.cmdAdd.Text = "&Add..."
    Me.cmdAdd.UseVisualStyleBackColor = True
    '
    'cmdDelete
    '
    Me.cmdDelete.Location = New System.Drawing.Point(8, 232)
    Me.cmdDelete.Name = "cmdDelete"
    Me.cmdDelete.Size = New System.Drawing.Size(96, 27)
    Me.cmdDelete.TabIndex = 7
    Me.cmdDelete.Text = "&Delete"
    Me.cmdDelete.UseVisualStyleBackColor = True
    '
    'pnlButton
    '
    Me.pnlButton.Controls.Add(Me.ButtonPanel1)
    Me.pnlButton.Dock = System.Windows.Forms.DockStyle.Right
    Me.pnlButton.Location = New System.Drawing.Point(490, 0)
    Me.pnlButton.Name = "pnlButton"
    Me.pnlButton.Size = New System.Drawing.Size(116, 521)
    Me.pnlButton.TabIndex = 12
    '
    'ButtonPanel1
    '
    Me.ButtonPanel1.Controls.Add(Me.cmdOK)
    Me.ButtonPanel1.Controls.Add(Me.cmdCancel)
    Me.ButtonPanel1.Controls.Add(Me.cmdCriteria)
    Me.ButtonPanel1.Controls.Add(Me.cmdLists)
    Me.ButtonPanel1.Controls.Add(Me.cmdAdd)
    Me.ButtonPanel1.Controls.Add(Me.cmdUpdate)
    Me.ButtonPanel1.Controls.Add(Me.cmdDelete)
    Me.ButtonPanel1.Controls.Add(Me.cmdSaveCriteria)
    Me.ButtonPanel1.Controls.Add(Me.cmdClear)
    Me.ButtonPanel1.Controls.Add(Me.cmdCount)
    Me.ButtonPanel1.Dock = System.Windows.Forms.DockStyle.Fill
    Me.ButtonPanel1.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsCenter
    Me.ButtonPanel1.Location = New System.Drawing.Point(0, 0)
    Me.ButtonPanel1.Name = "ButtonPanel1"
    Me.ButtonPanel1.Size = New System.Drawing.Size(112, 521)
    Me.ButtonPanel1.TabIndex = 0
    '
    'timPolling
    '
    Me.timPolling.Interval = 30000
    '
    'SplitContainerForGrid
    '
    Me.SplitContainerForGrid.Dock = System.Windows.Forms.DockStyle.Fill
    Me.SplitContainerForGrid.Location = New System.Drawing.Point(0, 0)
    Me.SplitContainerForGrid.Name = "SplitContainerForGrid"
    Me.SplitContainerForGrid.Orientation = System.Windows.Forms.Orientation.Horizontal
    '
    'SplitContainerForGrid.Panel1
    '
    Me.SplitContainerForGrid.Panel1.Controls.Add(Me.grpCurrentCriteria)
    '
    'SplitContainerForGrid.Panel2
    '
    Me.SplitContainerForGrid.Panel2.Controls.Add(Me.grpStandardExclusions)
    Me.SplitContainerForGrid.Size = New System.Drawing.Size(480, 355)
    Me.SplitContainerForGrid.SplitterDistance = 170
    Me.SplitContainerForGrid.TabIndex = 4
    '
    'grpCurrentCriteria
    '
    Me.grpCurrentCriteria.Controls.Add(Me.dgrCriteria)
    Me.grpCurrentCriteria.Dock = System.Windows.Forms.DockStyle.Fill
    Me.grpCurrentCriteria.Location = New System.Drawing.Point(0, 0)
    Me.grpCurrentCriteria.Name = "grpCurrentCriteria"
    Me.grpCurrentCriteria.Size = New System.Drawing.Size(480, 170)
    Me.grpCurrentCriteria.TabIndex = 0
    Me.grpCurrentCriteria.TabStop = False
    Me.grpCurrentCriteria.Text = "Current Criteria"
    '
    'dgrCriteria
    '
    Me.dgrCriteria.AccessibleName = "Display Grid"
    Me.dgrCriteria.ActiveColumn = 0
    Me.dgrCriteria.AllowColumnResize = True
    Me.dgrCriteria.AllowSorting = True
    Me.dgrCriteria.AutoSetHeight = False
    Me.dgrCriteria.AutoSetRowHeight = False
    Me.dgrCriteria.DisplayTitle = Nothing
    Me.dgrCriteria.Dock = System.Windows.Forms.DockStyle.Fill
    Me.dgrCriteria.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!)
    Me.dgrCriteria.Location = New System.Drawing.Point(3, 16)
    Me.dgrCriteria.MaintenanceDesc = Nothing
    Me.dgrCriteria.MaxGridRows = 6
    Me.dgrCriteria.MultipleSelect = False
    Me.dgrCriteria.Name = "dgrCriteria"
    Me.dgrCriteria.RowCount = 10
    Me.dgrCriteria.ShowIfEmpty = False
    Me.dgrCriteria.Size = New System.Drawing.Size(474, 151)
    Me.dgrCriteria.SuppressHyperLinkFormat = False
    Me.dgrCriteria.TabIndex = 2
    '
    'grpStandardExclusions
    '
    Me.grpStandardExclusions.Controls.Add(Me.dgrDefault)
    Me.grpStandardExclusions.Dock = System.Windows.Forms.DockStyle.Fill
    Me.grpStandardExclusions.Location = New System.Drawing.Point(0, 0)
    Me.grpStandardExclusions.Name = "grpStandardExclusions"
    Me.grpStandardExclusions.Size = New System.Drawing.Size(480, 181)
    Me.grpStandardExclusions.TabIndex = 4
    Me.grpStandardExclusions.TabStop = False
    Me.grpStandardExclusions.Text = "Standard Exclusions"
    '
    'dgrDefault
    '
    Me.dgrDefault.AccessibleName = "Display Grid"
    Me.dgrDefault.ActiveColumn = 0
    Me.dgrDefault.AllowColumnResize = True
    Me.dgrDefault.AllowSorting = True
    Me.dgrDefault.AutoSetHeight = False
    Me.dgrDefault.AutoSetRowHeight = False
    Me.dgrDefault.DisplayTitle = Nothing
    Me.dgrDefault.Dock = System.Windows.Forms.DockStyle.Fill
    Me.dgrDefault.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!)
    Me.dgrDefault.Location = New System.Drawing.Point(3, 16)
    Me.dgrDefault.MaintenanceDesc = Nothing
    Me.dgrDefault.MaxGridRows = 6
    Me.dgrDefault.MultipleSelect = False
    Me.dgrDefault.Name = "dgrDefault"
    Me.dgrDefault.RowCount = 10
    Me.dgrDefault.ShowIfEmpty = False
    Me.dgrDefault.Size = New System.Drawing.Size(474, 162)
    Me.dgrDefault.SuppressHyperLinkFormat = False
    Me.dgrDefault.TabIndex = 3
    '
    'pnlBottom
    '
    Me.pnlBottom.Controls.Add(Me.vseOptions)
    Me.pnlBottom.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.pnlBottom.Location = New System.Drawing.Point(5, 426)
    Me.pnlBottom.Name = "pnlBottom"
    Me.pnlBottom.Size = New System.Drawing.Size(480, 90)
    Me.pnlBottom.TabIndex = 15
    '
    'vseOptions
    '
    Me.vseOptions.BackColor = System.Drawing.Color.Transparent
    Me.vseOptions.Controls.Add(Me.vseStdExclOptions)
    Me.vseOptions.Controls.Add(Me.chkSkipCriteriaCount)
    Me.vseOptions.Dock = System.Windows.Forms.DockStyle.Fill
    Me.vseOptions.Location = New System.Drawing.Point(0, 0)
    Me.vseOptions.Name = "vseOptions"
    Me.vseOptions.Size = New System.Drawing.Size(480, 90)
    Me.vseOptions.TabIndex = 2
    '
    'vseStdExclOptions
    '
    Me.vseStdExclOptions.Controls.Add(Me.optStandardExclusions2)
    Me.vseStdExclOptions.Controls.Add(Me.optStandardExclusions)
    Me.vseStdExclOptions.Dock = System.Windows.Forms.DockStyle.Top
    Me.vseStdExclOptions.Location = New System.Drawing.Point(0, 0)
    Me.vseStdExclOptions.Name = "vseStdExclOptions"
    Me.vseStdExclOptions.Size = New System.Drawing.Size(480, 62)
    Me.vseStdExclOptions.TabIndex = 0
    Me.vseStdExclOptions.TabStop = False
    '
    'optStandardExclusions2
    '
    Me.optStandardExclusions2.AutoSize = True
    Me.optStandardExclusions2.Location = New System.Drawing.Point(7, 37)
    Me.optStandardExclusions2.Name = "optStandardExclusions2"
    Me.optStandardExclusions2.Size = New System.Drawing.Size(180, 17)
    Me.optStandardExclusions2.TabIndex = 1
    Me.optStandardExclusions2.TabStop = True
    Me.optStandardExclusions2.Text = "Pre-process Standard Exclusions"
    Me.optStandardExclusions2.UseVisualStyleBackColor = True
    '
    'optStandardExclusions
    '
    Me.optStandardExclusions.AutoSize = True
    Me.optStandardExclusions.Location = New System.Drawing.Point(7, 14)
    Me.optStandardExclusions.Name = "optStandardExclusions"
    Me.optStandardExclusions.Size = New System.Drawing.Size(256, 17)
    Me.optStandardExclusions.TabIndex = 0
    Me.optStandardExclusions.TabStop = True
    Me.optStandardExclusions.Text = "Process Standard Exclusions with Current Criteria"
    Me.optStandardExclusions.UseVisualStyleBackColor = True
    '
    'chkSkipCriteriaCount
    '
    Me.chkSkipCriteriaCount.AutoSize = True
    Me.chkSkipCriteriaCount.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.chkSkipCriteriaCount.Location = New System.Drawing.Point(0, 73)
    Me.chkSkipCriteriaCount.Name = "chkSkipCriteriaCount"
    Me.chkSkipCriteriaCount.Size = New System.Drawing.Size(480, 17)
    Me.chkSkipCriteriaCount.TabIndex = 1
    Me.chkSkipCriteriaCount.Text = "Skip Counting of each Criteria Line"
    Me.chkSkipCriteriaCount.UseVisualStyleBackColor = True
    '
    'pnlParent
    '
    Me.pnlParent.Controls.Add(Me.pnlCenterMain)
    Me.pnlParent.Controls.Add(Me.pnlTop)
    Me.pnlParent.Controls.Add(Me.pnlBottom)
    Me.pnlParent.Dock = System.Windows.Forms.DockStyle.Fill
    Me.pnlParent.Location = New System.Drawing.Point(0, 0)
    Me.pnlParent.Name = "pnlParent"
    Me.pnlParent.Padding = New System.Windows.Forms.Padding(5)
    Me.pnlParent.Size = New System.Drawing.Size(490, 521)
    Me.pnlParent.TabIndex = 14
    '
    'pnlCenterMain
    '
    Me.pnlCenterMain.BackColor = System.Drawing.SystemColors.Control
    Me.pnlCenterMain.Controls.Add(Me.SplitContainerForGrid)
    Me.pnlCenterMain.Dock = System.Windows.Forms.DockStyle.Fill
    Me.pnlCenterMain.Location = New System.Drawing.Point(5, 71)
    Me.pnlCenterMain.Name = "pnlCenterMain"
    Me.pnlCenterMain.Size = New System.Drawing.Size(480, 355)
    Me.pnlCenterMain.TabIndex = 15
    '
    'pnlTop
    '
    Me.pnlTop.Controls.Add(Me.vseInfo)
    Me.pnlTop.Dock = System.Windows.Forms.DockStyle.Top
    Me.pnlTop.Location = New System.Drawing.Point(5, 5)
    Me.pnlTop.Name = "pnlTop"
    Me.pnlTop.Size = New System.Drawing.Size(480, 66)
    Me.pnlTop.TabIndex = 15
    '
    'frmEditCriteria
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(606, 521)
    Me.Controls.Add(Me.pnlParent)
    Me.Controls.Add(Me.pnlButton)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmEditCriteria"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    Me.Text = "frmEditCriteria"
    Me.vseInfo.ResumeLayout(False)
    Me.vseInfo.PerformLayout()
    Me.pnlButton.ResumeLayout(False)
    Me.ButtonPanel1.ResumeLayout(False)
    Me.SplitContainerForGrid.Panel1.ResumeLayout(False)
    Me.SplitContainerForGrid.Panel2.ResumeLayout(False)
    CType(Me.SplitContainerForGrid, System.ComponentModel.ISupportInitialize).EndInit()
    Me.SplitContainerForGrid.ResumeLayout(False)
    Me.grpCurrentCriteria.ResumeLayout(False)
    Me.grpStandardExclusions.ResumeLayout(False)
    Me.pnlBottom.ResumeLayout(False)
    Me.vseOptions.ResumeLayout(False)
    Me.vseOptions.PerformLayout()
    Me.vseStdExclOptions.ResumeLayout(False)
    Me.vseStdExclOptions.PerformLayout()
    Me.pnlParent.ResumeLayout(False)
    Me.pnlCenterMain.ResumeLayout(False)
    Me.pnlTop.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents cmdOK As System.Windows.Forms.Button
  Friend WithEvents vseInfo As System.Windows.Forms.GroupBox
  Friend WithEvents lblCurrentList As System.Windows.Forms.Label
  Friend WithEvents lblCurrentCriteria As System.Windows.Forms.Label
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents cmdCriteria As System.Windows.Forms.Button
  Friend WithEvents cmdLists As System.Windows.Forms.Button
  Friend WithEvents cmdAdd As System.Windows.Forms.Button
  Friend WithEvents cmdDelete As System.Windows.Forms.Button
  Friend WithEvents cmdSaveCriteria As System.Windows.Forms.Button
  Friend WithEvents cmdClear As System.Windows.Forms.Button
  Friend WithEvents cmdCount As System.Windows.Forms.Button
  Friend WithEvents cmdUpdate As System.Windows.Forms.Button
  Friend WithEvents pnlButton As System.Windows.Forms.Panel
  Friend WithEvents ButtonPanel1 As CDBNETCL.ButtonPanel
  Friend WithEvents timPolling As System.Windows.Forms.Timer
  Friend WithEvents dgrCriteria As CDBNETCL.DisplayGrid
  Friend WithEvents vseOptions As CDBNETCL.PanelEx
  Friend WithEvents vseStdExclOptions As System.Windows.Forms.GroupBox
  Friend WithEvents optStandardExclusions2 As System.Windows.Forms.RadioButton
  Friend WithEvents optStandardExclusions As System.Windows.Forms.RadioButton
  Friend WithEvents chkSkipCriteriaCount As System.Windows.Forms.CheckBox
  Friend WithEvents dgrDefault As CDBNETCL.DisplayGrid
  Friend WithEvents SplitContainerForGrid As System.Windows.Forms.SplitContainer
  Friend WithEvents pnlParent As System.Windows.Forms.Panel
  Friend WithEvents pnlCenterMain As System.Windows.Forms.Panel
  Friend WithEvents pnlTop As System.Windows.Forms.Panel
  Friend WithEvents pnlBottom As System.Windows.Forms.Panel
  Friend WithEvents grpCurrentCriteria As System.Windows.Forms.GroupBox
  Friend WithEvents grpStandardExclusions As System.Windows.Forms.GroupBox
End Class
