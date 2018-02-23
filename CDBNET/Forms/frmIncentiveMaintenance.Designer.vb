<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmIncentiveMaintenance
  Inherits ThemedForm

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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmIncentiveMaintenance))
    Me.tbpFind = New System.Windows.Forms.TabPage()
    Me.PanelContainerRadioButtons = New System.Windows.Forms.Panel()
    Me.optUnFulFilledIncentives = New System.Windows.Forms.RadioButton()
    Me.optFulFilledIncentives = New System.Windows.Forms.RadioButton()
    Me.GroupBox2 = New System.Windows.Forms.GroupBox()
    Me.Label1 = New System.Windows.Forms.Label()
    Me.txtLookupContactNo = New CDBNETCL.TextLookupBox()
    Me.GroupBox1 = New System.Windows.Forms.GroupBox()
    Me.lblPayPlanNo = New System.Windows.Forms.Label()
    Me.txtLookupPayPlanNo = New CDBNETCL.TextLookupBox()
    Me.bpl = New CDBNETCL.ButtonPanel()
    Me.cmdClose = New System.Windows.Forms.Button()
    Me.cmdReset = New System.Windows.Forms.Button()
    Me.cmdFind = New System.Windows.Forms.Button()
    Me.cmdClear = New System.Windows.Forms.Button()
    Me.lblResults = New System.Windows.Forms.Label()
    Me.tbpResults = New System.Windows.Forms.TabPage()
    Me.dgrResults = New CDBNETCL.DisplayGrid()
    Me.tab = New CDBNETCL.TabControl()
    Me.tbpFind.SuspendLayout()
    Me.PanelContainerRadioButtons.SuspendLayout()
    Me.GroupBox2.SuspendLayout()
    Me.GroupBox1.SuspendLayout()
    Me.bpl.SuspendLayout()
    Me.tbpResults.SuspendLayout()
    Me.tab.SuspendLayout()
    Me.SuspendLayout()
    '
    'tbpFind
    '
    Me.tbpFind.AutoScroll = True
    Me.tbpFind.Controls.Add(Me.PanelContainerRadioButtons)
    Me.tbpFind.Controls.Add(Me.GroupBox2)
    Me.tbpFind.Controls.Add(Me.GroupBox1)
    Me.tbpFind.Location = New System.Drawing.Point(4, 26)
    Me.tbpFind.Name = "tbpFind"
    Me.tbpFind.Size = New System.Drawing.Size(567, 275)
    Me.tbpFind.TabIndex = 0
    Me.tbpFind.Text = "Criteria"
    Me.tbpFind.UseVisualStyleBackColor = True
    '
    'PanelContainerRadioButtons
    '
    Me.PanelContainerRadioButtons.Controls.Add(Me.optUnFulFilledIncentives)
    Me.PanelContainerRadioButtons.Controls.Add(Me.optFulFilledIncentives)
    Me.PanelContainerRadioButtons.Location = New System.Drawing.Point(9, 173)
    Me.PanelContainerRadioButtons.Name = "PanelContainerRadioButtons"
    Me.PanelContainerRadioButtons.Size = New System.Drawing.Size(550, 100)
    Me.PanelContainerRadioButtons.TabIndex = 2
    '
    'optUnFulFilledIncentives
    '
    Me.optUnFulFilledIncentives.AutoSize = True
    Me.optUnFulFilledIncentives.Location = New System.Drawing.Point(14, 50)
    Me.optUnFulFilledIncentives.Name = "optUnFulFilledIncentives"
    Me.optUnFulFilledIncentives.Size = New System.Drawing.Size(129, 17)
    Me.optUnFulFilledIncentives.TabIndex = 1
    Me.optUnFulFilledIncentives.TabStop = True
    Me.optUnFulFilledIncentives.Text = "UnFulFilled Incentives"
    Me.optUnFulFilledIncentives.UseVisualStyleBackColor = True
    '
    'optFulFilledIncentives
    '
    Me.optFulFilledIncentives.AutoSize = True
    Me.optFulFilledIncentives.Location = New System.Drawing.Point(14, 18)
    Me.optFulFilledIncentives.Name = "optFulFilledIncentives"
    Me.optFulFilledIncentives.Size = New System.Drawing.Size(115, 17)
    Me.optFulFilledIncentives.TabIndex = 0
    Me.optFulFilledIncentives.TabStop = True
    Me.optFulFilledIncentives.Text = "FulFilled Incentives"
    Me.optFulFilledIncentives.UseVisualStyleBackColor = True
    '
    'GroupBox2
    '
    Me.GroupBox2.Controls.Add(Me.Label1)
    Me.GroupBox2.Controls.Add(Me.txtLookupContactNo)
    Me.GroupBox2.Location = New System.Drawing.Point(9, 94)
    Me.GroupBox2.Name = "GroupBox2"
    Me.GroupBox2.Size = New System.Drawing.Size(550, 72)
    Me.GroupBox2.TabIndex = 1
    Me.GroupBox2.TabStop = False
    Me.GroupBox2.Text = "Contact CreationRelated Incentives"
    '
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.Location = New System.Drawing.Point(19, 32)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(64, 13)
    Me.Label1.TabIndex = 2
    Me.Label1.Text = "Contact No:"
    '
    'txtLookupContactNo
    '
    Me.txtLookupContactNo.ActiveOnly = False
    Me.txtLookupContactNo.BackColor = System.Drawing.Color.Transparent
    Me.txtLookupContactNo.CustomFormNumber = 0
    Me.txtLookupContactNo.Description = ""
    Me.txtLookupContactNo.EnabledProperty = True
    Me.txtLookupContactNo.HasDependancies = False
    Me.txtLookupContactNo.IsDesign = False
    Me.txtLookupContactNo.Location = New System.Drawing.Point(136, 32)
    Me.txtLookupContactNo.MaxLength = 32767
    Me.txtLookupContactNo.MultipleValuesSupported = False
    Me.txtLookupContactNo.Name = "txtLookupContactNo"
    Me.txtLookupContactNo.OriginalText = Nothing
    Me.txtLookupContactNo.ReadOnlyProperty = False
    Me.txtLookupContactNo.Size = New System.Drawing.Size(408, 28)
    Me.txtLookupContactNo.TabIndex = 1
    Me.txtLookupContactNo.TextReadOnly = False
    Me.txtLookupContactNo.TotalWidth = 408
    Me.txtLookupContactNo.ValidationRequired = True
    '
    'GroupBox1
    '
    Me.GroupBox1.Controls.Add(Me.lblPayPlanNo)
    Me.GroupBox1.Controls.Add(Me.txtLookupPayPlanNo)
    Me.GroupBox1.Location = New System.Drawing.Point(9, 16)
    Me.GroupBox1.Name = "GroupBox1"
    Me.GroupBox1.Size = New System.Drawing.Size(550, 72)
    Me.GroupBox1.TabIndex = 0
    Me.GroupBox1.TabStop = False
    Me.GroupBox1.Text = "Payment Plan Related Incentives"
    '
    'lblPayPlanNo
    '
    Me.lblPayPlanNo.AutoSize = True
    Me.lblPayPlanNo.Location = New System.Drawing.Point(19, 36)
    Me.lblPayPlanNo.Name = "lblPayPlanNo"
    Me.lblPayPlanNo.Size = New System.Drawing.Size(69, 13)
    Me.lblPayPlanNo.TabIndex = 1
    Me.lblPayPlanNo.Text = "Pay Plan No:"
    '
    'txtLookupPayPlanNo
    '
    Me.txtLookupPayPlanNo.ActiveOnly = False
    Me.txtLookupPayPlanNo.BackColor = System.Drawing.Color.Transparent
    Me.txtLookupPayPlanNo.CustomFormNumber = 0
    Me.txtLookupPayPlanNo.Description = ""
    Me.txtLookupPayPlanNo.EnabledProperty = True
    Me.txtLookupPayPlanNo.HasDependancies = False
    Me.txtLookupPayPlanNo.IsDesign = False
    Me.txtLookupPayPlanNo.Location = New System.Drawing.Point(136, 32)
    Me.txtLookupPayPlanNo.MaxLength = 32767
    Me.txtLookupPayPlanNo.MultipleValuesSupported = False
    Me.txtLookupPayPlanNo.Name = "txtLookupPayPlanNo"
    Me.txtLookupPayPlanNo.OriginalText = Nothing
    Me.txtLookupPayPlanNo.ReadOnlyProperty = False
    Me.txtLookupPayPlanNo.Size = New System.Drawing.Size(408, 28)
    Me.txtLookupPayPlanNo.TabIndex = 0
    Me.txtLookupPayPlanNo.TextReadOnly = False
    Me.txtLookupPayPlanNo.TotalWidth = 408
    Me.txtLookupPayPlanNo.ValidationRequired = True
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdClose)
    Me.bpl.Controls.Add(Me.cmdReset)
    Me.bpl.Controls.Add(Me.cmdFind)
    Me.bpl.Controls.Add(Me.cmdClear)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bpl.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bpl.Location = New System.Drawing.Point(0, 305)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(575, 39)
    Me.bpl.TabIndex = 3
    '
    'cmdClose
    '
    Me.cmdClose.Location = New System.Drawing.Point(73, 6)
    Me.cmdClose.Name = "cmdClose"
    Me.cmdClose.Size = New System.Drawing.Size(96, 27)
    Me.cmdClose.TabIndex = 8
    Me.cmdClose.Text = "&Close"
    '
    'cmdReset
    '
    Me.cmdReset.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdReset.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdReset.Location = New System.Drawing.Point(184, 6)
    Me.cmdReset.Name = "cmdReset"
    Me.cmdReset.Size = New System.Drawing.Size(96, 27)
    Me.cmdReset.TabIndex = 4
    Me.cmdReset.Text = "&Reset"
    '
    'cmdFind
    '
    Me.cmdFind.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdFind.Location = New System.Drawing.Point(295, 6)
    Me.cmdFind.Name = "cmdFind"
    Me.cmdFind.Size = New System.Drawing.Size(96, 27)
    Me.cmdFind.TabIndex = 2
    Me.cmdFind.Text = "&Find"
    '
    'cmdClear
    '
    Me.cmdClear.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdClear.Location = New System.Drawing.Point(406, 6)
    Me.cmdClear.Name = "cmdClear"
    Me.cmdClear.Size = New System.Drawing.Size(96, 27)
    Me.cmdClear.TabIndex = 3
    Me.cmdClear.Text = "Cle&ar"
    '
    'lblResults
    '
    Me.lblResults.Dock = System.Windows.Forms.DockStyle.Top
    Me.lblResults.Location = New System.Drawing.Point(0, 0)
    Me.lblResults.Name = "lblResults"
    Me.lblResults.Size = New System.Drawing.Size(567, 24)
    Me.lblResults.TabIndex = 1
    Me.lblResults.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    '
    'tbpResults
    '
    Me.tbpResults.Controls.Add(Me.dgrResults)
    Me.tbpResults.Controls.Add(Me.lblResults)
    Me.tbpResults.Location = New System.Drawing.Point(4, 26)
    Me.tbpResults.Name = "tbpResults"
    Me.tbpResults.Size = New System.Drawing.Size(567, 275)
    Me.tbpResults.TabIndex = 1
    Me.tbpResults.Text = "Results"
    Me.tbpResults.UseVisualStyleBackColor = True
    '
    'dgrResults
    '
    Me.dgrResults.AccessibleDescription = "Display List"
    Me.dgrResults.AccessibleName = "Display List"
    Me.dgrResults.AccessibleRole = System.Windows.Forms.AccessibleRole.Table
    Me.dgrResults.ActiveColumn = 0
    Me.dgrResults.AllowSorting = True
    Me.dgrResults.AutoSetHeight = False
    Me.dgrResults.AutoSetRowHeight = False
    Me.dgrResults.DisplayTitle = Nothing
    Me.dgrResults.Dock = System.Windows.Forms.DockStyle.Fill
    Me.dgrResults.Location = New System.Drawing.Point(0, 24)
    Me.dgrResults.MaintenanceDesc = Nothing
    Me.dgrResults.MaxGridRows = 8
    Me.dgrResults.MultipleSelect = False
    Me.dgrResults.Name = "dgrResults"
    Me.dgrResults.RowCount = 10
    Me.dgrResults.ShowIfEmpty = False
    Me.dgrResults.Size = New System.Drawing.Size(567, 251)
    Me.dgrResults.SuppressHyperLinkFormat = False
    Me.dgrResults.TabIndex = 0
    '
    'tab
    '
    Me.tab.Controls.Add(Me.tbpFind)
    Me.tab.Controls.Add(Me.tbpResults)
    Me.tab.Dock = System.Windows.Forms.DockStyle.Fill
    Me.tab.ItemSize = New System.Drawing.Size(71, 22)
    Me.tab.Location = New System.Drawing.Point(0, 0)
    Me.tab.Name = "tab"
    Me.tab.SelectedIndex = 0
    Me.tab.Size = New System.Drawing.Size(575, 305)
    Me.tab.SizeMode = System.Windows.Forms.TabSizeMode.Fixed
    Me.tab.TabIndex = 2
    '
    'frmIncentiveMaintenance
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(575, 344)
    Me.Controls.Add(Me.tab)
    Me.Controls.Add(Me.bpl)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmIncentiveMaintenance"
    Me.Text = "Incentive Maintenance"
    Me.tbpFind.ResumeLayout(False)
    Me.PanelContainerRadioButtons.ResumeLayout(False)
    Me.PanelContainerRadioButtons.PerformLayout()
    Me.GroupBox2.ResumeLayout(False)
    Me.GroupBox2.PerformLayout()
    Me.GroupBox1.ResumeLayout(False)
    Me.GroupBox1.PerformLayout()
    Me.bpl.ResumeLayout(False)
    Me.tbpResults.ResumeLayout(False)
    Me.tab.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents tbpFind As System.Windows.Forms.TabPage
  Protected WithEvents bpl As CDBNETCL.ButtonPanel
  Friend WithEvents cmdFind As System.Windows.Forms.Button
  Friend WithEvents cmdClose As System.Windows.Forms.Button
  Friend WithEvents cmdClear As System.Windows.Forms.Button
  Friend WithEvents cmdReset As System.Windows.Forms.Button
  Friend WithEvents lblResults As System.Windows.Forms.Label
  Friend WithEvents tbpResults As System.Windows.Forms.TabPage
  Protected WithEvents dgrResults As CDBNETCL.DisplayGrid
  Friend WithEvents tab As CDBNETCL.TabControl
  Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
  Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
  Friend WithEvents PanelContainerRadioButtons As System.Windows.Forms.Panel
  Friend WithEvents optFulFilledIncentives As System.Windows.Forms.RadioButton
  Friend WithEvents optUnFulFilledIncentives As System.Windows.Forms.RadioButton
  Friend WithEvents txtLookupContactNo As CDBNETCL.TextLookupBox
  Friend WithEvents lblPayPlanNo As System.Windows.Forms.Label
  Friend WithEvents txtLookupPayPlanNo As CDBNETCL.TextLookupBox
  Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
