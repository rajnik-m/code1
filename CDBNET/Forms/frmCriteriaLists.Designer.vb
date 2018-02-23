<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCriteriaLists
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCriteriaLists))
    Me.ButtonPanel1 = New CDBNETCL.ButtonPanel()
    Me.cmdOK = New System.Windows.Forms.Button()
    Me.cmdCancel = New System.Windows.Forms.Button()
    Me.cmdSelect = New System.Windows.Forms.Button()
    Me.cmdUpdate = New System.Windows.Forms.Button()
    Me.cmdDelete = New System.Windows.Forms.Button()
    Me.cmdClear = New System.Windows.Forms.Button()
    Me.Panel1 = New System.Windows.Forms.Panel()
    Me.Panel2 = New System.Windows.Forms.Panel()
    Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
    Me.lblMessage = New System.Windows.Forms.Label()
    Me.vseGrid = New System.Windows.Forms.GroupBox()
    Me.dgrCriteriaSets = New CDBNETCL.DisplayGrid()
    Me.vse3 = New System.Windows.Forms.GroupBox()
    Me.txtLookupDepartment = New CDBNETCL.TextLookupBox()
    Me.txtOwner = New System.Windows.Forms.TextBox()
    Me.txtDescription = New System.Windows.Forms.TextBox()
    Me.lblOwner = New System.Windows.Forms.Label()
    Me.lblDepartment = New System.Windows.Forms.Label()
    Me.lblDescription = New System.Windows.Forms.Label()
    Me.ButtonPanel1.SuspendLayout()
    Me.Panel1.SuspendLayout()
    Me.Panel2.SuspendLayout()
    CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SplitContainer1.Panel1.SuspendLayout()
    Me.SplitContainer1.Panel2.SuspendLayout()
    Me.SplitContainer1.SuspendLayout()
    Me.vseGrid.SuspendLayout()
    Me.vse3.SuspendLayout()
    Me.SuspendLayout()
    '
    'ButtonPanel1
    '
    Me.ButtonPanel1.Controls.Add(Me.cmdOK)
    Me.ButtonPanel1.Controls.Add(Me.cmdCancel)
    Me.ButtonPanel1.Controls.Add(Me.cmdSelect)
    Me.ButtonPanel1.Controls.Add(Me.cmdUpdate)
    Me.ButtonPanel1.Controls.Add(Me.cmdDelete)
    Me.ButtonPanel1.Controls.Add(Me.cmdClear)
    Me.ButtonPanel1.Dock = System.Windows.Forms.DockStyle.Fill
    Me.ButtonPanel1.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsCenter
    Me.ButtonPanel1.Location = New System.Drawing.Point(0, 0)
    Me.ButtonPanel1.Name = "ButtonPanel1"
    Me.ButtonPanel1.Size = New System.Drawing.Size(112, 409)
    Me.ButtonPanel1.TabIndex = 0
    '
    'cmdOK
    '
    Me.cmdOK.Location = New System.Drawing.Point(8, 10)
    Me.cmdOK.Name = "cmdOK"
    Me.cmdOK.Size = New System.Drawing.Size(96, 27)
    Me.cmdOK.TabIndex = 1
    Me.cmdOK.Text = "OK"
    Me.cmdOK.UseVisualStyleBackColor = True
    '
    'cmdCancel
    '
    Me.cmdCancel.Location = New System.Drawing.Point(8, 47)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(96, 27)
    Me.cmdCancel.TabIndex = 2
    Me.cmdCancel.Text = "Cancel"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'cmdSelect
    '
    Me.cmdSelect.Location = New System.Drawing.Point(8, 84)
    Me.cmdSelect.Name = "cmdSelect"
    Me.cmdSelect.Size = New System.Drawing.Size(96, 27)
    Me.cmdSelect.TabIndex = 3
    Me.cmdSelect.Text = "Select"
    Me.cmdSelect.UseVisualStyleBackColor = True
    '
    'cmdUpdate
    '
    Me.cmdUpdate.Location = New System.Drawing.Point(8, 121)
    Me.cmdUpdate.Name = "cmdUpdate"
    Me.cmdUpdate.Size = New System.Drawing.Size(96, 27)
    Me.cmdUpdate.TabIndex = 4
    Me.cmdUpdate.Text = "Update"
    Me.cmdUpdate.UseVisualStyleBackColor = True
    '
    'cmdDelete
    '
    Me.cmdDelete.Location = New System.Drawing.Point(8, 158)
    Me.cmdDelete.Name = "cmdDelete"
    Me.cmdDelete.Size = New System.Drawing.Size(96, 27)
    Me.cmdDelete.TabIndex = 5
    Me.cmdDelete.Text = "Delete"
    Me.cmdDelete.UseVisualStyleBackColor = True
    '
    'cmdClear
    '
    Me.cmdClear.Location = New System.Drawing.Point(8, 195)
    Me.cmdClear.Name = "cmdClear"
    Me.cmdClear.Size = New System.Drawing.Size(96, 27)
    Me.cmdClear.TabIndex = 6
    Me.cmdClear.Text = "Clear"
    Me.cmdClear.UseVisualStyleBackColor = True
    '
    'Panel1
    '
    Me.Panel1.Controls.Add(Me.ButtonPanel1)
    Me.Panel1.Dock = System.Windows.Forms.DockStyle.Right
    Me.Panel1.Location = New System.Drawing.Point(488, 0)
    Me.Panel1.Name = "Panel1"
    Me.Panel1.Size = New System.Drawing.Size(104, 409)
    Me.Panel1.TabIndex = 1
    '
    'Panel2
    '
    Me.Panel2.Controls.Add(Me.SplitContainer1)
    Me.Panel2.Controls.Add(Me.vse3)
    Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
    Me.Panel2.Location = New System.Drawing.Point(0, 0)
    Me.Panel2.Name = "Panel2"
    Me.Panel2.Size = New System.Drawing.Size(488, 409)
    Me.Panel2.TabIndex = 2
    '
    'SplitContainer1
    '
    Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
    Me.SplitContainer1.Location = New System.Drawing.Point(0, 78)
    Me.SplitContainer1.Name = "SplitContainer1"
    Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
    '
    'SplitContainer1.Panel1
    '
    Me.SplitContainer1.Panel1.Controls.Add(Me.lblMessage)
    Me.SplitContainer1.Panel1MinSize = 10
    '
    'SplitContainer1.Panel2
    '
    Me.SplitContainer1.Panel2.Controls.Add(Me.vseGrid)
    Me.SplitContainer1.Size = New System.Drawing.Size(488, 331)
    Me.SplitContainer1.SplitterDistance = 36
    Me.SplitContainer1.TabIndex = 7
    '
    'lblMessage
    '
    Me.lblMessage.AutoSize = True
    Me.lblMessage.Dock = System.Windows.Forms.DockStyle.Fill
    Me.lblMessage.Location = New System.Drawing.Point(0, 0)
    Me.lblMessage.Name = "lblMessage"
    Me.lblMessage.Size = New System.Drawing.Size(73, 13)
    Me.lblMessage.TabIndex = 3
    Me.lblMessage.Text = "Record Count"
    '
    'vseGrid
    '
    Me.vseGrid.Controls.Add(Me.dgrCriteriaSets)
    Me.vseGrid.Dock = System.Windows.Forms.DockStyle.Fill
    Me.vseGrid.Location = New System.Drawing.Point(0, 0)
    Me.vseGrid.Name = "vseGrid"
    Me.vseGrid.Size = New System.Drawing.Size(488, 291)
    Me.vseGrid.TabIndex = 2
    Me.vseGrid.TabStop = False
    '
    'dgrCriteriaSets
    '
    Me.dgrCriteriaSets.AccessibleName = "Display Grid"
    Me.dgrCriteriaSets.ActiveColumn = 0
    Me.dgrCriteriaSets.AllowSorting = True
    Me.dgrCriteriaSets.AutoSetHeight = False
    Me.dgrCriteriaSets.AutoSetRowHeight = False
    Me.dgrCriteriaSets.DisplayTitle = Nothing
    Me.dgrCriteriaSets.Dock = System.Windows.Forms.DockStyle.Fill
    Me.dgrCriteriaSets.Location = New System.Drawing.Point(3, 16)
    Me.dgrCriteriaSets.MaintenanceDesc = Nothing
    Me.dgrCriteriaSets.MaxGridRows = 6
    Me.dgrCriteriaSets.MultipleSelect = False
    Me.dgrCriteriaSets.Name = "dgrCriteriaSets"
    Me.dgrCriteriaSets.RowCount = 10
    Me.dgrCriteriaSets.ShowIfEmpty = False
    Me.dgrCriteriaSets.Size = New System.Drawing.Size(482, 272)
    Me.dgrCriteriaSets.SuppressHyperLinkFormat = False
    Me.dgrCriteriaSets.TabIndex = 0
    '
    'vse3
    '
    Me.vse3.Controls.Add(Me.txtLookupDepartment)
    Me.vse3.Controls.Add(Me.txtOwner)
    Me.vse3.Controls.Add(Me.txtDescription)
    Me.vse3.Controls.Add(Me.lblOwner)
    Me.vse3.Controls.Add(Me.lblDepartment)
    Me.vse3.Controls.Add(Me.lblDescription)
    Me.vse3.Dock = System.Windows.Forms.DockStyle.Top
    Me.vse3.Location = New System.Drawing.Point(0, 0)
    Me.vse3.Name = "vse3"
    Me.vse3.Size = New System.Drawing.Size(488, 78)
    Me.vse3.TabIndex = 0
    Me.vse3.TabStop = False
    '
    'txtLookupDepartment
    '
    Me.txtLookupDepartment.ActiveOnly = False
    Me.txtLookupDepartment.BackColor = System.Drawing.SystemColors.Control
    Me.txtLookupDepartment.CustomFormNumber = 0
    Me.txtLookupDepartment.Description = ""
    Me.txtLookupDepartment.EnabledProperty = True
    Me.txtLookupDepartment.HasDependancies = False
    Me.txtLookupDepartment.IsDesign = False
    Me.txtLookupDepartment.Location = New System.Drawing.Point(93, 40)
    Me.txtLookupDepartment.MaxLength = 32767
    Me.txtLookupDepartment.MultipleValuesSupported = False
    Me.txtLookupDepartment.Name = "txtLookupDepartment"
    Me.txtLookupDepartment.OriginalText = Nothing
    Me.txtLookupDepartment.ReadOnlyProperty = False
    Me.txtLookupDepartment.Size = New System.Drawing.Size(375, 33)
    Me.txtLookupDepartment.TabIndex = 5
    Me.txtLookupDepartment.TextReadOnly = False
    Me.txtLookupDepartment.TotalWidth = 408
    Me.txtLookupDepartment.ValidationRequired = True
    '
    'txtOwner
    '
    Me.txtOwner.Location = New System.Drawing.Point(366, 12)
    Me.txtOwner.Name = "txtOwner"
    Me.txtOwner.Size = New System.Drawing.Size(101, 20)
    Me.txtOwner.TabIndex = 4
    '
    'txtDescription
    '
    Me.txtDescription.Location = New System.Drawing.Point(92, 12)
    Me.txtDescription.Name = "txtDescription"
    Me.txtDescription.Size = New System.Drawing.Size(209, 20)
    Me.txtDescription.TabIndex = 3
    '
    'lblOwner
    '
    Me.lblOwner.AutoSize = True
    Me.lblOwner.Location = New System.Drawing.Point(307, 15)
    Me.lblOwner.Name = "lblOwner"
    Me.lblOwner.Size = New System.Drawing.Size(41, 13)
    Me.lblOwner.TabIndex = 2
    Me.lblOwner.Text = "Owner:"
    '
    'lblDepartment
    '
    Me.lblDepartment.AutoSize = True
    Me.lblDepartment.Location = New System.Drawing.Point(6, 41)
    Me.lblDepartment.Name = "lblDepartment"
    Me.lblDepartment.Size = New System.Drawing.Size(65, 13)
    Me.lblDepartment.TabIndex = 1
    Me.lblDepartment.Text = "Department:"
    '
    'lblDescription
    '
    Me.lblDescription.AutoSize = True
    Me.lblDescription.Location = New System.Drawing.Point(6, 15)
    Me.lblDescription.Name = "lblDescription"
    Me.lblDescription.Size = New System.Drawing.Size(63, 13)
    Me.lblDescription.TabIndex = 0
    Me.lblDescription.Text = "Description:"
    '
    'frmCriteriaLists
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(592, 409)
    Me.Controls.Add(Me.Panel2)
    Me.Controls.Add(Me.Panel1)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmCriteriaLists"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    Me.Text = "frmCriteriaLists"
    Me.ButtonPanel1.ResumeLayout(False)
    Me.Panel1.ResumeLayout(False)
    Me.Panel2.ResumeLayout(False)
    Me.SplitContainer1.Panel1.ResumeLayout(False)
    Me.SplitContainer1.Panel1.PerformLayout()
    Me.SplitContainer1.Panel2.ResumeLayout(False)
    CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
    Me.SplitContainer1.ResumeLayout(False)
    Me.vseGrid.ResumeLayout(False)
    Me.vse3.ResumeLayout(False)
    Me.vse3.PerformLayout()
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents ButtonPanel1 As CDBNETCL.ButtonPanel
  Friend WithEvents Panel1 As System.Windows.Forms.Panel
  Friend WithEvents cmdOK As System.Windows.Forms.Button
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents cmdClear As System.Windows.Forms.Button
  Friend WithEvents cmdDelete As System.Windows.Forms.Button
  Friend WithEvents cmdUpdate As System.Windows.Forms.Button
  Friend WithEvents cmdSelect As System.Windows.Forms.Button
  Friend WithEvents Panel2 As System.Windows.Forms.Panel
  Friend WithEvents vse3 As System.Windows.Forms.GroupBox
  Friend WithEvents lblDescription As System.Windows.Forms.Label
  Friend WithEvents lblOwner As System.Windows.Forms.Label
  Friend WithEvents lblDepartment As System.Windows.Forms.Label
  Friend WithEvents txtLookupDepartment As CDBNETCL.TextLookupBox
  Friend WithEvents txtOwner As System.Windows.Forms.TextBox
  Friend WithEvents txtDescription As System.Windows.Forms.TextBox
  Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
  Friend WithEvents vseGrid As System.Windows.Forms.GroupBox
  Friend WithEvents dgrCriteriaSets As CDBNETCL.DisplayGrid
  Friend WithEvents lblMessage As System.Windows.Forms.Label
End Class
