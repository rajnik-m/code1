<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmReportDataSelection
    Inherits System.Windows.Forms.Form

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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmReportDataSelection))
    Me.bpl = New CDBNETCL.ButtonPanel
    Me.cmdOK = New System.Windows.Forms.Button
    Me.cmdSave = New System.Windows.Forms.Button
    Me.cmdDelete = New System.Windows.Forms.Button
    Me.cmdCancel = New System.Windows.Forms.Button
    Me.PanelEx1 = New CDBNETCL.PanelEx
    Me.lblSourceReport = New CDBNETCL.TransparentLabel
    Me.lblReportDataSet = New CDBNETCL.TransparentLabel
    Me.lblReportType = New CDBNETCL.TransparentLabel
    Me.lblOrder = New CDBNETCL.TransparentLabel
    Me.lblAvailableItems = New CDBNETCL.TransparentLabel
    Me.lblSelectedItems = New CDBNETCL.TransparentLabel
    Me.cboReport = New System.Windows.Forms.ComboBox
    Me.cboDataSet = New System.Windows.Forms.ComboBox
    Me.cboOutputType = New System.Windows.Forms.ComboBox
    Me.cboOrder = New System.Windows.Forms.ComboBox
    Me.dgrAvailable = New CDBNETCL.DisplayGrid
    Me.DisplayGrid1 = New CDBNETCL.DisplayGrid
    Me.cmdAdd = New System.Windows.Forms.Button
    Me.Button1 = New System.Windows.Forms.Button
    Me.Button2 = New System.Windows.Forms.Button
    Me.Button3 = New System.Windows.Forms.Button
    Me.bpl.SuspendLayout()
    Me.PanelEx1.SuspendLayout()
    Me.SuspendLayout()
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdOK)
    Me.bpl.Controls.Add(Me.cmdSave)
    Me.bpl.Controls.Add(Me.cmdDelete)
    Me.bpl.Controls.Add(Me.cmdCancel)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bpl.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bpl.Location = New System.Drawing.Point(0, 382)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(635, 39)
    Me.bpl.TabIndex = 0
    '
    'cmdOK
    '
    Me.cmdOK.Location = New System.Drawing.Point(107, 6)
    Me.cmdOK.Name = "cmdOK"
    Me.cmdOK.Size = New System.Drawing.Size(94, 27)
    Me.cmdOK.TabIndex = 0
    Me.cmdOK.Text = "OK"
    Me.cmdOK.UseVisualStyleBackColor = True
    '
    'cmdSave
    '
    Me.cmdSave.Location = New System.Drawing.Point(216, 6)
    Me.cmdSave.Name = "cmdSave"
    Me.cmdSave.Size = New System.Drawing.Size(94, 27)
    Me.cmdSave.TabIndex = 1
    Me.cmdSave.Text = "&Save"
    Me.cmdSave.UseVisualStyleBackColor = True
    '
    'cmdDelete
    '
    Me.cmdDelete.Location = New System.Drawing.Point(325, 6)
    Me.cmdDelete.Name = "cmdDelete"
    Me.cmdDelete.Size = New System.Drawing.Size(94, 27)
    Me.cmdDelete.TabIndex = 2
    Me.cmdDelete.Text = "&Delete"
    Me.cmdDelete.UseVisualStyleBackColor = True
    '
    'cmdCancel
    '
    Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdCancel.Location = New System.Drawing.Point(434, 6)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(94, 27)
    Me.cmdCancel.TabIndex = 3
    Me.cmdCancel.Text = "Cancel"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'PanelEx1
    '
    Me.PanelEx1.BackColor = System.Drawing.Color.Transparent
    Me.PanelEx1.Controls.Add(Me.Button3)
    Me.PanelEx1.Controls.Add(Me.Button2)
    Me.PanelEx1.Controls.Add(Me.Button1)
    Me.PanelEx1.Controls.Add(Me.cmdAdd)
    Me.PanelEx1.Controls.Add(Me.DisplayGrid1)
    Me.PanelEx1.Controls.Add(Me.dgrAvailable)
    Me.PanelEx1.Controls.Add(Me.cboOrder)
    Me.PanelEx1.Controls.Add(Me.cboOutputType)
    Me.PanelEx1.Controls.Add(Me.cboDataSet)
    Me.PanelEx1.Controls.Add(Me.cboReport)
    Me.PanelEx1.Controls.Add(Me.lblSelectedItems)
    Me.PanelEx1.Controls.Add(Me.lblAvailableItems)
    Me.PanelEx1.Controls.Add(Me.lblOrder)
    Me.PanelEx1.Controls.Add(Me.lblReportType)
    Me.PanelEx1.Controls.Add(Me.lblReportDataSet)
    Me.PanelEx1.Controls.Add(Me.lblSourceReport)
    Me.PanelEx1.Dock = System.Windows.Forms.DockStyle.Fill
    Me.PanelEx1.Location = New System.Drawing.Point(0, 0)
    Me.PanelEx1.Name = "PanelEx1"
    Me.PanelEx1.Size = New System.Drawing.Size(635, 382)
    Me.PanelEx1.TabIndex = 1
    '
    'lblSourceReport
    '
    Me.lblSourceReport.AutoSize = True
    Me.lblSourceReport.Location = New System.Drawing.Point(12, 9)
    Me.lblSourceReport.Name = "lblSourceReport"
    Me.lblSourceReport.Size = New System.Drawing.Size(79, 13)
    Me.lblSourceReport.TabIndex = 0
    Me.lblSourceReport.Text = "Source Report:"
    Me.lblSourceReport.Visible = False
    '
    'lblReportDataSet
    '
    Me.lblReportDataSet.AutoSize = True
    Me.lblReportDataSet.Location = New System.Drawing.Point(12, 33)
    Me.lblReportDataSet.Name = "lblReportDataSet"
    Me.lblReportDataSet.Size = New System.Drawing.Size(87, 13)
    Me.lblReportDataSet.TabIndex = 1
    Me.lblReportDataSet.Text = "Report Data Set:"
    Me.lblReportDataSet.Visible = False
    '
    'lblReportType
    '
    Me.lblReportType.AutoSize = True
    Me.lblReportType.Location = New System.Drawing.Point(12, 56)
    Me.lblReportType.Name = "lblReportType"
    Me.lblReportType.Size = New System.Drawing.Size(69, 13)
    Me.lblReportType.TabIndex = 2
    Me.lblReportType.Text = "Report Type:"
    Me.lblReportType.Visible = False
    '
    'lblOrder
    '
    Me.lblOrder.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.lblOrder.AutoSize = True
    Me.lblOrder.Location = New System.Drawing.Point(380, 56)
    Me.lblOrder.Name = "lblOrder"
    Me.lblOrder.Size = New System.Drawing.Size(36, 13)
    Me.lblOrder.TabIndex = 3
    Me.lblOrder.Text = "Order:"
    Me.lblOrder.Visible = False
    '
    'lblAvailableItems
    '
    Me.lblAvailableItems.AutoSize = True
    Me.lblAvailableItems.Location = New System.Drawing.Point(12, 87)
    Me.lblAvailableItems.Name = "lblAvailableItems"
    Me.lblAvailableItems.Size = New System.Drawing.Size(78, 13)
    Me.lblAvailableItems.TabIndex = 4
    Me.lblAvailableItems.Text = "Available Items"
    Me.lblAvailableItems.Visible = False
    '
    'lblSelectedItems
    '
    Me.lblSelectedItems.AutoSize = True
    Me.lblSelectedItems.Location = New System.Drawing.Point(292, 87)
    Me.lblSelectedItems.Name = "lblSelectedItems"
    Me.lblSelectedItems.Size = New System.Drawing.Size(77, 13)
    Me.lblSelectedItems.TabIndex = 5
    Me.lblSelectedItems.Text = "Selected Items"
    Me.lblSelectedItems.Visible = False
    '
    'cboReport
    '
    Me.cboReport.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cboReport.FormattingEnabled = True
    Me.cboReport.Location = New System.Drawing.Point(113, 6)
    Me.cboReport.Name = "cboReport"
    Me.cboReport.Size = New System.Drawing.Size(510, 21)
    Me.cboReport.TabIndex = 6
    '
    'cboDataSet
    '
    Me.cboDataSet.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cboDataSet.FormattingEnabled = True
    Me.cboDataSet.Location = New System.Drawing.Point(113, 30)
    Me.cboDataSet.Name = "cboDataSet"
    Me.cboDataSet.Size = New System.Drawing.Size(510, 21)
    Me.cboDataSet.TabIndex = 7
    '
    'cboOutputType
    '
    Me.cboOutputType.FormattingEnabled = True
    Me.cboOutputType.Location = New System.Drawing.Point(113, 53)
    Me.cboOutputType.Name = "cboOutputType"
    Me.cboOutputType.Size = New System.Drawing.Size(189, 21)
    Me.cboOutputType.TabIndex = 8
    '
    'cboOrder
    '
    Me.cboOrder.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cboOrder.FormattingEnabled = True
    Me.cboOrder.Location = New System.Drawing.Point(434, 53)
    Me.cboOrder.Name = "cboOrder"
    Me.cboOrder.Size = New System.Drawing.Size(189, 21)
    Me.cboOrder.TabIndex = 9
    '
    'dgrAvailable
    '
    Me.dgrAvailable.AllowSorting = True
    Me.dgrAvailable.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.dgrAvailable.AutoSetHeight = False
    Me.dgrAvailable.AutoSetRowHeight = False
    Me.dgrAvailable.DisplayTitle = Nothing
    Me.dgrAvailable.Location = New System.Drawing.Point(15, 103)
    Me.dgrAvailable.MaxGridRows = 8
    Me.dgrAvailable.MultipleSelect = False
    Me.dgrAvailable.Name = "dgrAvailable"
    Me.dgrAvailable.ShowIfEmpty = False
    Me.dgrAvailable.Size = New System.Drawing.Size(174, 270)
    Me.dgrAvailable.TabIndex = 10
    '
    'DisplayGrid1
    '
    Me.DisplayGrid1.AllowSorting = True
    Me.DisplayGrid1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.DisplayGrid1.AutoSetHeight = False
    Me.DisplayGrid1.AutoSetRowHeight = False
    Me.DisplayGrid1.DisplayTitle = Nothing
    Me.DisplayGrid1.Location = New System.Drawing.Point(295, 103)
    Me.DisplayGrid1.MaxGridRows = 8
    Me.DisplayGrid1.MultipleSelect = False
    Me.DisplayGrid1.Name = "DisplayGrid1"
    Me.DisplayGrid1.ShowIfEmpty = False
    Me.DisplayGrid1.Size = New System.Drawing.Size(328, 270)
    Me.DisplayGrid1.TabIndex = 11
    '
    'cmdAdd
    '
    Me.cmdAdd.Location = New System.Drawing.Point(195, 103)
    Me.cmdAdd.Name = "cmdAdd"
    Me.cmdAdd.Size = New System.Drawing.Size(94, 27)
    Me.cmdAdd.TabIndex = 12
    Me.cmdAdd.Text = "&Add >>"
    Me.cmdAdd.UseVisualStyleBackColor = True
    '
    'Button1
    '
    Me.Button1.Location = New System.Drawing.Point(195, 136)
    Me.Button1.Name = "Button1"
    Me.Button1.Size = New System.Drawing.Size(94, 27)
    Me.Button1.TabIndex = 13
    Me.Button1.Text = "<< &Remove"
    Me.Button1.UseVisualStyleBackColor = True
    '
    'Button2
    '
    Me.Button2.Location = New System.Drawing.Point(195, 201)
    Me.Button2.Name = "Button2"
    Me.Button2.Size = New System.Drawing.Size(94, 27)
    Me.Button2.TabIndex = 14
    Me.Button2.Text = "A&dd All >>"
    Me.Button2.UseVisualStyleBackColor = True
    '
    'Button3
    '
    Me.Button3.Location = New System.Drawing.Point(195, 234)
    Me.Button3.Name = "Button3"
    Me.Button3.Size = New System.Drawing.Size(94, 27)
    Me.Button3.TabIndex = 15
    Me.Button3.Text = "<< Re&move All"
    Me.Button3.UseVisualStyleBackColor = True
    '
    'frmReportDataSelection
    '
    Me.AcceptButton = Me.cmdOK
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.cmdCancel
    Me.ClientSize = New System.Drawing.Size(635, 421)
    Me.Controls.Add(Me.PanelEx1)
    Me.Controls.Add(Me.bpl)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmReportDataSelection"
    Me.Text = "Report Data Selection"
    Me.bpl.ResumeLayout(False)
    Me.PanelEx1.ResumeLayout(False)
    Me.PanelEx1.PerformLayout()
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents bpl As CDBNETCL.ButtonPanel
  Friend WithEvents cmdDelete As System.Windows.Forms.Button
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents cmdOK As System.Windows.Forms.Button
  Friend WithEvents cmdSave As System.Windows.Forms.Button
  Friend WithEvents PanelEx1 As CDBNETCL.PanelEx
  Friend WithEvents lblSelectedItems As CDBNETCL.TransparentLabel
  Friend WithEvents lblAvailableItems As CDBNETCL.TransparentLabel
  Friend WithEvents lblOrder As CDBNETCL.TransparentLabel
  Friend WithEvents lblReportType As CDBNETCL.TransparentLabel
  Friend WithEvents lblReportDataSet As CDBNETCL.TransparentLabel
  Friend WithEvents lblSourceReport As CDBNETCL.TransparentLabel
  Friend WithEvents cboOutputType As System.Windows.Forms.ComboBox
  Friend WithEvents cboDataSet As System.Windows.Forms.ComboBox
  Friend WithEvents cboReport As System.Windows.Forms.ComboBox
  Friend WithEvents cboOrder As System.Windows.Forms.ComboBox
  Friend WithEvents dgrAvailable As CDBNETCL.DisplayGrid
  Friend WithEvents DisplayGrid1 As CDBNETCL.DisplayGrid
  Friend WithEvents cmdAdd As System.Windows.Forms.Button
  Friend WithEvents Button3 As System.Windows.Forms.Button
  Friend WithEvents Button2 As System.Windows.Forms.Button
  Friend WithEvents Button1 As System.Windows.Forms.Button
End Class
