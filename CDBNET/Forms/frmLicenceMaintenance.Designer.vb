<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmLicenceMaintenance
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmLicenceMaintenance))
    Me.dgr = New CDBNETCL.DisplayGrid()
    Me.bpl = New CDBNETCL.ButtonPanel()
    Me.cmdDeDup = New System.Windows.Forms.Button()
    Me.cmdRefresh = New System.Windows.Forms.Button()
    Me.cmdClose = New System.Windows.Forms.Button()
    Me.lblModuleName = New System.Windows.Forms.Label()
    Me.lblUser = New System.Windows.Forms.Label()
    Me.lblNoOfLicences = New System.Windows.Forms.Label()
    Me.chkNamedUser = New System.Windows.Forms.CheckBox()
    Me.cboModuleName = New System.Windows.Forms.ComboBox()
    Me.cboUser = New System.Windows.Forms.ComboBox()
    Me.lblActiveUsers = New System.Windows.Forms.Label()
    Me.bpl.SuspendLayout()
    Me.SuspendLayout()
    '
    'dgr
    '
    Me.dgr.AccessibleName = "Display Grid"
    Me.dgr.ActiveColumn = 0
    Me.dgr.AllowSorting = True
    Me.dgr.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.dgr.AutoSetHeight = False
    Me.dgr.AutoSetRowHeight = False
    Me.dgr.DisplayTitle = Nothing
    Me.dgr.Location = New System.Drawing.Point(12, 112)
    Me.dgr.MaintenanceDesc = Nothing
    Me.dgr.MaxGridRows = 6
    Me.dgr.MultipleSelect = False
    Me.dgr.Name = "dgr"
    Me.dgr.RowCount = 10
    Me.dgr.ShowIfEmpty = False
    Me.dgr.Size = New System.Drawing.Size(627, 287)
    Me.dgr.SuppressHyperLinkFormat = False
    Me.dgr.TabIndex = 0
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdDeDup)
    Me.bpl.Controls.Add(Me.cmdRefresh)
    Me.bpl.Controls.Add(Me.cmdClose)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bpl.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bpl.Location = New System.Drawing.Point(0, 405)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(651, 39)
    Me.bpl.TabIndex = 2
    '
    'cmdDeDup
    '
    Me.cmdDeDup.Location = New System.Drawing.Point(166, 6)
    Me.cmdDeDup.Name = "cmdDeDup"
    Me.cmdDeDup.Size = New System.Drawing.Size(96, 27)
    Me.cmdDeDup.TabIndex = 0
    Me.cmdDeDup.Text = "&DeDup"
    Me.cmdDeDup.UseVisualStyleBackColor = True
    '
    'cmdRefresh
    '
    Me.cmdRefresh.Location = New System.Drawing.Point(277, 6)
    Me.cmdRefresh.Name = "cmdRefresh"
    Me.cmdRefresh.Size = New System.Drawing.Size(96, 27)
    Me.cmdRefresh.TabIndex = 1
    Me.cmdRefresh.Text = "&Refresh"
    Me.cmdRefresh.UseVisualStyleBackColor = True
    '
    'cmdClose
    '
    Me.cmdClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdClose.Location = New System.Drawing.Point(388, 6)
    Me.cmdClose.Name = "cmdClose"
    Me.cmdClose.Size = New System.Drawing.Size(96, 27)
    Me.cmdClose.TabIndex = 0
    Me.cmdClose.Text = "Close"
    Me.cmdClose.UseVisualStyleBackColor = True
    '
    'lblModuleName
    '
    Me.lblModuleName.AutoSize = True
    Me.lblModuleName.Location = New System.Drawing.Point(12, 18)
    Me.lblModuleName.Name = "lblModuleName"
    Me.lblModuleName.Size = New System.Drawing.Size(76, 13)
    Me.lblModuleName.TabIndex = 3
    Me.lblModuleName.Text = "Module Name:"
    '
    'lblUser
    '
    Me.lblUser.AutoSize = True
    Me.lblUser.Location = New System.Drawing.Point(12, 51)
    Me.lblUser.Name = "lblUser"
    Me.lblUser.Size = New System.Drawing.Size(32, 13)
    Me.lblUser.TabIndex = 5
    Me.lblUser.Text = "User:"
    '
    'lblNoOfLicences
    '
    Me.lblNoOfLicences.AutoSize = True
    Me.lblNoOfLicences.Location = New System.Drawing.Point(430, 18)
    Me.lblNoOfLicences.Name = "lblNoOfLicences"
    Me.lblNoOfLicences.Size = New System.Drawing.Size(0, 13)
    Me.lblNoOfLicences.TabIndex = 6
    '
    'chkNamedUser
    '
    Me.chkNamedUser.AutoSize = True
    Me.chkNamedUser.Location = New System.Drawing.Point(433, 51)
    Me.chkNamedUser.Name = "chkNamedUser"
    Me.chkNamedUser.Size = New System.Drawing.Size(138, 17)
    Me.chkNamedUser.TabIndex = 9
    Me.chkNamedUser.Text = "Named User for Module"
    Me.chkNamedUser.UseVisualStyleBackColor = True
    '
    'cboModuleName
    '
    Me.cboModuleName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cboModuleName.FormattingEnabled = True
    Me.cboModuleName.Location = New System.Drawing.Point(117, 15)
    Me.cboModuleName.Name = "cboModuleName"
    Me.cboModuleName.Size = New System.Drawing.Size(256, 21)
    Me.cboModuleName.TabIndex = 10
    '
    'cboUser
    '
    Me.cboUser.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cboUser.FormattingEnabled = True
    Me.cboUser.Location = New System.Drawing.Point(117, 48)
    Me.cboUser.Name = "cboUser"
    Me.cboUser.Size = New System.Drawing.Size(256, 21)
    Me.cboUser.TabIndex = 11
    '
    'lblActiveUsers
    '
    Me.lblActiveUsers.AutoSize = True
    Me.lblActiveUsers.Location = New System.Drawing.Point(12, 86)
    Me.lblActiveUsers.Name = "lblActiveUsers"
    Me.lblActiveUsers.Size = New System.Drawing.Size(0, 13)
    Me.lblActiveUsers.TabIndex = 12
    '
    'frmLicenceMaintenance
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.cmdClose
    Me.ClientSize = New System.Drawing.Size(651, 444)
    Me.Controls.Add(Me.lblActiveUsers)
    Me.Controls.Add(Me.cboUser)
    Me.Controls.Add(Me.cboModuleName)
    Me.Controls.Add(Me.chkNamedUser)
    Me.Controls.Add(Me.lblNoOfLicences)
    Me.Controls.Add(Me.lblUser)
    Me.Controls.Add(Me.lblModuleName)
    Me.Controls.Add(Me.bpl)
    Me.Controls.Add(Me.dgr)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmLicenceMaintenance"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    Me.Text = "Licence Maintenance"
    Me.bpl.ResumeLayout(False)
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents dgr As CDBNETCL.DisplayGrid
  Friend WithEvents bpl As CDBNETCL.ButtonPanel
  Friend WithEvents cmdDeDup As System.Windows.Forms.Button
  Friend WithEvents cmdRefresh As System.Windows.Forms.Button
  Friend WithEvents cmdClose As System.Windows.Forms.Button
  Friend WithEvents lblModuleName As System.Windows.Forms.Label
  Friend WithEvents lblUser As System.Windows.Forms.Label
  Friend WithEvents lblNoOfLicences As System.Windows.Forms.Label
  Friend WithEvents chkNamedUser As System.Windows.Forms.CheckBox
  Friend WithEvents cboModuleName As System.Windows.Forms.ComboBox
  Friend WithEvents cboUser As System.Windows.Forms.ComboBox
  Friend WithEvents lblActiveUsers As System.Windows.Forms.Label
End Class
