<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCLIBrowser
    Inherits CDBNETCL.PersistentForm

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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCLIBrowser))
    Me.epl = New CDBNETCL.EditPanel()
    Me.bpl = New CDBNETCL.ButtonPanel()
    Me.cmdSelect = New System.Windows.Forms.Button()
    Me.cmdSelectOrg = New System.Windows.Forms.Button()
    Me.cmdFind = New System.Windows.Forms.Button()
    Me.cmdClear = New System.Windows.Forms.Button()
    Me.cmdCancel = New System.Windows.Forms.Button()
    Me.spl = New System.Windows.Forms.SplitContainer()
    Me.dgr = New CDBNETCL.DisplayGrid()
    Me.pnlCombo = New CDBNETCL.PanelEx()
    Me.lblAddresses = New System.Windows.Forms.Label()
    Me.cboAddresses = New System.Windows.Forms.ComboBox()
    Me.lblOrganisation = New System.Windows.Forms.Label()
    Me.cboOrganisations = New System.Windows.Forms.ComboBox()
    Me.bpl.SuspendLayout()
    CType(Me.spl, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.spl.Panel1.SuspendLayout()
    Me.spl.Panel2.SuspendLayout()
    Me.spl.SuspendLayout()
    Me.pnlCombo.SuspendLayout()
    Me.SuspendLayout()
    '
    'epl
    '
    Me.epl.AddressChanged = False
    Me.epl.BackColor = System.Drawing.Color.Transparent
    Me.epl.DataChanged = False
    Me.epl.Dock = System.Windows.Forms.DockStyle.Fill
    Me.epl.Location = New System.Drawing.Point(0, 0)
    Me.epl.Name = "epl"
    Me.epl.Recipients = Nothing
    Me.epl.Size = New System.Drawing.Size(777, 71)
    Me.epl.SuppressDrawing = False
    Me.epl.TabIndex = 0
    Me.epl.TabSelectedIndex = 0
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdSelect)
    Me.bpl.Controls.Add(Me.cmdSelectOrg)
    Me.bpl.Controls.Add(Me.cmdFind)
    Me.bpl.Controls.Add(Me.cmdClear)
    Me.bpl.Controls.Add(Me.cmdCancel)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bpl.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bpl.Location = New System.Drawing.Point(0, 423)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(777, 39)
    Me.bpl.TabIndex = 3
    Me.bpl.TabStop = True
    '
    'cmdSelect
    '
    Me.cmdSelect.Location = New System.Drawing.Point(118, 6)
    Me.cmdSelect.Name = "cmdSelect"
    Me.cmdSelect.Size = New System.Drawing.Size(96, 27)
    Me.cmdSelect.TabIndex = 0
    Me.cmdSelect.Text = "&Select Contact"
    Me.cmdSelect.UseVisualStyleBackColor = True
    '
    'cmdSelectOrg
    '
    Me.cmdSelectOrg.Location = New System.Drawing.Point(229, 6)
    Me.cmdSelectOrg.Name = "cmdSelectOrg"
    Me.cmdSelectOrg.Size = New System.Drawing.Size(96, 27)
    Me.cmdSelectOrg.TabIndex = 1
    Me.cmdSelectOrg.Text = "Select &Org"
    Me.cmdSelectOrg.UseVisualStyleBackColor = True
    '
    'cmdFind
    '
    Me.cmdFind.Location = New System.Drawing.Point(340, 6)
    Me.cmdFind.Name = "cmdFind"
    Me.cmdFind.Size = New System.Drawing.Size(96, 27)
    Me.cmdFind.TabIndex = 2
    Me.cmdFind.Text = "Find"
    Me.cmdFind.UseVisualStyleBackColor = True
    '
    'cmdClear
    '
    Me.cmdClear.Location = New System.Drawing.Point(451, 6)
    Me.cmdClear.Name = "cmdClear"
    Me.cmdClear.Size = New System.Drawing.Size(96, 27)
    Me.cmdClear.TabIndex = 3
    Me.cmdClear.Text = "Clear"
    Me.cmdClear.UseVisualStyleBackColor = True
    '
    'cmdCancel
    '
    Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdCancel.Location = New System.Drawing.Point(562, 6)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(96, 27)
    Me.cmdCancel.TabIndex = 4
    Me.cmdCancel.Text = "Cancel"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'spl
    '
    Me.spl.Dock = System.Windows.Forms.DockStyle.Fill
    Me.spl.FixedPanel = System.Windows.Forms.FixedPanel.Panel2
    Me.spl.IsSplitterFixed = True
    Me.spl.Location = New System.Drawing.Point(0, 0)
    Me.spl.Name = "spl"
    Me.spl.Orientation = System.Windows.Forms.Orientation.Horizontal
    '
    'spl.Panel1
    '
    Me.spl.Panel1.Controls.Add(Me.dgr)
    Me.spl.Panel1.Controls.Add(Me.pnlCombo)
    '
    'spl.Panel2
    '
    Me.spl.Panel2.Controls.Add(Me.epl)
    Me.spl.Size = New System.Drawing.Size(777, 423)
    Me.spl.SplitterDistance = 348
    Me.spl.TabIndex = 2
    Me.spl.TabStop = False
    '
    'dgr
    '
    Me.dgr.AccessibleName = "Display Grid"
    Me.dgr.ActiveColumn = 0
    Me.dgr.AllowSorting = True
    Me.dgr.AutoSetHeight = False
    Me.dgr.AutoSetRowHeight = False
    Me.dgr.DisplayTitle = Nothing
    Me.dgr.Dock = System.Windows.Forms.DockStyle.Fill
    Me.dgr.Location = New System.Drawing.Point(0, 97)
    Me.dgr.MaintenanceDesc = Nothing
    Me.dgr.MaxGridRows = 8
    Me.dgr.MultipleSelect = False
    Me.dgr.Name = "dgr"
    Me.dgr.RowCount = 10
    Me.dgr.ShowIfEmpty = False
    Me.dgr.Size = New System.Drawing.Size(777, 251)
    Me.dgr.SuppressHyperLinkFormat = False
    Me.dgr.TabIndex = 1
    '
    'pnlCombo
    '
    Me.pnlCombo.BackColor = System.Drawing.Color.Transparent
    Me.pnlCombo.Controls.Add(Me.lblAddresses)
    Me.pnlCombo.Controls.Add(Me.cboAddresses)
    Me.pnlCombo.Controls.Add(Me.lblOrganisation)
    Me.pnlCombo.Controls.Add(Me.cboOrganisations)
    Me.pnlCombo.Dock = System.Windows.Forms.DockStyle.Top
    Me.pnlCombo.Location = New System.Drawing.Point(0, 0)
    Me.pnlCombo.Name = "pnlCombo"
    Me.pnlCombo.Size = New System.Drawing.Size(777, 97)
    Me.pnlCombo.TabIndex = 0
    Me.pnlCombo.TabStop = True
    '
    'lblAddresses
    '
    Me.lblAddresses.AutoSize = True
    Me.lblAddresses.Location = New System.Drawing.Point(12, 50)
    Me.lblAddresses.Name = "lblAddresses"
    Me.lblAddresses.Size = New System.Drawing.Size(56, 13)
    Me.lblAddresses.TabIndex = 2
    Me.lblAddresses.Text = "Addresses"
    '
    'cboAddresses
    '
    Me.cboAddresses.FormattingEnabled = True
    Me.cboAddresses.Location = New System.Drawing.Point(12, 68)
    Me.cboAddresses.Name = "cboAddresses"
    Me.cboAddresses.Size = New System.Drawing.Size(482, 21)
    Me.cboAddresses.TabIndex = 3
    '
    'lblOrganisation
    '
    Me.lblOrganisation.AutoSize = True
    Me.lblOrganisation.Location = New System.Drawing.Point(12, 6)
    Me.lblOrganisation.Name = "lblOrganisation"
    Me.lblOrganisation.Size = New System.Drawing.Size(71, 13)
    Me.lblOrganisation.TabIndex = 0
    Me.lblOrganisation.Text = "Organisations"
    '
    'cboOrganisations
    '
    Me.cboOrganisations.FormattingEnabled = True
    Me.cboOrganisations.Location = New System.Drawing.Point(12, 25)
    Me.cboOrganisations.Name = "cboOrganisations"
    Me.cboOrganisations.Size = New System.Drawing.Size(482, 21)
    Me.cboOrganisations.TabIndex = 1
    '
    'frmCLIBrowser
    '
    Me.AcceptButton = Me.cmdSelect
    Me.CancelButton = Me.cmdCancel
    Me.ClientSize = New System.Drawing.Size(777, 462)
    Me.Controls.Add(Me.spl)
    Me.Controls.Add(Me.bpl)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmCLIBrowser"
    Me.bpl.ResumeLayout(False)
    Me.spl.Panel1.ResumeLayout(False)
    Me.spl.Panel2.ResumeLayout(False)
    CType(Me.spl, System.ComponentModel.ISupportInitialize).EndInit()
    Me.spl.ResumeLayout(False)
    Me.pnlCombo.ResumeLayout(False)
    Me.pnlCombo.PerformLayout()
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents epl As CDBNETCL.EditPanel
  Friend WithEvents bpl As CDBNETCL.ButtonPanel
  Friend WithEvents cmdSelectOrg As System.Windows.Forms.Button
  Friend WithEvents cmdFind As System.Windows.Forms.Button
  Friend WithEvents cmdClear As System.Windows.Forms.Button
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents spl As System.Windows.Forms.SplitContainer
  Friend WithEvents dgr As CDBNETCL.DisplayGrid
  Friend WithEvents pnlCombo As CDBNETCL.PanelEx
  Friend WithEvents cboAddresses As System.Windows.Forms.ComboBox
  Friend WithEvents lblAddresses As System.Windows.Forms.Label
  Friend WithEvents lblOrganisation As System.Windows.Forms.Label
  Friend WithEvents cboOrganisations As System.Windows.Forms.ComboBox
  Friend WithEvents cmdSelect As System.Windows.Forms.Button

End Class
