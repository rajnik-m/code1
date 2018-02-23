<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmGenMLists
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmGenMLists))
    Me.Panel1 = New System.Windows.Forms.Panel()
    Me.bpl = New CDBNETCL.ButtonPanel()
    Me.cmdOK = New System.Windows.Forms.Button()
    Me.cmdCancel = New System.Windows.Forms.Button()
    Me.cmdSelect = New System.Windows.Forms.Button()
    Me.cmdUpdate = New System.Windows.Forms.Button()
    Me.cmdDelete = New System.Windows.Forms.Button()
    Me.cmdClear = New System.Windows.Forms.Button()
    Me.splt = New System.Windows.Forms.SplitContainer()
    Me.epl = New CDBNETCL.EditPanel()
    Me.lblContents = New System.Windows.Forms.Label()
    Me.dgr = New CDBNETCL.DisplayGrid()
    Me.Panel1.SuspendLayout()
    Me.bpl.SuspendLayout()
    CType(Me.splt, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.splt.Panel1.SuspendLayout()
    Me.splt.Panel2.SuspendLayout()
    Me.splt.SuspendLayout()
    Me.SuspendLayout()
    '
    'Panel1
    '
    Me.Panel1.Controls.Add(Me.bpl)
    Me.Panel1.Dock = System.Windows.Forms.DockStyle.Right
    Me.Panel1.Location = New System.Drawing.Point(543, 0)
    Me.Panel1.Name = "Panel1"
    Me.Panel1.Size = New System.Drawing.Size(112, 372)
    Me.Panel1.TabIndex = 0
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdOK)
    Me.bpl.Controls.Add(Me.cmdCancel)
    Me.bpl.Controls.Add(Me.cmdSelect)
    Me.bpl.Controls.Add(Me.cmdUpdate)
    Me.bpl.Controls.Add(Me.cmdDelete)
    Me.bpl.Controls.Add(Me.cmdClear)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Fill
    Me.bpl.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsCenter
    Me.bpl.Location = New System.Drawing.Point(0, 0)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(112, 372)
    Me.bpl.TabIndex = 0
    '
    'cmdOK
    '
    Me.cmdOK.Location = New System.Drawing.Point(8, 10)
    Me.cmdOK.Name = "cmdOK"
    Me.cmdOK.Size = New System.Drawing.Size(96, 27)
    Me.cmdOK.TabIndex = 6
    Me.cmdOK.Text = "OK"
    Me.cmdOK.UseVisualStyleBackColor = True
    '
    'cmdCancel
    '
    Me.cmdCancel.Location = New System.Drawing.Point(8, 47)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(96, 27)
    Me.cmdCancel.TabIndex = 5
    Me.cmdCancel.Text = "Cancel"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'cmdSelect
    '
    Me.cmdSelect.Location = New System.Drawing.Point(8, 84)
    Me.cmdSelect.Name = "cmdSelect"
    Me.cmdSelect.Size = New System.Drawing.Size(96, 27)
    Me.cmdSelect.TabIndex = 4
    Me.cmdSelect.Text = "&Select"
    Me.cmdSelect.UseVisualStyleBackColor = True
    '
    'cmdUpdate
    '
    Me.cmdUpdate.Location = New System.Drawing.Point(8, 121)
    Me.cmdUpdate.Name = "cmdUpdate"
    Me.cmdUpdate.Size = New System.Drawing.Size(96, 27)
    Me.cmdUpdate.TabIndex = 3
    Me.cmdUpdate.Text = "&Update"
    Me.cmdUpdate.UseVisualStyleBackColor = True
    '
    'cmdDelete
    '
    Me.cmdDelete.Location = New System.Drawing.Point(8, 158)
    Me.cmdDelete.Name = "cmdDelete"
    Me.cmdDelete.Size = New System.Drawing.Size(96, 27)
    Me.cmdDelete.TabIndex = 2
    Me.cmdDelete.Text = "&Delete"
    Me.cmdDelete.UseVisualStyleBackColor = True
    '
    'cmdClear
    '
    Me.cmdClear.Location = New System.Drawing.Point(8, 195)
    Me.cmdClear.Name = "cmdClear"
    Me.cmdClear.Size = New System.Drawing.Size(96, 27)
    Me.cmdClear.TabIndex = 1
    Me.cmdClear.Text = "&Clear"
    Me.cmdClear.UseVisualStyleBackColor = True
    '
    'splt
    '
    Me.splt.Dock = System.Windows.Forms.DockStyle.Fill
    Me.splt.Location = New System.Drawing.Point(0, 0)
    Me.splt.Name = "splt"
    Me.splt.Orientation = System.Windows.Forms.Orientation.Horizontal
    '
    'splt.Panel1
    '
    Me.splt.Panel1.Controls.Add(Me.epl)
    '
    'splt.Panel2
    '
    Me.splt.Panel2.Controls.Add(Me.lblContents)
    Me.splt.Panel2.Controls.Add(Me.dgr)
    Me.splt.Size = New System.Drawing.Size(543, 372)
    Me.splt.SplitterDistance = 114
    Me.splt.TabIndex = 1
    '
    'epl
    '
    Me.epl.AddressChanged = False
    Me.epl.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.epl.BackColor = System.Drawing.Color.Transparent
    Me.epl.DataChanged = False
    Me.epl.Location = New System.Drawing.Point(0, 0)
    Me.epl.Name = "epl"
    Me.epl.Recipients = Nothing
    Me.epl.Size = New System.Drawing.Size(543, 114)
    Me.epl.SuppressDrawing = False
    Me.epl.TabIndex = 0
    Me.epl.TabSelectedIndex = 0
    '
    'lblContents
    '
    Me.lblContents.AutoSize = True
    Me.lblContents.Location = New System.Drawing.Point(0, 0)
    Me.lblContents.Name = "lblContents"
    Me.lblContents.Size = New System.Drawing.Size(0, 13)
    Me.lblContents.TabIndex = 1
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
    Me.dgr.Location = New System.Drawing.Point(0, 36)
    Me.dgr.MaintenanceDesc = Nothing
    Me.dgr.MaxGridRows = 8
    Me.dgr.MultipleSelect = False
    Me.dgr.Name = "dgr"
    Me.dgr.RowCount = 10
    Me.dgr.ShowIfEmpty = False
    Me.dgr.Size = New System.Drawing.Size(543, 177)
    Me.dgr.SuppressHyperLinkFormat = False
    Me.dgr.TabIndex = 0
    '
    'frmGenMLists
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(655, 372)
    Me.Controls.Add(Me.splt)
    Me.Controls.Add(Me.Panel1)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmGenMLists"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    Me.Text = "Previously Defined Lists"
    Me.Panel1.ResumeLayout(False)
    Me.bpl.ResumeLayout(False)
    Me.splt.Panel1.ResumeLayout(False)
    Me.splt.Panel2.ResumeLayout(False)
    Me.splt.Panel2.PerformLayout()
    CType(Me.splt, System.ComponentModel.ISupportInitialize).EndInit()
    Me.splt.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents Panel1 As System.Windows.Forms.Panel
  Friend WithEvents bpl As CDBNETCL.ButtonPanel
  Friend WithEvents cmdOK As System.Windows.Forms.Button
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents cmdSelect As System.Windows.Forms.Button
  Friend WithEvents cmdUpdate As System.Windows.Forms.Button
  Friend WithEvents cmdDelete As System.Windows.Forms.Button
  Friend WithEvents cmdClear As System.Windows.Forms.Button
  Friend WithEvents splt As System.Windows.Forms.SplitContainer
  Friend WithEvents dgr As CDBNETCL.DisplayGrid
  Friend WithEvents epl As CDBNETCL.EditPanel
  Friend WithEvents lblContents As System.Windows.Forms.Label
End Class
