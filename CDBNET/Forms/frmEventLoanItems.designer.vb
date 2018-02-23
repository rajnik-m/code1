<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmEventLoanItems
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmEventLoanItems))
    Me.splt = New System.Windows.Forms.SplitContainer()
    Me.dgr = New CDBNETCL.DisplayGrid()
    Me.bpl = New CDBNETCL.ButtonPanel()
    Me.cmdReturned = New System.Windows.Forms.Button()
    Me.cmdOK = New System.Windows.Forms.Button()
    Me.cmdCancel = New System.Windows.Forms.Button()
    Me.epl = New CDBNETCL.EditPanel()
    CType(Me.splt, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.splt.Panel1.SuspendLayout()
    Me.splt.Panel2.SuspendLayout()
    Me.splt.SuspendLayout()
    Me.bpl.SuspendLayout()
    Me.SuspendLayout()
    '
    'splt
    '
    Me.splt.Dock = System.Windows.Forms.DockStyle.Fill
    Me.splt.FixedPanel = System.Windows.Forms.FixedPanel.Panel2
    Me.splt.Location = New System.Drawing.Point(0, 0)
    Me.splt.Name = "splt"
    Me.splt.Orientation = System.Windows.Forms.Orientation.Horizontal
    '
    'splt.Panel1
    '
    Me.splt.Panel1.Controls.Add(Me.dgr)
    '
    'splt.Panel2
    '
    Me.splt.Panel2.Controls.Add(Me.bpl)
    Me.splt.Panel2.Controls.Add(Me.epl)
    Me.splt.Size = New System.Drawing.Size(691, 328)
    Me.splt.SplitterDistance = 153
    Me.splt.TabIndex = 0
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
    Me.dgr.Location = New System.Drawing.Point(0, 0)
    Me.dgr.MaintenanceDesc = Nothing
    Me.dgr.MaxGridRows = 6
    Me.dgr.MultipleSelect = False
    Me.dgr.Name = "dgr"
    Me.dgr.RowCount = 10
    Me.dgr.ShowIfEmpty = False
    Me.dgr.Size = New System.Drawing.Size(691, 153)
    Me.dgr.SuppressHyperLinkFormat = False
    Me.dgr.TabIndex = 0
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdReturned)
    Me.bpl.Controls.Add(Me.cmdOK)
    Me.bpl.Controls.Add(Me.cmdCancel)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bpl.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bpl.Location = New System.Drawing.Point(0, 132)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(691, 39)
    Me.bpl.TabIndex = 1
    '
    'cmdReturned
    '
    Me.cmdReturned.Location = New System.Drawing.Point(186, 6)
    Me.cmdReturned.Name = "cmdReturned"
    Me.cmdReturned.Size = New System.Drawing.Size(96, 27)
    Me.cmdReturned.TabIndex = 0
    Me.cmdReturned.Text = "Returned"
    Me.cmdReturned.UseVisualStyleBackColor = True
    '
    'cmdOK
    '
    Me.cmdOK.Location = New System.Drawing.Point(297, 6)
    Me.cmdOK.Name = "cmdOK"
    Me.cmdOK.Size = New System.Drawing.Size(96, 27)
    Me.cmdOK.TabIndex = 1
    Me.cmdOK.Text = "OK"
    Me.cmdOK.UseVisualStyleBackColor = True
    '
    'cmdCancel
    '
    Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdCancel.Location = New System.Drawing.Point(408, 6)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(96, 27)
    Me.cmdCancel.TabIndex = 2
    Me.cmdCancel.Text = "Cancel"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'epl
    '
    Me.epl.AddressChanged = False
    Me.epl.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.epl.BackColor = System.Drawing.Color.Transparent
    Me.epl.DataChanged = False
    Me.epl.Location = New System.Drawing.Point(3, 4)
    Me.epl.Name = "epl"
    Me.epl.Recipients = Nothing
    Me.epl.Size = New System.Drawing.Size(685, 126)
    Me.epl.SuppressDrawing = False
    Me.epl.TabIndex = 0
    Me.epl.TabSelectedIndex = 0
    '
    'frmEventLoanItems
    '
    Me.AcceptButton = Me.cmdOK
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.cmdCancel
    Me.ClientSize = New System.Drawing.Size(691, 328)
    Me.Controls.Add(Me.splt)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmEventLoanItems"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    Me.Text = "Loan Items"
    Me.splt.Panel1.ResumeLayout(False)
    Me.splt.Panel2.ResumeLayout(False)
    CType(Me.splt, System.ComponentModel.ISupportInitialize).EndInit()
    Me.splt.ResumeLayout(False)
    Me.bpl.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents splt As System.Windows.Forms.SplitContainer
  Friend WithEvents dgr As CDBNETCL.DisplayGrid
  Friend WithEvents epl As CDBNETCL.EditPanel
  Friend WithEvents bpl As CDBNETCL.ButtonPanel
  Friend WithEvents cmdOK As System.Windows.Forms.Button
  Friend WithEvents cmdReturned As System.Windows.Forms.Button
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
End Class
