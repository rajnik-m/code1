<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmWaitingList
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmWaitingList))
    Me.pnl = New CDBNETCL.PanelEx()
    Me.spl = New System.Windows.Forms.SplitContainer()
    Me.dgrBookings = New CDBNETCL.DisplayGrid()
    Me.lblBookings = New System.Windows.Forms.Label()
    Me.dgrDelegates = New CDBNETCL.DisplayGrid()
    Me.lblDelegates = New System.Windows.Forms.Label()
    Me.bpl = New CDBNETCL.ButtonPanel()
    Me.cmdOK = New System.Windows.Forms.Button()
    Me.cmdApply = New System.Windows.Forms.Button()
    Me.cmdCancel = New System.Windows.Forms.Button()
    Me.pnl.SuspendLayout()
    CType(Me.spl, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.spl.Panel1.SuspendLayout()
    Me.spl.Panel2.SuspendLayout()
    Me.spl.SuspendLayout()
    Me.bpl.SuspendLayout()
    Me.SuspendLayout()
    '
    'pnl
    '
    Me.pnl.BackColor = System.Drawing.Color.Transparent
    Me.pnl.Controls.Add(Me.spl)
    Me.pnl.Dock = System.Windows.Forms.DockStyle.Fill
    Me.pnl.Location = New System.Drawing.Point(0, 0)
    Me.pnl.Name = "pnl"
    Me.pnl.Size = New System.Drawing.Size(656, 296)
    Me.pnl.TabIndex = 0
    '
    'spl
    '
    Me.spl.Dock = System.Windows.Forms.DockStyle.Fill
    Me.spl.Location = New System.Drawing.Point(0, 0)
    Me.spl.Name = "spl"
    '
    'spl.Panel1
    '
    Me.spl.Panel1.Controls.Add(Me.dgrBookings)
    Me.spl.Panel1.Controls.Add(Me.lblBookings)
    '
    'spl.Panel2
    '
    Me.spl.Panel2.Controls.Add(Me.dgrDelegates)
    Me.spl.Panel2.Controls.Add(Me.lblDelegates)
    Me.spl.Size = New System.Drawing.Size(656, 296)
    Me.spl.SplitterDistance = 308
    Me.spl.SplitterWidth = 8
    Me.spl.TabIndex = 0
    '
    'dgrBookings
    '
    Me.dgrBookings.AccessibleName = "Display Grid"
    Me.dgrBookings.ActiveColumn = 0
    Me.dgrBookings.AllowSorting = True
    Me.dgrBookings.AutoSetHeight = False
    Me.dgrBookings.AutoSetRowHeight = False
    Me.dgrBookings.DisplayTitle = Nothing
    Me.dgrBookings.Dock = System.Windows.Forms.DockStyle.Fill
    Me.dgrBookings.Location = New System.Drawing.Point(0, 29)
    Me.dgrBookings.MaintenanceDesc = Nothing
    Me.dgrBookings.MaxGridRows = 8
    Me.dgrBookings.MultipleSelect = True
    Me.dgrBookings.Name = "dgrBookings"
    Me.dgrBookings.Padding = New System.Windows.Forms.Padding(4)
    Me.dgrBookings.RowCount = 10
    Me.dgrBookings.ShowIfEmpty = False
    Me.dgrBookings.Size = New System.Drawing.Size(308, 267)
    Me.dgrBookings.SuppressHyperLinkFormat = False
    Me.dgrBookings.TabIndex = 1
    '
    'lblBookings
    '
    Me.lblBookings.AutoSize = True
    Me.lblBookings.Dock = System.Windows.Forms.DockStyle.Top
    Me.lblBookings.Location = New System.Drawing.Point(0, 0)
    Me.lblBookings.Name = "lblBookings"
    Me.lblBookings.Padding = New System.Windows.Forms.Padding(8, 8, 0, 8)
    Me.lblBookings.Size = New System.Drawing.Size(59, 29)
    Me.lblBookings.TabIndex = 0
    Me.lblBookings.Text = "Bookings"
    '
    'dgrDelegates
    '
    Me.dgrDelegates.AccessibleName = "Display Grid"
    Me.dgrDelegates.ActiveColumn = 0
    Me.dgrDelegates.AllowSorting = True
    Me.dgrDelegates.AutoSetHeight = False
    Me.dgrDelegates.AutoSetRowHeight = False
    Me.dgrDelegates.DisplayTitle = Nothing
    Me.dgrDelegates.Dock = System.Windows.Forms.DockStyle.Fill
    Me.dgrDelegates.Location = New System.Drawing.Point(0, 29)
    Me.dgrDelegates.MaintenanceDesc = Nothing
    Me.dgrDelegates.MaxGridRows = 8
    Me.dgrDelegates.MultipleSelect = True
    Me.dgrDelegates.Name = "dgrDelegates"
    Me.dgrDelegates.Padding = New System.Windows.Forms.Padding(4)
    Me.dgrDelegates.RowCount = 10
    Me.dgrDelegates.ShowIfEmpty = False
    Me.dgrDelegates.Size = New System.Drawing.Size(340, 267)
    Me.dgrDelegates.SuppressHyperLinkFormat = False
    Me.dgrDelegates.TabIndex = 2
    '
    'lblDelegates
    '
    Me.lblDelegates.AutoSize = True
    Me.lblDelegates.Dock = System.Windows.Forms.DockStyle.Top
    Me.lblDelegates.Location = New System.Drawing.Point(0, 0)
    Me.lblDelegates.Name = "lblDelegates"
    Me.lblDelegates.Padding = New System.Windows.Forms.Padding(8, 8, 0, 8)
    Me.lblDelegates.Size = New System.Drawing.Size(63, 29)
    Me.lblDelegates.TabIndex = 1
    Me.lblDelegates.Text = "Delegates"
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdOK)
    Me.bpl.Controls.Add(Me.cmdApply)
    Me.bpl.Controls.Add(Me.cmdCancel)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bpl.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bpl.Location = New System.Drawing.Point(0, 296)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(656, 39)
    Me.bpl.TabIndex = 1
    '
    'cmdOK
    '
    Me.cmdOK.Location = New System.Drawing.Point(169, 6)
    Me.cmdOK.Name = "cmdOK"
    Me.cmdOK.Size = New System.Drawing.Size(96, 27)
    Me.cmdOK.TabIndex = 2
    Me.cmdOK.Text = "OK"
    Me.cmdOK.UseVisualStyleBackColor = True
    '
    'cmdApply
    '
    Me.cmdApply.Location = New System.Drawing.Point(280, 6)
    Me.cmdApply.Name = "cmdApply"
    Me.cmdApply.Size = New System.Drawing.Size(96, 27)
    Me.cmdApply.TabIndex = 1
    Me.cmdApply.Text = "Apply"
    Me.cmdApply.UseVisualStyleBackColor = True
    '
    'cmdCancel
    '
    Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdCancel.Location = New System.Drawing.Point(391, 6)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(96, 27)
    Me.cmdCancel.TabIndex = 0
    Me.cmdCancel.Text = "Cancel"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'frmWaitingList
    '
    Me.AcceptButton = Me.cmdOK
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.cmdCancel
    Me.ClientSize = New System.Drawing.Size(656, 335)
    Me.Controls.Add(Me.pnl)
    Me.Controls.Add(Me.bpl)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmWaitingList"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.pnl.ResumeLayout(False)
    Me.spl.Panel1.ResumeLayout(False)
    Me.spl.Panel1.PerformLayout()
    Me.spl.Panel2.ResumeLayout(False)
    Me.spl.Panel2.PerformLayout()
    CType(Me.spl, System.ComponentModel.ISupportInitialize).EndInit()
    Me.spl.ResumeLayout(False)
    Me.bpl.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents pnl As CDBNETCL.PanelEx
  Friend WithEvents spl As System.Windows.Forms.SplitContainer
  Friend WithEvents bpl As CDBNETCL.ButtonPanel
  Friend WithEvents dgrBookings As CDBNETCL.DisplayGrid
  Friend WithEvents lblBookings As System.Windows.Forms.Label
  Friend WithEvents lblDelegates As System.Windows.Forms.Label
  Friend WithEvents cmdOK As System.Windows.Forms.Button
  Friend WithEvents cmdApply As System.Windows.Forms.Button
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents dgrDelegates As CDBNETCL.DisplayGrid
End Class
