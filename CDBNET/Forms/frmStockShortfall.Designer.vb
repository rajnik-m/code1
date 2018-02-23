<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmStockShortfall
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
    Me.components = New System.ComponentModel.Container
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmStockShortfall))
    Me.dgr = New CDBNETCL.DisplayGrid
    Me.lblPickingList = New System.Windows.Forms.Label
    Me.bpl = New CDBNETCL.ButtonPanel
    Me.cmdOK = New System.Windows.Forms.Button
    Me.cmdCancel = New System.Windows.Forms.Button
    Me.bpl.SuspendLayout()
    Me.SuspendLayout()
    '
    'dgr
    '
    Me.dgr.AccessibleDescription = "Display List"
    Me.dgr.AccessibleName = "Display List"
    Me.dgr.AccessibleRole = System.Windows.Forms.AccessibleRole.Table
    Me.dgr.AllowSorting = True
    Me.dgr.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.dgr.AutoSetHeight = False
    Me.dgr.AutoSetRowHeight = False
    Me.dgr.DisplayTitle = Nothing
    Me.dgr.Location = New System.Drawing.Point(15, 25)
    Me.dgr.MaxGridRows = 8
    Me.dgr.MultipleSelect = False
    Me.dgr.Name = "dgr"
    Me.dgr.RowCount = 10
    Me.dgr.ShowIfEmpty = False
    Me.dgr.Size = New System.Drawing.Size(374, 225)
    Me.dgr.TabIndex = 0
    Me.dgr.TabStop = False
    '
    'lblPickingList
    '
    Me.lblPickingList.AutoSize = True
    Me.lblPickingList.Location = New System.Drawing.Point(12, 9)
    Me.lblPickingList.Name = "lblPickingList"
    Me.lblPickingList.Size = New System.Drawing.Size(104, 13)
    Me.lblPickingList.TabIndex = 3
    Me.lblPickingList.Text = "Picking List Number:"
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdOK)
    Me.bpl.Controls.Add(Me.cmdCancel)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bpl.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bpl.Location = New System.Drawing.Point(0, 252)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(405, 39)
    Me.bpl.TabIndex = 4
    '
    'cmdOK
    '
    Me.cmdOK.Location = New System.Drawing.Point(99, 6)
    Me.cmdOK.Name = "cmdOK"
    Me.cmdOK.Size = New System.Drawing.Size(96, 27)
    Me.cmdOK.TabIndex = 0
    Me.cmdOK.Text = "&OK"
    Me.cmdOK.UseVisualStyleBackColor = True
    '
    'cmdCancel
    '
    Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdCancel.Location = New System.Drawing.Point(210, 6)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(96, 27)
    Me.cmdCancel.TabIndex = 1
    Me.cmdCancel.Text = "&Cancel"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'frmStockShortfall
    '
    Me.AcceptButton = Me.cmdOK
    Me.CancelButton = Me.cmdCancel
    Me.ClientSize = New System.Drawing.Size(405, 291)
    Me.Controls.Add(Me.bpl)
    Me.Controls.Add(Me.lblPickingList)
    Me.Controls.Add(Me.dgr)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmStockShortfall"
    Me.Text = "Stock Allocation"
    Me.bpl.ResumeLayout(False)
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents dgr As CDBNETCL.DisplayGrid
  Friend WithEvents lblPickingList As System.Windows.Forms.Label
  Friend WithEvents bpl As CDBNETCL.ButtonPanel
  Friend WithEvents cmdOK As System.Windows.Forms.Button
  Friend WithEvents cmdCancel As System.Windows.Forms.Button

End Class
