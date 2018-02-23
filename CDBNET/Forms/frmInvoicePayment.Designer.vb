<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmInvoicePayment
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmInvoicePayment))
    Me.bpl = New CDBNETCL.ButtonPanel()
    Me.cmdOK = New System.Windows.Forms.Button()
    Me.cmdCancel = New System.Windows.Forms.Button()
    Me.eplTop = New CDBNETCL.EditPanel()
    Me.eplBottom = New CDBNETCL.EditPanel()
    Me.dgrCashInvoices = New CDBNETCL.DisplayGrid()
    Me.bpl.SuspendLayout()
    Me.SuspendLayout()
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdOK)
    Me.bpl.Controls.Add(Me.cmdCancel)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bpl.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bpl.Location = New System.Drawing.Point(0, 199)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(509, 39)
    Me.bpl.TabIndex = 1
    '
    'cmdOK
    '
    Me.cmdOK.DialogResult = System.Windows.Forms.DialogResult.OK
    Me.cmdOK.Location = New System.Drawing.Point(151, 6)
    Me.cmdOK.Name = "cmdOK"
    Me.cmdOK.Size = New System.Drawing.Size(96, 27)
    Me.cmdOK.TabIndex = 0
    Me.cmdOK.Text = "OK"
    Me.cmdOK.UseVisualStyleBackColor = True
    '
    'cmdCancel
    '
    Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdCancel.Location = New System.Drawing.Point(262, 6)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(96, 27)
    Me.cmdCancel.TabIndex = 1
    Me.cmdCancel.Text = "Cancel"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'eplTop
    '
    Me.eplTop.AddressChanged = False
    Me.eplTop.BackColor = System.Drawing.Color.Transparent
    Me.eplTop.DataChanged = False
    Me.eplTop.Dock = System.Windows.Forms.DockStyle.Top
    Me.eplTop.Location = New System.Drawing.Point(0, 0)
    Me.eplTop.Name = "eplTop"
    Me.eplTop.Recipients = Nothing
    Me.eplTop.Size = New System.Drawing.Size(509, 37)
    Me.eplTop.SuppressDrawing = False
    Me.eplTop.TabIndex = 1
    Me.eplTop.TabSelectedIndex = 0
    Me.eplTop.TabStop = False
    '
    'eplBottom
    '
    Me.eplBottom.AddressChanged = False
    Me.eplBottom.BackColor = System.Drawing.Color.Transparent
    Me.eplBottom.DataChanged = False
    Me.eplBottom.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.eplBottom.Location = New System.Drawing.Point(0, 140)
    Me.eplBottom.Name = "eplBottom"
    Me.eplBottom.Recipients = Nothing
    Me.eplBottom.Size = New System.Drawing.Size(509, 59)
    Me.eplBottom.SuppressDrawing = False
    Me.eplBottom.TabIndex = 2
    Me.eplBottom.TabSelectedIndex = 0
    Me.eplBottom.TabStop = False
    '
    'dgrCashInvoices
    '
    Me.dgrCashInvoices.AccessibleName = "Display Grid"
    Me.dgrCashInvoices.ActiveColumn = 0
    Me.dgrCashInvoices.AllowSorting = True
    Me.dgrCashInvoices.AutoSetHeight = False
    Me.dgrCashInvoices.AutoSetRowHeight = False
    Me.dgrCashInvoices.DisplayTitle = Nothing
    Me.dgrCashInvoices.Dock = System.Windows.Forms.DockStyle.Fill
    Me.dgrCashInvoices.Location = New System.Drawing.Point(0, 37)
    Me.dgrCashInvoices.MaintenanceDesc = Nothing
    Me.dgrCashInvoices.MaxGridRows = 8
    Me.dgrCashInvoices.MultipleSelect = True
    Me.dgrCashInvoices.Name = "dgrCashInvoices"
    Me.dgrCashInvoices.RowCount = 10
    Me.dgrCashInvoices.ShowIfEmpty = False
    Me.dgrCashInvoices.Size = New System.Drawing.Size(509, 103)
    Me.dgrCashInvoices.SuppressHyperLinkFormat = False
    Me.dgrCashInvoices.TabIndex = 0
    '
    'frmInvoicePayment
    '
    Me.AcceptButton = Me.cmdOK
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
    Me.CancelButton = Me.cmdCancel
    Me.ClientSize = New System.Drawing.Size(509, 238)
    Me.Controls.Add(Me.dgrCashInvoices)
    Me.Controls.Add(Me.eplBottom)
    Me.Controls.Add(Me.eplTop)
    Me.Controls.Add(Me.bpl)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmInvoicePayment"
    Me.Text = "Invoice Payment"
    Me.bpl.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents bpl As CDBNETCL.ButtonPanel
  Friend WithEvents eplTop As CDBNETCL.EditPanel
  Friend WithEvents cmdOK As System.Windows.Forms.Button
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents eplBottom As CDBNETCL.EditPanel
  Friend WithEvents dgrCashInvoices As CDBNETCL.DisplayGrid
End Class
