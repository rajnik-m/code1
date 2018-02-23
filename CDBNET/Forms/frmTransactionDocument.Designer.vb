<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTransactionDocument
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmTransactionDocument))
    Me.cmdOK = New System.Windows.Forms.Button
    Me.dgr = New CDBNETCL.DisplayGrid
    Me.bpl = New CDBNETCL.ButtonPanel
    Me.cmdEdit = New System.Windows.Forms.Button
    Me.epl = New CDBNETCL.EditPanel
    Me.bpl.SuspendLayout()
    Me.SuspendLayout()
    '
    'cmdOK
    '
    Me.cmdOK.Location = New System.Drawing.Point(180, 6)
    Me.cmdOK.Name = "cmdOK"
    Me.cmdOK.Size = New System.Drawing.Size(94, 27)
    Me.cmdOK.TabIndex = 1
    Me.cmdOK.Text = "OK"
    Me.cmdOK.UseVisualStyleBackColor = True
    '
    'dgr
    '
    Me.dgr.AllowSorting = True
    Me.dgr.AutoSetHeight = False
    Me.dgr.AutoSetRowHeight = False
    Me.dgr.DisplayTitle = Nothing
    Me.dgr.Dock = System.Windows.Forms.DockStyle.Fill
    Me.dgr.Location = New System.Drawing.Point(0, 78)
    Me.dgr.MaxGridRows = 8
    Me.dgr.MultipleSelect = True
    Me.dgr.Name = "dgr"
    Me.dgr.ShowIfEmpty = False
    Me.dgr.Size = New System.Drawing.Size(564, 231)
    Me.dgr.TabIndex = 2
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdOK)
    Me.bpl.Controls.Add(Me.cmdEdit)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bpl.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bpl.Location = New System.Drawing.Point(0, 309)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(564, 39)
    Me.bpl.TabIndex = 1
    '
    'cmdEdit
    '
    Me.cmdEdit.Location = New System.Drawing.Point(289, 6)
    Me.cmdEdit.Name = "cmdEdit"
    Me.cmdEdit.Size = New System.Drawing.Size(94, 27)
    Me.cmdEdit.TabIndex = 0
    Me.cmdEdit.Text = "&Edit"
    Me.cmdEdit.UseVisualStyleBackColor = True
    '
    'epl
    '
    Me.epl.AddressChanged = False
    Me.epl.BackColor = System.Drawing.Color.Transparent
    Me.epl.DataChanged = False
    Me.epl.Dock = System.Windows.Forms.DockStyle.Top
    Me.epl.Location = New System.Drawing.Point(0, 0)
    Me.epl.Name = "epl"
    Me.epl.Recipients = Nothing
    Me.epl.Size = New System.Drawing.Size(564, 78)
    Me.epl.SuppressDrawing = False
    Me.epl.TabIndex = 0
    Me.epl.TabSelectedIndex = 0
    '
    'frmTransactionDocument
    '
    Me.AcceptButton = Me.cmdOK
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(564, 348)
    Me.Controls.Add(Me.dgr)
    Me.Controls.Add(Me.bpl)
    Me.Controls.Add(Me.epl)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmTransactionDocument"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.bpl.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents epl As CDBNETCL.EditPanel
  Friend WithEvents bpl As CDBNETCL.ButtonPanel
  Friend WithEvents cmdOK As System.Windows.Forms.Button
  Friend WithEvents cmdEdit As System.Windows.Forms.Button
  Friend WithEvents dgr As CDBNETCL.DisplayGrid
End Class
