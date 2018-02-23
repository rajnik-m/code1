<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCustomise
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCustomise))
    Me.lvw = New System.Windows.Forms.ListView
    Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
    Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
    Me.bpl = New CDBNETCL.ButtonPanel
    Me.cmdReset = New System.Windows.Forms.Button
    Me.cmdClose = New System.Windows.Forms.Button
    Me.pnl = New System.Windows.Forms.Panel
    Me.txtLabel = New System.Windows.Forms.TextBox
    Me.txtToolTip = New System.Windows.Forms.TextBox
    Me.lblLabel = New System.Windows.Forms.Label
    Me.lblToolTip = New System.Windows.Forms.Label
    Me.chkLabelsBelow = New System.Windows.Forms.CheckBox
    Me.bpl.SuspendLayout()
    Me.pnl.SuspendLayout()
    Me.SuspendLayout()
    '
    'lvw
    '
    Me.lvw.AllowDrop = True
    Me.lvw.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2})
    Me.lvw.Dock = System.Windows.Forms.DockStyle.Fill
    Me.lvw.LabelEdit = True
    Me.lvw.Location = New System.Drawing.Point(0, 0)
    Me.lvw.Name = "lvw"
    Me.lvw.Size = New System.Drawing.Size(403, 313)
    Me.lvw.TabIndex = 1
    Me.lvw.UseCompatibleStateImageBehavior = False
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdReset)
    Me.bpl.Controls.Add(Me.cmdClose)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bpl.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bpl.Location = New System.Drawing.Point(0, 410)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(403, 39)
    Me.bpl.TabIndex = 0
    '
    'cmdReset
    '
    Me.cmdReset.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdReset.Location = New System.Drawing.Point(100, 6)
    Me.cmdReset.Name = "cmdReset"
    Me.cmdReset.Size = New System.Drawing.Size(94, 27)
    Me.cmdReset.TabIndex = 9
    Me.cmdReset.Text = "&Reset"
    '
    'cmdClose
    '
    Me.cmdClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdClose.Location = New System.Drawing.Point(209, 6)
    Me.cmdClose.Name = "cmdClose"
    Me.cmdClose.Size = New System.Drawing.Size(94, 27)
    Me.cmdClose.TabIndex = 8
    Me.cmdClose.Text = "Close"
    '
    'pnl
    '
    Me.pnl.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
    Me.pnl.Controls.Add(Me.chkLabelsBelow)
    Me.pnl.Controls.Add(Me.txtLabel)
    Me.pnl.Controls.Add(Me.txtToolTip)
    Me.pnl.Controls.Add(Me.lblLabel)
    Me.pnl.Controls.Add(Me.lblToolTip)
    Me.pnl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.pnl.Location = New System.Drawing.Point(0, 313)
    Me.pnl.Name = "pnl"
    Me.pnl.Size = New System.Drawing.Size(403, 97)
    Me.pnl.TabIndex = 2
    '
    'txtLabel
    '
    Me.txtLabel.Location = New System.Drawing.Point(100, 34)
    Me.txtLabel.Name = "txtLabel"
    Me.txtLabel.Size = New System.Drawing.Size(291, 22)
    Me.txtLabel.TabIndex = 3
    '
    'txtToolTip
    '
    Me.txtToolTip.Location = New System.Drawing.Point(99, 7)
    Me.txtToolTip.Name = "txtToolTip"
    Me.txtToolTip.Size = New System.Drawing.Size(291, 22)
    Me.txtToolTip.TabIndex = 2
    '
    'lblLabel
    '
    Me.lblLabel.AutoSize = True
    Me.lblLabel.Location = New System.Drawing.Point(11, 37)
    Me.lblLabel.Name = "lblLabel"
    Me.lblLabel.Size = New System.Drawing.Size(47, 17)
    Me.lblLabel.TabIndex = 1
    Me.lblLabel.Text = "Label:"
    '
    'lblToolTip
    '
    Me.lblToolTip.AutoSize = True
    Me.lblToolTip.Location = New System.Drawing.Point(12, 10)
    Me.lblToolTip.Name = "lblToolTip"
    Me.lblToolTip.Size = New System.Drawing.Size(60, 17)
    Me.lblToolTip.TabIndex = 0
    Me.lblToolTip.Text = "ToolTip:"
    '
    'chkLabelsBelow
    '
    Me.chkLabelsBelow.AutoSize = True
    Me.chkLabelsBelow.Location = New System.Drawing.Point(15, 64)
    Me.chkLabelsBelow.Name = "chkLabelsBelow"
    Me.chkLabelsBelow.Size = New System.Drawing.Size(175, 21)
    Me.chkLabelsBelow.TabIndex = 4
    Me.chkLabelsBelow.Text = "Show labels below icon"
    Me.chkLabelsBelow.UseVisualStyleBackColor = True
    '
    'frmCustomise
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.cmdClose
    Me.ClientSize = New System.Drawing.Size(403, 449)
    Me.Controls.Add(Me.lvw)
    Me.Controls.Add(Me.pnl)
    Me.Controls.Add(Me.bpl)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Margin = New System.Windows.Forms.Padding(4)
    Me.Name = "frmCustomise"
    Me.ShowInTaskbar = False
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.Text = "Drag Commands to the Toolbar"
    Me.bpl.ResumeLayout(False)
    Me.pnl.ResumeLayout(False)
    Me.pnl.PerformLayout()
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents bpl As CDBNETCL.ButtonPanel
  Friend WithEvents cmdClose As System.Windows.Forms.Button
  Friend WithEvents lvw As System.Windows.Forms.ListView
  Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
  Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
  Friend WithEvents cmdReset As System.Windows.Forms.Button
  Friend WithEvents pnl As System.Windows.Forms.Panel
  Friend WithEvents txtLabel As System.Windows.Forms.TextBox
  Friend WithEvents txtToolTip As System.Windows.Forms.TextBox
  Friend WithEvents lblLabel As System.Windows.Forms.Label
  Friend WithEvents lblToolTip As System.Windows.Forms.Label
  Friend WithEvents chkLabelsBelow As System.Windows.Forms.CheckBox

End Class
