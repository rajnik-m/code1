<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmHTMLEditor
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmHTMLEditor))
    Me.pnl = New System.Windows.Forms.Panel
    Me.bpl = New CDBNETCL.ButtonPanel
    Me.cmdOK = New System.Windows.Forms.Button
    Me.cmdEdit = New System.Windows.Forms.Button
    Me.cmdCancel = New System.Windows.Forms.Button
    Me.bpl.SuspendLayout()
    Me.SuspendLayout()
    '
    'pnl
    '
    Me.pnl.Dock = System.Windows.Forms.DockStyle.Fill
    Me.pnl.Location = New System.Drawing.Point(0, 0)
    Me.pnl.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
    Me.pnl.Name = "pnl"
    Me.pnl.Size = New System.Drawing.Size(1053, 662)
    Me.pnl.TabIndex = 0
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdOK)
    Me.bpl.Controls.Add(Me.cmdEdit)
    Me.bpl.Controls.Add(Me.cmdCancel)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bpl.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bpl.Location = New System.Drawing.Point(0, 662)
    Me.bpl.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(1053, 39)
    Me.bpl.TabIndex = 1
    '
    'cmdOK
    '
    Me.cmdOK.DialogResult = System.Windows.Forms.DialogResult.OK
    Me.cmdOK.Location = New System.Drawing.Point(370, 6)
    Me.cmdOK.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
    Me.cmdOK.Name = "cmdOK"
    Me.cmdOK.Size = New System.Drawing.Size(94, 27)
    Me.cmdOK.TabIndex = 2
    Me.cmdOK.Text = "OK"
    Me.cmdOK.UseVisualStyleBackColor = True
    '
    'cmdEdit
    '
    Me.cmdEdit.Location = New System.Drawing.Point(479, 6)
    Me.cmdEdit.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
    Me.cmdEdit.Name = "cmdEdit"
    Me.cmdEdit.Size = New System.Drawing.Size(94, 27)
    Me.cmdEdit.TabIndex = 1
    Me.cmdEdit.Text = "&Edit HTML"
    Me.cmdEdit.UseVisualStyleBackColor = True
    '
    'cmdCancel
    '
    Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdCancel.Location = New System.Drawing.Point(588, 6)
    Me.cmdCancel.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(94, 27)
    Me.cmdCancel.TabIndex = 0
    Me.cmdCancel.Text = "Cancel"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'frmHTMLEditor
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(1053, 701)
    Me.Controls.Add(Me.pnl)
    Me.Controls.Add(Me.bpl)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
    Me.Name = "frmHTMLEditor"
    Me.ShowInTaskbar = False
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    Me.Text = "HTML Editor"
    Me.bpl.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents pnl As System.Windows.Forms.Panel
  Friend WithEvents bpl As CDBNETCL.ButtonPanel
  Friend WithEvents cmdEdit As System.Windows.Forms.Button
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents cmdOK As System.Windows.Forms.Button
End Class
