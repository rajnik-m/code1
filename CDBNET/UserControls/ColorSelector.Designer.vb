<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ColorSelector
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
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
    Me.cmd = New System.Windows.Forms.Button
    Me.pnl = New System.Windows.Forms.Panel
    Me.lbl = New System.Windows.Forms.Label
    Me.SuspendLayout()
    '
    'cmd
    '
    Me.cmd.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmd.Location = New System.Drawing.Point(295, 4)
    Me.cmd.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
    Me.cmd.Name = "cmd"
    Me.cmd.Size = New System.Drawing.Size(21, 26)
    Me.cmd.TabIndex = 2
    Me.cmd.Text = "?"
    Me.cmd.UseVisualStyleBackColor = True
    '
    'pnl
    '
    Me.pnl.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.pnl.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
    Me.pnl.Location = New System.Drawing.Point(219, 2)
    Me.pnl.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
    Me.pnl.Name = "pnl"
    Me.pnl.Size = New System.Drawing.Size(67, 27)
    Me.pnl.TabIndex = 1
    '
    'lbl
    '
    Me.lbl.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.lbl.Location = New System.Drawing.Point(3, 7)
    Me.lbl.Name = "lbl"
    Me.lbl.Size = New System.Drawing.Size(207, 21)
    Me.lbl.TabIndex = 0
    '
    'ColorSelector
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.Controls.Add(Me.cmd)
    Me.Controls.Add(Me.pnl)
    Me.Controls.Add(Me.lbl)
    Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
    Me.Name = "ColorSelector"
    Me.Size = New System.Drawing.Size(317, 34)
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents cmd As System.Windows.Forms.Button
  Friend WithEvents pnl As System.Windows.Forms.Panel
  Friend WithEvents lbl As System.Windows.Forms.Label
End Class
