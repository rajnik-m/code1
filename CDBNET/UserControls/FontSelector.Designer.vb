<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FontSelector
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
    Me.txtFont = New System.Windows.Forms.TextBox
    Me.cmdFont = New System.Windows.Forms.Button
    Me.fnt = New System.Windows.Forms.FontDialog
    Me.chk = New System.Windows.Forms.CheckBox
    Me.SuspendLayout()
    '
    'txtFont
    '
    Me.txtFont.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.txtFont.Location = New System.Drawing.Point(80, 5)
    Me.txtFont.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
    Me.txtFont.Name = "txtFont"
    Me.txtFont.ReadOnly = True
    Me.txtFont.Size = New System.Drawing.Size(187, 22)
    Me.txtFont.TabIndex = 1
    '
    'cmdFont
    '
    Me.cmdFont.Location = New System.Drawing.Point(3, 5)
    Me.cmdFont.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
    Me.cmdFont.Name = "cmdFont"
    Me.cmdFont.Size = New System.Drawing.Size(71, 22)
    Me.cmdFont.TabIndex = 0
    Me.cmdFont.Text = "Font..."
    Me.cmdFont.UseVisualStyleBackColor = True
    '
    'chk
    '
    Me.chk.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.chk.AutoSize = True
    Me.chk.Location = New System.Drawing.Point(272, 9)
    Me.chk.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
    Me.chk.Name = "chk"
    Me.chk.Size = New System.Drawing.Size(18, 17)
    Me.chk.TabIndex = 2
    Me.chk.UseVisualStyleBackColor = True
    '
    'FontSelector
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.Controls.Add(Me.chk)
    Me.Controls.Add(Me.cmdFont)
    Me.Controls.Add(Me.txtFont)
    Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
    Me.Name = "FontSelector"
    Me.Size = New System.Drawing.Size(295, 34)
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents txtFont As System.Windows.Forms.TextBox
  Friend WithEvents cmdFont As System.Windows.Forms.Button
  Friend WithEvents fnt As System.Windows.Forms.FontDialog
  Friend WithEvents chk As System.Windows.Forms.CheckBox

End Class
