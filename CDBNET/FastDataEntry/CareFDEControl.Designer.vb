<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CareFDEControl
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
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
    Me.pnl = New System.Windows.Forms.Panel
    Me.epl = New CDBNETCL.EditPanel
    Me.pnl.SuspendLayout()
    Me.SuspendLayout()
    '
    'pnl
    '
    Me.pnl.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.pnl.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
    Me.pnl.Controls.Add(Me.epl)
    Me.pnl.Location = New System.Drawing.Point(0, 0)
    Me.pnl.Margin = New System.Windows.Forms.Padding(4, 0, 4, 4)
    Me.pnl.Name = "pnl"
    Me.pnl.Size = New System.Drawing.Size(746, 61)
    Me.pnl.TabIndex = 0
    '
    'epl
    '
    Me.epl.AddressChanged = False
    Me.epl.BackColor = System.Drawing.Color.Transparent
    Me.epl.DataChanged = False
    Me.epl.Dock = System.Windows.Forms.DockStyle.Fill
    Me.epl.Location = New System.Drawing.Point(0, 0)
    Me.epl.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
    Me.epl.Name = "epl"
    Me.epl.Recipients = Nothing
    Me.epl.Size = New System.Drawing.Size(744, 59)
    Me.epl.SuppressDrawing = False
    Me.epl.TabIndex = 0
    Me.epl.TabSelectedIndex = 0
    '
    'CareFDEControl
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.Controls.Add(Me.pnl)
    Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
    Me.Name = "CareFDEControl"
    Me.Padding = New System.Windows.Forms.Padding(0, 4, 0, 4)
    Me.Size = New System.Drawing.Size(747, 68)
    Me.pnl.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents pnl As System.Windows.Forms.Panel
  Friend WithEvents epl As CDBNETCL.EditPanel

End Class
