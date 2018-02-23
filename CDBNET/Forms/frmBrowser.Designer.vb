<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmBrowser
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBrowser))
    Me.sts = New System.Windows.Forms.StatusStrip
    Me.tsl = New System.Windows.Forms.ToolStripStatusLabel
    Me.web = New CDBNETCL.CareWebBrowser
    Me.sts.SuspendLayout()
    Me.SuspendLayout()
    '
    'sts
    '
    Me.sts.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsl})
    Me.sts.Location = New System.Drawing.Point(0, 538)
    Me.sts.Name = "sts"
    Me.sts.Size = New System.Drawing.Size(792, 22)
    Me.sts.TabIndex = 2
    '
    'tsl
    '
    Me.tsl.Name = "tsl"
    Me.tsl.Size = New System.Drawing.Size(0, 17)
    '
    'web
    '
    Me.web.Dock = System.Windows.Forms.DockStyle.Fill
    Me.web.Location = New System.Drawing.Point(0, 0)
    Me.web.Name = "web"
    Me.web.Size = New System.Drawing.Size(792, 538)
    Me.web.TabIndex = 3
    '
    'frmBrowser
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(792, 560)
    Me.Controls.Add(Me.web)
    Me.Controls.Add(Me.sts)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmBrowser"
    Me.sts.ResumeLayout(False)
    Me.sts.PerformLayout()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents sts As System.Windows.Forms.StatusStrip
  Friend WithEvents tsl As System.Windows.Forms.ToolStripStatusLabel
  Friend WithEvents web As CDBNETCL.CareWebBrowser
End Class
