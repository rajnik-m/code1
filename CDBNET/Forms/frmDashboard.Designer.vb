<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDashboard
  Inherits MaintenanceParentForm

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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDashboard))
    Me.dtc = New CDBNETCL.DashboardTabControl()
    Me.SuspendLayout()
    '
    'dtc
    '
    Me.dtc.BackColor = System.Drawing.Color.Transparent
    Me.dtc.Dock = System.Windows.Forms.DockStyle.Fill
    Me.dtc.Location = New System.Drawing.Point(0, 0)
    Me.dtc.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
    Me.dtc.Name = "dtc"
    Me.dtc.Size = New System.Drawing.Size(857, 533)
    Me.dtc.TabIndex = 2
    '
    'frmDashboard
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(857, 533)
    Me.Controls.Add(Me.dtc)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmDashboard"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.Text = "Dashboard"
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents dtc As CDBNETCL.DashboardTabControl
End Class
