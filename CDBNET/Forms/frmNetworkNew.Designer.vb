<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmNetworkNew
  Inherits CDBNETCL.MaintenanceParentForm

  'Form overrides dispose to clean up the component list.
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(NetworkTreeView))
    Me.pnlTreeView = New CDBNETCL.PanelEx
    'Me.tvw = New CDBNETCL.VistaTreeView
    Me.pnlTreeView.SuspendLayout()
    '
    'pnlTreeView
    '
    Me.pnlTreeView.BackColor = System.Drawing.Color.Transparent
    'Me.pnlTreeView.Controls.Add(Me.tvw)
    Me.pnlTreeView.Dock = System.Windows.Forms.DockStyle.Fill
    Me.pnlTreeView.Location = New System.Drawing.Point(0, 0)
    Me.pnlTreeView.Name = "pnlTreeView"
    Me.pnlTreeView.Size = New System.Drawing.Size(387, 400)
    Me.pnlTreeView.TabIndex = 2
    '
    'frmNetworkTreeView
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(387, 400)
    Me.Controls.Add(Me.pnlTreeView)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmNetworkNew"
    Me.pnlTreeView.ResumeLayout(False)
    Me.ResumeLayout(False)
  End Sub

  Friend WithEvents pnlTreeView As CDBNETCL.PanelEx
  'Friend WithEvents tvw As VistaTreeView
End Class
