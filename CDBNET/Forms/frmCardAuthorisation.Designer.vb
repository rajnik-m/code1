<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCardAuthorisation
  Inherits CDBNETCL.ThemedForm

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
    Me.browser = New System.Windows.Forms.WebBrowser()
    Me.SuspendLayout()
    '
    'browser
    '
    Me.browser.Dock = System.Windows.Forms.DockStyle.Fill
    Me.browser.Location = New System.Drawing.Point(0, 0)
    Me.browser.MinimumSize = New System.Drawing.Size(20, 20)
    Me.browser.Name = "browser"
    Me.browser.Size = New System.Drawing.Size(1006, 517)
    Me.browser.TabIndex = 0
    '
    'frmCardAuthorisation
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(1006, 517)
    Me.Controls.Add(Me.browser)
    Me.Name = "frmCardAuthorisation"
    Me.Text = "frmCardAuthorisation"
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents browser As System.Windows.Forms.WebBrowser
End Class
