﻿Imports CDBNETXAMLPages

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmModernMenuViewer
  Inherits System.Windows.Forms.Form

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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmModernMenuViewer))
    Me.ElementHost1 = New System.Windows.Forms.Integration.ElementHost()
    Me.SuspendLayout()
    HomeScreen1 = New HomeScreen()
    '
    'ElementHost1
    '
    Me.ElementHost1.Dock = System.Windows.Forms.DockStyle.Fill
    Me.ElementHost1.Location = New System.Drawing.Point(0, 0)
    Me.ElementHost1.Margin = New System.Windows.Forms.Padding(4)
    Me.ElementHost1.Name = "ElementHost1"
    Me.ElementHost1.Size = New System.Drawing.Size(840, 549)
    Me.ElementHost1.TabIndex = 0
    Me.ElementHost1.Text = "ElementHost1"
    Me.ElementHost1.Child = HomeScreen1
    '
    'frmModernMenuViewer
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(840, 549)
    Me.Controls.Add(Me.ElementHost1)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Margin = New System.Windows.Forms.Padding(4)
    Me.Name = "frmModernMenuViewer"
    Me.Text = "HOME     "
    Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents ElementHost1 As System.Windows.Forms.Integration.ElementHost
  Friend HomeScreen1 As HomeScreen
End Class
