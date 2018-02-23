<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPagedTreeView
  Inherits CDBNETCL.PersistentForm

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
    Me.PagedTreeView = New CDBNETCLPages.PagedTreeView()
    Me.SuspendLayout()
    '
    'PagedTreeView
    '
    Me.PagedTreeView.Dock = System.Windows.Forms.DockStyle.Fill
    Me.PagedTreeView.IDataContext = Nothing
    Me.PagedTreeView.Location = New System.Drawing.Point(0, 0)
    Me.PagedTreeView.Name = "PagedTreeView"
    Me.PagedTreeView.Size = New System.Drawing.Size(733, 459)
    Me.PagedTreeView.TabIndex = 0
    '
    'frmPagedTreeView
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(733, 459)
    Me.Controls.Add(Me.PagedTreeView)
    Me.DoubleBuffered = True
    Me.Name = "frmPagedTreeView"
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents PagedTreeView As CDBNETCLPages.PagedTreeView
End Class
