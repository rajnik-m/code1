<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmHtmlSettings
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
    Me.HtmlBrowser = New System.Windows.Forms.WebBrowser()
    Me.Panel1 = New System.Windows.Forms.Panel()
    Me.btnOk = New System.Windows.Forms.Button()
    Me.btnCancel = New System.Windows.Forms.Button()
    Me.Panel1.SuspendLayout()
    Me.SuspendLayout()
    '
    'HtmlBrowser
    '
    Me.HtmlBrowser.Dock = System.Windows.Forms.DockStyle.Fill
    Me.HtmlBrowser.Location = New System.Drawing.Point(0, 0)
    Me.HtmlBrowser.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
    Me.HtmlBrowser.MinimumSize = New System.Drawing.Size(15, 16)
    Me.HtmlBrowser.Name = "HtmlBrowser"
    Me.HtmlBrowser.Size = New System.Drawing.Size(640, 480)
    Me.HtmlBrowser.TabIndex = 0
    '
    'Panel1
    '
    Me.Panel1.Controls.Add(Me.btnOk)
    Me.Panel1.Controls.Add(Me.btnCancel)
    Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.Panel1.Location = New System.Drawing.Point(0, 480)
    Me.Panel1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
    Me.Panel1.Name = "Panel1"
    Me.Panel1.Size = New System.Drawing.Size(640, 29)
    Me.Panel1.TabIndex = 1
    '
    'btnOk
    '
    Me.btnOk.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.btnOk.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.btnOk.Location = New System.Drawing.Point(520, 5)
    Me.btnOk.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
    Me.btnOk.Name = "btnOk"
    Me.btnOk.Size = New System.Drawing.Size(56, 20)
    Me.btnOk.TabIndex = 1
    Me.btnOk.Text = "&Ok"
    Me.btnOk.UseVisualStyleBackColor = True
    '
    'btnCancel
    '
    Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.btnCancel.Location = New System.Drawing.Point(581, 6)
    Me.btnCancel.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
    Me.btnCancel.Name = "btnCancel"
    Me.btnCancel.Size = New System.Drawing.Size(56, 20)
    Me.btnCancel.TabIndex = 0
    Me.btnCancel.Text = "&Cancel"
    Me.btnCancel.UseVisualStyleBackColor = True
    '
    'frmHtmlSettings
    '
    Me.AcceptButton = Me.btnOk
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.btnCancel
    Me.ClientSize = New System.Drawing.Size(640, 509)
    Me.Controls.Add(Me.HtmlBrowser)
    Me.Controls.Add(Me.Panel1)
    Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
    Me.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
    Me.MaximizeBox = False
    Me.MinimizeBox = False
    Me.Name = "frmHtmlSettings"
    Me.ShowIcon = False
    Me.Panel1.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents HtmlBrowser As System.Windows.Forms.WebBrowser
  Friend WithEvents Panel1 As System.Windows.Forms.Panel
  Friend WithEvents btnOk As System.Windows.Forms.Button
  Friend WithEvents btnCancel As System.Windows.Forms.Button
End Class
