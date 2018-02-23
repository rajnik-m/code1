<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmEventEligibility
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmEventEligibility))
    Me.bpl = New CDBNETCL.ButtonPanel()
    Me.cmdApprove = New System.Windows.Forms.Button()
    Me.cmdDefer = New System.Windows.Forms.Button()
    Me.cmdReject = New System.Windows.Forms.Button()
    Me.cmdView = New System.Windows.Forms.Button()
    Me.epl = New CDBNETCL.EditPanel()
    Me.bpl.SuspendLayout()
    Me.SuspendLayout()
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdApprove)
    Me.bpl.Controls.Add(Me.cmdDefer)
    Me.bpl.Controls.Add(Me.cmdReject)
    Me.bpl.Controls.Add(Me.cmdView)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bpl.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bpl.Location = New System.Drawing.Point(0, 245)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(520, 39)
    Me.bpl.TabIndex = 0
    '
    'cmdApprove
    '
    Me.cmdApprove.Location = New System.Drawing.Point(45, 6)
    Me.cmdApprove.Name = "cmdApprove"
    Me.cmdApprove.Size = New System.Drawing.Size(96, 27)
    Me.cmdApprove.TabIndex = 3
    Me.cmdApprove.Text = "&Approve"
    Me.cmdApprove.UseVisualStyleBackColor = True
    '
    'cmdDefer
    '
    Me.cmdDefer.Location = New System.Drawing.Point(156, 6)
    Me.cmdDefer.Name = "cmdDefer"
    Me.cmdDefer.Size = New System.Drawing.Size(96, 27)
    Me.cmdDefer.TabIndex = 2
    Me.cmdDefer.Text = "&Defer"
    Me.cmdDefer.UseVisualStyleBackColor = True
    '
    'cmdReject
    '
    Me.cmdReject.Location = New System.Drawing.Point(267, 6)
    Me.cmdReject.Name = "cmdReject"
    Me.cmdReject.Size = New System.Drawing.Size(96, 27)
    Me.cmdReject.TabIndex = 1
    Me.cmdReject.Text = "&Reject"
    Me.cmdReject.UseVisualStyleBackColor = True
    '
    'cmdView
    '
    Me.cmdView.Location = New System.Drawing.Point(378, 6)
    Me.cmdView.Name = "cmdView"
    Me.cmdView.Size = New System.Drawing.Size(96, 27)
    Me.cmdView.TabIndex = 0
    Me.cmdView.Text = "&View"
    Me.cmdView.UseVisualStyleBackColor = True
    '
    'epl
    '
    Me.epl.AddressChanged = False
    Me.epl.BackColor = System.Drawing.Color.Transparent
    Me.epl.DataChanged = False
    Me.epl.Dock = System.Windows.Forms.DockStyle.Fill
    Me.epl.Location = New System.Drawing.Point(0, 0)
    Me.epl.Name = "epl"
    Me.epl.Recipients = Nothing
    Me.epl.Size = New System.Drawing.Size(520, 245)
    Me.epl.SuppressDrawing = False
    Me.epl.TabIndex = 1
    Me.epl.TabSelectedIndex = 0
    '
    'frmEventEligibility
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(520, 284)
    Me.Controls.Add(Me.epl)
    Me.Controls.Add(Me.bpl)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmEventEligibility"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.bpl.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents bpl As CDBNETCL.ButtonPanel
  Friend WithEvents epl As CDBNETCL.EditPanel
  Friend WithEvents cmdApprove As System.Windows.Forms.Button
  Friend WithEvents cmdDefer As System.Windows.Forms.Button
  Friend WithEvents cmdReject As System.Windows.Forms.Button
  Friend WithEvents cmdView As System.Windows.Forms.Button
End Class
