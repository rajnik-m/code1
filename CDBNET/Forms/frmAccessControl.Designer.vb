<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAccessControl
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAccessControl))
    Me.splt = New System.Windows.Forms.SplitContainer()
    Me.tvw = New CDBNETCL.VistaTreeView()
    Me.bpl = New CDBNETCL.ButtonPanel()
    Me.cmdOK = New System.Windows.Forms.Button()
    Me.txtAccessLevelReqd = New System.Windows.Forms.TextBox()
    Me.lblAccessLevelReqd = New System.Windows.Forms.Label()
    CType(Me.splt, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.splt.Panel1.SuspendLayout()
    Me.splt.Panel2.SuspendLayout()
    Me.splt.SuspendLayout()
    Me.bpl.SuspendLayout()
    Me.SuspendLayout()
    '
    'splt
    '
    Me.splt.Dock = System.Windows.Forms.DockStyle.Fill
    Me.splt.FixedPanel = System.Windows.Forms.FixedPanel.Panel2
    Me.splt.Location = New System.Drawing.Point(0, 0)
    Me.splt.Name = "splt"
    Me.splt.Orientation = System.Windows.Forms.Orientation.Horizontal
    '
    'splt.Panel1
    '
    Me.splt.Panel1.Controls.Add(Me.tvw)
    '
    'splt.Panel2
    '
    Me.splt.Panel2.Controls.Add(Me.bpl)
    Me.splt.Panel2.Controls.Add(Me.txtAccessLevelReqd)
    Me.splt.Panel2.Controls.Add(Me.lblAccessLevelReqd)
    Me.splt.Size = New System.Drawing.Size(759, 438)
    Me.splt.SplitterDistance = 342
    Me.splt.TabIndex = 0
    '
    'tvw
    '
    Me.tvw.BackColor = System.Drawing.Color.FromArgb(CType(CType(232, Byte), Integer), CType(CType(232, Byte), Integer), CType(CType(232, Byte), Integer))
    Me.tvw.Dock = System.Windows.Forms.DockStyle.Fill
    Me.tvw.FontHotTracking = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tvw.Location = New System.Drawing.Point(0, 0)
    Me.tvw.Name = "tvw"
    Me.tvw.Size = New System.Drawing.Size(759, 342)
    Me.tvw.TabIndex = 0
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdOK)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bpl.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bpl.Location = New System.Drawing.Point(0, 53)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(759, 39)
    Me.bpl.TabIndex = 2
    '
    'cmdOK
    '
    Me.cmdOK.Location = New System.Drawing.Point(331, 6)
    Me.cmdOK.Name = "cmdOK"
    Me.cmdOK.Size = New System.Drawing.Size(96, 27)
    Me.cmdOK.TabIndex = 0
    Me.cmdOK.Text = "OK"
    Me.cmdOK.UseVisualStyleBackColor = True
    '
    'txtAccessLevelReqd
    '
    Me.txtAccessLevelReqd.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
    Me.txtAccessLevelReqd.Location = New System.Drawing.Point(229, 10)
    Me.txtAccessLevelReqd.Name = "txtAccessLevelReqd"
    Me.txtAccessLevelReqd.ReadOnly = True
    Me.txtAccessLevelReqd.Size = New System.Drawing.Size(459, 20)
    Me.txtAccessLevelReqd.TabIndex = 1
    '
    'lblAccessLevelReqd
    '
    Me.lblAccessLevelReqd.AutoSize = True
    Me.lblAccessLevelReqd.Location = New System.Drawing.Point(66, 13)
    Me.lblAccessLevelReqd.Name = "lblAccessLevelReqd"
    Me.lblAccessLevelReqd.Size = New System.Drawing.Size(120, 13)
    Me.lblAccessLevelReqd.TabIndex = 0
    Me.lblAccessLevelReqd.Text = "Access Level Required:"
    '
    'frmAccessControl
    '
    Me.AcceptButton = Me.cmdOK
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(759, 438)
    Me.Controls.Add(Me.splt)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmAccessControl"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    Me.Text = "System Access Control"
    Me.splt.Panel1.ResumeLayout(False)
    Me.splt.Panel2.ResumeLayout(False)
    Me.splt.Panel2.PerformLayout()
    CType(Me.splt, System.ComponentModel.ISupportInitialize).EndInit()
    Me.splt.ResumeLayout(False)
    Me.bpl.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents splt As System.Windows.Forms.SplitContainer
  Friend WithEvents tvw As CDBNETCL.VistaTreeView
  Friend WithEvents bpl As CDBNETCL.ButtonPanel
  Friend WithEvents txtAccessLevelReqd As System.Windows.Forms.TextBox
  Friend WithEvents lblAccessLevelReqd As System.Windows.Forms.Label
  Friend WithEvents cmdOK As System.Windows.Forms.Button
End Class
