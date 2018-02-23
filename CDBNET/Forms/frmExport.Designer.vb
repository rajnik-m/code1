<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmExport
  Inherits ThemedForm

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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmExport))
    Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
    Me.eplExport = New CDBNETCL.EditPanel()
    Me.bpl = New CDBNETCL.ButtonPanel()
    Me.cmdExport = New System.Windows.Forms.Button()
    Me.cmdCancel = New System.Windows.Forms.Button()
    CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SplitContainer1.Panel1.SuspendLayout()
    Me.SplitContainer1.Panel2.SuspendLayout()
    Me.SplitContainer1.SuspendLayout()
    Me.bpl.SuspendLayout()
    Me.SuspendLayout()
    '
    'SplitContainer1
    '
    Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
    Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
    Me.SplitContainer1.Margin = New System.Windows.Forms.Padding(2)
    Me.SplitContainer1.Name = "SplitContainer1"
    Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
    '
    'SplitContainer1.Panel1
    '
    Me.SplitContainer1.Panel1.Controls.Add(Me.eplExport)
    '
    'SplitContainer1.Panel2
    '
    Me.SplitContainer1.Panel2.Controls.Add(Me.bpl)
    Me.SplitContainer1.Panel2MinSize = 39
    Me.SplitContainer1.Size = New System.Drawing.Size(700, 201)
    Me.SplitContainer1.SplitterDistance = 158
    Me.SplitContainer1.SplitterWidth = 3
    Me.SplitContainer1.TabIndex = 0
    Me.SplitContainer1.TabStop = False
    '
    'eplExport
    '
    Me.eplExport.AddressChanged = False
    Me.eplExport.BackColor = System.Drawing.Color.Transparent
    Me.eplExport.DataChanged = False
    Me.eplExport.Dock = System.Windows.Forms.DockStyle.Fill
    Me.eplExport.Location = New System.Drawing.Point(0, 0)
    Me.eplExport.Margin = New System.Windows.Forms.Padding(2)
    Me.eplExport.Name = "eplExport"
    Me.eplExport.Recipients = Nothing
    Me.eplExport.Size = New System.Drawing.Size(700, 158)
    Me.eplExport.SuppressDrawing = False
    Me.eplExport.TabIndex = 0
    Me.eplExport.TabSelectedIndex = 0
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdExport)
    Me.bpl.Controls.Add(Me.cmdCancel)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bpl.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bpl.Location = New System.Drawing.Point(0, 1)
    Me.bpl.Margin = New System.Windows.Forms.Padding(2)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(700, 39)
    Me.bpl.TabIndex = 0
    '
    'cmdExport
    '
    Me.cmdExport.Location = New System.Drawing.Point(246, 6)
    Me.cmdExport.Margin = New System.Windows.Forms.Padding(2)
    Me.cmdExport.Name = "cmdExport"
    Me.cmdExport.Size = New System.Drawing.Size(96, 27)
    Me.cmdExport.TabIndex = 1
    Me.cmdExport.Text = "E&xport"
    Me.cmdExport.UseVisualStyleBackColor = True
    '
    'cmdCancel
    '
    Me.cmdCancel.Location = New System.Drawing.Point(357, 6)
    Me.cmdCancel.Margin = New System.Windows.Forms.Padding(2)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(96, 27)
    Me.cmdCancel.TabIndex = 2
    Me.cmdCancel.Text = "&Close"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'frmExport
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(700, 201)
    Me.Controls.Add(Me.SplitContainer1)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Margin = New System.Windows.Forms.Padding(2)
    Me.Name = "frmExport"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    Me.Text = "frmExport"
    Me.SplitContainer1.Panel1.ResumeLayout(False)
    Me.SplitContainer1.Panel2.ResumeLayout(False)
    CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
    Me.SplitContainer1.ResumeLayout(False)
    Me.bpl.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
  Friend WithEvents bpl As CDBNETCL.ButtonPanel
  Friend WithEvents cmdExport As System.Windows.Forms.Button
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents eplExport As CDBNETCL.EditPanel
End Class
