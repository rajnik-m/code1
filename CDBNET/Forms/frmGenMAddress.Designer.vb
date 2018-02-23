<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmGenMAddress
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
    Me.components = New System.ComponentModel.Container()
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmGenMAddress))
    Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
    Me.drg = New CDBNETCL.DisplayGrid()
    Me.ButtonPanel1 = New CDBNETCL.ButtonPanel()
    Me.cmdOk = New System.Windows.Forms.Button()
    Me.cmdCancel = New System.Windows.Forms.Button()
    CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SplitContainer1.Panel1.SuspendLayout()
    Me.SplitContainer1.Panel2.SuspendLayout()
    Me.SplitContainer1.SuspendLayout()
    Me.ButtonPanel1.SuspendLayout()
    Me.SuspendLayout()
    '
    'SplitContainer1
    '
    Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
    Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
    Me.SplitContainer1.Name = "SplitContainer1"
    Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
    '
    'SplitContainer1.Panel1
    '
    Me.SplitContainer1.Panel1.Controls.Add(Me.drg)
    '
    'SplitContainer1.Panel2
    '
    Me.SplitContainer1.Panel2.Controls.Add(Me.ButtonPanel1)
    Me.SplitContainer1.Panel2MinSize = 10
    Me.SplitContainer1.Size = New System.Drawing.Size(471, 275)
    Me.SplitContainer1.SplitterDistance = 232
    Me.SplitContainer1.TabIndex = 0
    '
    'drg
    '
    Me.drg.AccessibleName = "Display Grid"
    Me.drg.ActiveColumn = 0
    Me.drg.AllowSorting = True
    Me.drg.AutoSetHeight = False
    Me.drg.AutoSetRowHeight = False
    Me.drg.DisplayTitle = Nothing
    Me.drg.Dock = System.Windows.Forms.DockStyle.Fill
    Me.drg.Location = New System.Drawing.Point(0, 0)
    Me.drg.MaintenanceDesc = Nothing
    Me.drg.MaxGridRows = 8
    Me.drg.MultipleSelect = False
    Me.drg.Name = "drg"
    Me.drg.RowCount = 10
    Me.drg.ShowIfEmpty = False
    Me.drg.Size = New System.Drawing.Size(471, 232)
    Me.drg.SuppressHyperLinkFormat = False
    Me.drg.TabIndex = 3
    '
    'ButtonPanel1
    '
    Me.ButtonPanel1.Controls.Add(Me.cmdOk)
    Me.ButtonPanel1.Controls.Add(Me.cmdCancel)
    Me.ButtonPanel1.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.ButtonPanel1.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.ButtonPanel1.Location = New System.Drawing.Point(0, 0)
    Me.ButtonPanel1.Name = "ButtonPanel1"
    Me.ButtonPanel1.Size = New System.Drawing.Size(471, 39)
    Me.ButtonPanel1.TabIndex = 4
    '
    'cmdOk
    '
    Me.cmdOk.Location = New System.Drawing.Point(132, 6)
    Me.cmdOk.Name = "cmdOk"
    Me.cmdOk.Size = New System.Drawing.Size(96, 27)
    Me.cmdOk.TabIndex = 0
    Me.cmdOk.Text = "OK"
    Me.cmdOk.UseVisualStyleBackColor = True
    '
    'cmdCancel
    '
    Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdCancel.Location = New System.Drawing.Point(243, 6)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(96, 27)
    Me.cmdCancel.TabIndex = 1
    Me.cmdCancel.Text = "Cancel"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'frmGenMAddress
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(471, 275)
    Me.Controls.Add(Me.SplitContainer1)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmGenMAddress"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    Me.Tag = ""
    Me.Text = "Selection Manager - Contact Address Selection"
    Me.SplitContainer1.Panel1.ResumeLayout(False)
    Me.SplitContainer1.Panel2.ResumeLayout(False)
    CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
    Me.SplitContainer1.ResumeLayout(False)
    Me.ButtonPanel1.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
  Friend WithEvents drg As CDBNETCL.DisplayGrid
  Friend WithEvents ButtonPanel1 As CDBNETCL.ButtonPanel
  Friend WithEvents cmdOk As System.Windows.Forms.Button
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
End Class
