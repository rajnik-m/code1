<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmReportMaintenance
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
    Me.components = New System.ComponentModel.Container()
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmReportMaintenance))
    Me.tvw = New CDBNETCL.VistaTreeView()
    Me.dgr = New CDBNETCL.DisplayGrid()
    Me.bpl = New CDBNETCL.ButtonPanel()
    Me.cmdOk = New System.Windows.Forms.Button()
    Me.cmdAmend = New System.Windows.Forms.Button()
    Me.cmdNew = New System.Windows.Forms.Button()
    Me.cmdDelete = New System.Windows.Forms.Button()
    Me.cmdRun = New System.Windows.Forms.Button()
    Me.cmdFind = New System.Windows.Forms.Button()
    Me.splt = New System.Windows.Forms.SplitContainer()
    Me.bpl.SuspendLayout()
    CType(Me.splt, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.splt.Panel1.SuspendLayout()
    Me.splt.Panel2.SuspendLayout()
    Me.splt.SuspendLayout()
    Me.SuspendLayout()
    '
    'tvw
    '
    Me.tvw.BackColor = System.Drawing.Color.FromArgb(CType(CType(232, Byte), Integer), CType(CType(232, Byte), Integer), CType(CType(232, Byte), Integer))
    Me.tvw.Dock = System.Windows.Forms.DockStyle.Fill
    Me.tvw.FontHotTracking = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tvw.Location = New System.Drawing.Point(0, 0)
    Me.tvw.Name = "tvw"
    Me.tvw.Size = New System.Drawing.Size(242, 411)
    Me.tvw.TabIndex = 0
    '
    'dgr
    '
    Me.dgr.AccessibleName = "Display Grid"
    Me.dgr.ActiveColumn = 0
    Me.dgr.AllowColumnResize = True
    Me.dgr.AllowSorting = True
    Me.dgr.AutoSetHeight = False
    Me.dgr.AutoSetRowHeight = False
    Me.dgr.DisplayTitle = Nothing
    Me.dgr.Dock = System.Windows.Forms.DockStyle.Fill
    Me.dgr.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!)
    Me.dgr.Location = New System.Drawing.Point(0, 0)
    Me.dgr.MaintenanceDesc = Nothing
    Me.dgr.MaxGridRows = 6
    Me.dgr.MultipleSelect = False
    Me.dgr.Name = "dgr"
    Me.dgr.RowCount = 10
    Me.dgr.ShowIfEmpty = False
    Me.dgr.Size = New System.Drawing.Size(480, 411)
    Me.dgr.SuppressHyperLinkFormat = False
    Me.dgr.TabIndex = 1
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdOk)
    Me.bpl.Controls.Add(Me.cmdAmend)
    Me.bpl.Controls.Add(Me.cmdNew)
    Me.bpl.Controls.Add(Me.cmdDelete)
    Me.bpl.Controls.Add(Me.cmdRun)
    Me.bpl.Controls.Add(Me.cmdFind)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bpl.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bpl.Location = New System.Drawing.Point(0, 411)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(726, 39)
    Me.bpl.TabIndex = 2
    '
    'cmdOk
    '
    Me.cmdOk.Location = New System.Drawing.Point(37, 6)
    Me.cmdOk.Name = "cmdOk"
    Me.cmdOk.Size = New System.Drawing.Size(96, 27)
    Me.cmdOk.TabIndex = 0
    Me.cmdOk.Text = "OK"
    Me.cmdOk.UseVisualStyleBackColor = True
    '
    'cmdAmend
    '
    Me.cmdAmend.Location = New System.Drawing.Point(148, 6)
    Me.cmdAmend.Name = "cmdAmend"
    Me.cmdAmend.Size = New System.Drawing.Size(96, 27)
    Me.cmdAmend.TabIndex = 1
    Me.cmdAmend.Text = "Amend"
    Me.cmdAmend.UseVisualStyleBackColor = True
    '
    'cmdNew
    '
    Me.cmdNew.Location = New System.Drawing.Point(259, 6)
    Me.cmdNew.Name = "cmdNew"
    Me.cmdNew.Size = New System.Drawing.Size(96, 27)
    Me.cmdNew.TabIndex = 2
    Me.cmdNew.Text = "New"
    Me.cmdNew.UseVisualStyleBackColor = True
    '
    'cmdDelete
    '
    Me.cmdDelete.Location = New System.Drawing.Point(370, 6)
    Me.cmdDelete.Name = "cmdDelete"
    Me.cmdDelete.Size = New System.Drawing.Size(96, 27)
    Me.cmdDelete.TabIndex = 3
    Me.cmdDelete.Text = "Delete"
    Me.cmdDelete.UseVisualStyleBackColor = True
    '
    'cmdRun
    '
    Me.cmdRun.Location = New System.Drawing.Point(481, 6)
    Me.cmdRun.Name = "cmdRun"
    Me.cmdRun.Size = New System.Drawing.Size(96, 27)
    Me.cmdRun.TabIndex = 4
    Me.cmdRun.Text = "Run"
    Me.cmdRun.UseVisualStyleBackColor = True
    '
    'cmdFind
    '
    Me.cmdFind.Location = New System.Drawing.Point(592, 6)
    Me.cmdFind.Name = "cmdFind"
    Me.cmdFind.Size = New System.Drawing.Size(96, 27)
    Me.cmdFind.TabIndex = 5
    Me.cmdFind.Text = "Find"
    Me.cmdFind.UseVisualStyleBackColor = True
    '
    'splt
    '
    Me.splt.Dock = System.Windows.Forms.DockStyle.Fill
    Me.splt.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
    Me.splt.Location = New System.Drawing.Point(0, 0)
    Me.splt.Name = "splt"
    '
    'splt.Panel1
    '
    Me.splt.Panel1.Controls.Add(Me.tvw)
    '
    'splt.Panel2
    '
    Me.splt.Panel2.Controls.Add(Me.dgr)
    Me.splt.Size = New System.Drawing.Size(726, 411)
    Me.splt.SplitterDistance = 242
    Me.splt.TabIndex = 3
    '
    'frmReportMaintenance
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(726, 450)
    Me.Controls.Add(Me.splt)
    Me.Controls.Add(Me.bpl)
    Me.DoubleBuffered = True
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmReportMaintenance"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    Me.Text = "Report Maintenance"
    Me.bpl.ResumeLayout(False)
    Me.splt.Panel1.ResumeLayout(False)
    Me.splt.Panel2.ResumeLayout(False)
    CType(Me.splt, System.ComponentModel.ISupportInitialize).EndInit()
    Me.splt.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents tvw As CDBNETCL.VistaTreeView
  Friend WithEvents dgr As CDBNETCL.DisplayGrid
  Friend WithEvents bpl As CDBNETCL.ButtonPanel
  Friend WithEvents cmdFind As System.Windows.Forms.Button
  Friend WithEvents cmdRun As System.Windows.Forms.Button
  Friend WithEvents cmdDelete As System.Windows.Forms.Button
  Friend WithEvents cmdNew As System.Windows.Forms.Button
  Friend WithEvents cmdAmend As System.Windows.Forms.Button
  Friend WithEvents cmdOk As System.Windows.Forms.Button
  Friend WithEvents splt As System.Windows.Forms.SplitContainer
End Class
