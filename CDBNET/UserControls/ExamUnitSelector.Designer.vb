<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ExamUnitSelector
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
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
    Me.spl = New System.Windows.Forms.SplitContainer()
    Me.sel = New CDBNETCL.ExamSelector()
    Me.tab = New CDBNETCL.TabControl()
    Me.tbp1 = New System.Windows.Forms.TabPage()
    Me.dpl = New CDBNETCL.DisplayPanel()
    Me.tbp2 = New System.Windows.Forms.TabPage()
    Me.dgr = New CDBNETCL.DisplayGrid()
    Me.tbp3 = New System.Windows.Forms.TabPage()
    Me.dgrActivities = New CDBNETCL.DisplayGrid()
    Me.tbp4 = New System.Windows.Forms.TabPage()
    Me.dgrGradeHistory = New CDBNETCL.DisplayGrid()
    Me.tbp5 = New System.Windows.Forms.TabPage()
    Me.dgrWorkstreams = New CDBNETCL.DisplayGrid()
    CType(Me.spl, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.spl.Panel1.SuspendLayout()
    Me.spl.Panel2.SuspendLayout()
    Me.spl.SuspendLayout()
    Me.tab.SuspendLayout()
    Me.tbp1.SuspendLayout()
    Me.tbp2.SuspendLayout()
    Me.tbp3.SuspendLayout()
    Me.tbp4.SuspendLayout()
    Me.tbp5.SuspendLayout()
    Me.SuspendLayout()
    '
    'spl
    '
    Me.spl.BackColor = System.Drawing.SystemColors.Control
    Me.spl.Dock = System.Windows.Forms.DockStyle.Fill
    Me.spl.Location = New System.Drawing.Point(0, 0)
    Me.spl.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
    Me.spl.Name = "spl"
    '
    'spl.Panel1
    '
    Me.spl.Panel1.Controls.Add(Me.sel)
    '
    'spl.Panel2
    '
    Me.spl.Panel2.Controls.Add(Me.tab)
    Me.spl.Size = New System.Drawing.Size(584, 332)
    Me.spl.SplitterDistance = 255
    Me.spl.SplitterWidth = 5
    Me.spl.TabIndex = 0
    '
    'sel
    '
    Me.sel.AutoSelectParents = False
    Me.sel.BackColor = System.Drawing.Color.Transparent
    Me.sel.Dock = System.Windows.Forms.DockStyle.Fill
    Me.sel.ExamMaintenance = False
    Me.sel.Location = New System.Drawing.Point(0, 0)
    Me.sel.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
    Me.sel.Name = "sel"
    Me.sel.Size = New System.Drawing.Size(255, 332)
    Me.sel.TabIndex = 0
    Me.sel.TreeContextMenu = Nothing
    '
    'tab
    '
    Me.tab.Controls.Add(Me.tbp1)
    Me.tab.Controls.Add(Me.tbp2)
    Me.tab.Controls.Add(Me.tbp3)
    Me.tab.Controls.Add(Me.tbp4)
    Me.tab.Controls.Add(Me.tbp5)
    Me.tab.Dock = System.Windows.Forms.DockStyle.Fill
    Me.tab.ItemSize = New System.Drawing.Size(85, 20)
    Me.tab.Location = New System.Drawing.Point(0, 0)
    Me.tab.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
    Me.tab.Name = "tab"
    Me.tab.SelectedIndex = 0
    Me.tab.Size = New System.Drawing.Size(324, 332)
    Me.tab.SizeMode = System.Windows.Forms.TabSizeMode.Fixed
    Me.tab.TabIndex = 1
    '
    'tbp1
    '
    Me.tbp1.Controls.Add(Me.dpl)
    Me.tbp1.Location = New System.Drawing.Point(4, 24)
    Me.tbp1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
    Me.tbp1.Name = "tbp1"
    Me.tbp1.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
    Me.tbp1.Size = New System.Drawing.Size(316, 304)
    Me.tbp1.TabIndex = 0
    Me.tbp1.Text = "Detail"
    '
    'dpl
    '
    Me.dpl.BackColor = System.Drawing.Color.Transparent
    Me.dpl.DataSelectionType = CDBNETCL.CareNetServices.XMLContactDataSelectionTypes.xcdtNone
    Me.dpl.Dock = System.Windows.Forms.DockStyle.Fill
    Me.dpl.Location = New System.Drawing.Point(4, 4)
    Me.dpl.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
    Me.dpl.Name = "dpl"
    Me.dpl.Size = New System.Drawing.Size(308, 296)
    Me.dpl.TabIndex = 0
    '
    'tbp2
    '
    Me.tbp2.Controls.Add(Me.dgr)
    Me.tbp2.Location = New System.Drawing.Point(4, 24)
    Me.tbp2.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
    Me.tbp2.Name = "tbp2"
    Me.tbp2.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
    Me.tbp2.Size = New System.Drawing.Size(376, 304)
    Me.tbp2.TabIndex = 1
    Me.tbp2.Text = "List"
    Me.tbp2.UseVisualStyleBackColor = True
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
    Me.dgr.Location = New System.Drawing.Point(4, 4)
    Me.dgr.MaintenanceDesc = Nothing
    Me.dgr.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
    Me.dgr.MaxGridRows = 8
    Me.dgr.MultipleSelect = False
    Me.dgr.Name = "dgr"
    Me.dgr.RowCount = 10
    Me.dgr.ShowIfEmpty = False
    Me.dgr.Size = New System.Drawing.Size(368, 296)
    Me.dgr.SuppressHyperLinkFormat = False
    Me.dgr.TabIndex = 0
    '
    'tbp3
    '
    Me.tbp3.Controls.Add(Me.dgrActivities)
    Me.tbp3.Location = New System.Drawing.Point(4, 24)
    Me.tbp3.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
    Me.tbp3.Name = "tbp3"
    Me.tbp3.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
    Me.tbp3.Size = New System.Drawing.Size(376, 304)
    Me.tbp3.TabIndex = 2
    Me.tbp3.Text = "Categories"
    Me.tbp3.UseVisualStyleBackColor = True
    '
    'dgrActivities
    '
    Me.dgrActivities.AccessibleName = "Display Grid Categories"
    Me.dgrActivities.ActiveColumn = 0
    Me.dgrActivities.AllowColumnResize = True
    Me.dgrActivities.AllowSorting = True
    Me.dgrActivities.AutoSetHeight = False
    Me.dgrActivities.AutoSetRowHeight = False
    Me.dgrActivities.DisplayTitle = Nothing
    Me.dgrActivities.Dock = System.Windows.Forms.DockStyle.Fill
    Me.dgrActivities.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!)
    Me.dgrActivities.Location = New System.Drawing.Point(4, 4)
    Me.dgrActivities.MaintenanceDesc = Nothing
    Me.dgrActivities.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
    Me.dgrActivities.MaxGridRows = 8
    Me.dgrActivities.MultipleSelect = False
    Me.dgrActivities.Name = "dgrActivities"
    Me.dgrActivities.RowCount = 10
    Me.dgrActivities.ShowIfEmpty = False
    Me.dgrActivities.Size = New System.Drawing.Size(368, 296)
    Me.dgrActivities.SuppressHyperLinkFormat = False
    Me.dgrActivities.TabIndex = 0
    '
    'tbp4
    '
    Me.tbp4.Controls.Add(Me.dgrGradeHistory)
    Me.tbp4.Location = New System.Drawing.Point(4, 24)
    Me.tbp4.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
    Me.tbp4.Name = "tbp4"
    Me.tbp4.Size = New System.Drawing.Size(376, 304)
    Me.tbp4.TabIndex = 3
    Me.tbp4.Text = "Grade History"
    Me.tbp4.UseVisualStyleBackColor = True
    '
    'dgrGradeHistory
    '
    Me.dgrGradeHistory.AccessibleName = "Display Grid"
    Me.dgrGradeHistory.ActiveColumn = 0
    Me.dgrGradeHistory.AllowColumnResize = True
    Me.dgrGradeHistory.AllowSorting = True
    Me.dgrGradeHistory.AutoSetHeight = False
    Me.dgrGradeHistory.AutoSetRowHeight = False
    Me.dgrGradeHistory.DisplayTitle = Nothing
    Me.dgrGradeHistory.Dock = System.Windows.Forms.DockStyle.Fill
    Me.dgrGradeHistory.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!)
    Me.dgrGradeHistory.Location = New System.Drawing.Point(0, 0)
    Me.dgrGradeHistory.MaintenanceDesc = Nothing
    Me.dgrGradeHistory.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
    Me.dgrGradeHistory.MaxGridRows = 8
    Me.dgrGradeHistory.MultipleSelect = False
    Me.dgrGradeHistory.Name = "dgrGradeHistory"
    Me.dgrGradeHistory.RowCount = 10
    Me.dgrGradeHistory.ShowIfEmpty = False
    Me.dgrGradeHistory.Size = New System.Drawing.Size(376, 304)
    Me.dgrGradeHistory.SuppressHyperLinkFormat = False
    Me.dgrGradeHistory.TabIndex = 0
    '
    'tbp5
    '
    Me.tbp5.Controls.Add(Me.dgrWorkstreams)
    Me.tbp5.Location = New System.Drawing.Point(4, 24)
    Me.tbp5.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
    Me.tbp5.Name = "tbp5"
    Me.tbp5.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
    Me.tbp5.Size = New System.Drawing.Size(376, 304)
    Me.tbp5.TabIndex = 4
    Me.tbp5.Text = "Workstreams"
    Me.tbp5.UseVisualStyleBackColor = True
    '
    'dgrWorkstreams
    '
    Me.dgrWorkstreams.AccessibleName = "Display Grid"
    Me.dgrWorkstreams.ActiveColumn = 0
    Me.dgrWorkstreams.AllowColumnResize = True
    Me.dgrWorkstreams.AllowSorting = True
    Me.dgrWorkstreams.AutoSetHeight = False
    Me.dgrWorkstreams.AutoSetRowHeight = False
    Me.dgrWorkstreams.DisplayTitle = Nothing
    Me.dgrWorkstreams.Dock = System.Windows.Forms.DockStyle.Fill
    Me.dgrWorkstreams.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!)
    Me.dgrWorkstreams.Location = New System.Drawing.Point(4, 4)
    Me.dgrWorkstreams.MaintenanceDesc = Nothing
    Me.dgrWorkstreams.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
    Me.dgrWorkstreams.MaxGridRows = 8
    Me.dgrWorkstreams.MultipleSelect = False
    Me.dgrWorkstreams.Name = "dgrWorkstreams"
    Me.dgrWorkstreams.RowCount = 10
    Me.dgrWorkstreams.ShowIfEmpty = False
    Me.dgrWorkstreams.Size = New System.Drawing.Size(368, 296)
    Me.dgrWorkstreams.SuppressHyperLinkFormat = False
    Me.dgrWorkstreams.TabIndex = 1
    '
    'ExamUnitSelector
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.BackColor = System.Drawing.Color.Red
    Me.Controls.Add(Me.spl)
    Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
    Me.Name = "ExamUnitSelector"
    Me.Size = New System.Drawing.Size(584, 332)
    Me.spl.Panel1.ResumeLayout(False)
    Me.spl.Panel2.ResumeLayout(False)
    CType(Me.spl, System.ComponentModel.ISupportInitialize).EndInit()
    Me.spl.ResumeLayout(False)
    Me.tab.ResumeLayout(False)
    Me.tbp1.ResumeLayout(False)
    Me.tbp2.ResumeLayout(False)
    Me.tbp3.ResumeLayout(False)
    Me.tbp4.ResumeLayout(False)
    Me.tbp5.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents spl As System.Windows.Forms.SplitContainer
  Friend WithEvents sel As CDBNETCL.ExamSelector
  Friend WithEvents dpl As CDBNETCL.DisplayPanel
  Friend WithEvents tab As CDBNETCL.TabControl
  Friend WithEvents tbp1 As System.Windows.Forms.TabPage
  Friend WithEvents tbp2 As System.Windows.Forms.TabPage
  Friend WithEvents dgr As CDBNETCL.DisplayGrid
  Friend WithEvents tbp3 As System.Windows.Forms.TabPage
  Friend WithEvents dgrActivities As CDBNETCL.DisplayGrid
  Friend WithEvents tbp4 As System.Windows.Forms.TabPage
  Friend WithEvents dgrGradeHistory As CDBNETCL.DisplayGrid
  Friend WithEvents tbp5 As System.Windows.Forms.TabPage
  Friend WithEvents dgrWorkstreams As CDBNETCL.DisplayGrid

End Class
