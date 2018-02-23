<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAuthoriseExpenses
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAuthoriseExpenses))
    Me.bpl = New CDBNETCL.ButtonPanel()
    Me.cmdAuthorise = New System.Windows.Forms.Button()
    Me.cmdAuthoriseAll = New System.Windows.Forms.Button()
    Me.cmdOK = New System.Windows.Forms.Button()
    Me.cmdCancel = New System.Windows.Forms.Button()
    Me.splt = New System.Windows.Forms.SplitContainer()
    Me.dgr = New CDBNETCL.DisplayGrid()
    Me.bpl.SuspendLayout()
    CType(Me.splt, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.splt.Panel1.SuspendLayout()
    Me.splt.Panel2.SuspendLayout()
    Me.splt.SuspendLayout()
    Me.SuspendLayout()
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdAuthorise)
    Me.bpl.Controls.Add(Me.cmdAuthoriseAll)
    Me.bpl.Controls.Add(Me.cmdOK)
    Me.bpl.Controls.Add(Me.cmdCancel)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bpl.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bpl.Location = New System.Drawing.Point(0, 6)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(649, 39)
    Me.bpl.TabIndex = 1
    '
    'cmdAuthorise
    '
    Me.cmdAuthorise.Location = New System.Drawing.Point(110, 6)
    Me.cmdAuthorise.Name = "cmdAuthorise"
    Me.cmdAuthorise.Size = New System.Drawing.Size(96, 27)
    Me.cmdAuthorise.TabIndex = 2
    Me.cmdAuthorise.Text = "Authorise"
    Me.cmdAuthorise.UseVisualStyleBackColor = True
    '
    'cmdAuthoriseAll
    '
    Me.cmdAuthoriseAll.Location = New System.Drawing.Point(221, 6)
    Me.cmdAuthoriseAll.Name = "cmdAuthoriseAll"
    Me.cmdAuthoriseAll.Size = New System.Drawing.Size(96, 27)
    Me.cmdAuthoriseAll.TabIndex = 3
    Me.cmdAuthoriseAll.Text = "Authorise All"
    Me.cmdAuthoriseAll.UseVisualStyleBackColor = True
    '
    'cmdOK
    '
    Me.cmdOK.Location = New System.Drawing.Point(332, 6)
    Me.cmdOK.Name = "cmdOK"
    Me.cmdOK.Size = New System.Drawing.Size(96, 27)
    Me.cmdOK.TabIndex = 0
    Me.cmdOK.Text = "OK"
    Me.cmdOK.UseVisualStyleBackColor = True
    '
    'cmdCancel
    '
    Me.cmdCancel.Location = New System.Drawing.Point(443, 6)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(96, 27)
    Me.cmdCancel.TabIndex = 1
    Me.cmdCancel.Text = "Cancel"
    Me.cmdCancel.UseVisualStyleBackColor = True
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
    Me.splt.Panel1.Controls.Add(Me.dgr)
    '
    'splt.Panel2
    '
    Me.splt.Panel2.Controls.Add(Me.bpl)
    Me.splt.Size = New System.Drawing.Size(649, 275)
    Me.splt.SplitterDistance = 226
    Me.splt.TabIndex = 2
    '
    'dgr
    '
    Me.dgr.AccessibleName = "Display Grid"
    Me.dgr.ActiveColumn = 0
    Me.dgr.AllowSorting = True
    Me.dgr.AutoSetHeight = False
    Me.dgr.AutoSetRowHeight = False
    Me.dgr.DisplayTitle = Nothing
    Me.dgr.Dock = System.Windows.Forms.DockStyle.Fill
    Me.dgr.Location = New System.Drawing.Point(0, 0)
    Me.dgr.MaintenanceDesc = Nothing
    Me.dgr.MaxGridRows = 8
    Me.dgr.MultipleSelect = False
    Me.dgr.Name = "dgr"
    Me.dgr.RowCount = 10
    Me.dgr.ShowIfEmpty = False
    Me.dgr.Size = New System.Drawing.Size(649, 226)
    Me.dgr.SuppressHyperLinkFormat = False
    Me.dgr.TabIndex = 0
    '
    'frmAuthoriseExpenses
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(649, 275)
    Me.Controls.Add(Me.splt)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmAuthoriseExpenses"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.Text = "Authorise Expenses"
    Me.bpl.ResumeLayout(False)
    Me.splt.Panel1.ResumeLayout(False)
    Me.splt.Panel2.ResumeLayout(False)
    CType(Me.splt, System.ComponentModel.ISupportInitialize).EndInit()
    Me.splt.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents bpl As CDBNETCL.ButtonPanel
  Friend WithEvents cmdAuthoriseAll As System.Windows.Forms.Button
  Friend WithEvents cmdAuthorise As System.Windows.Forms.Button
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents cmdOK As System.Windows.Forms.Button
  Friend WithEvents splt As System.Windows.Forms.SplitContainer
  Friend WithEvents dgr As CDBNETCL.DisplayGrid
End Class
