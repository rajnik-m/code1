<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmConfigurationMaintenance
  Inherits CDBNETCL.ThemedForm

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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmConfigurationMaintenance))
    Me.splMain = New System.Windows.Forms.SplitContainer()
    Me.splConfig = New System.Windows.Forms.SplitContainer()
    Me.tvw = New CDBNETCL.VistaTreeView()
    Me.dgr = New CDBNETCL.DisplayGrid()
    Me.lblNotes = New System.Windows.Forms.Label()
    Me.lblConfigName = New System.Windows.Forms.Label()
    Me.txtNotes = New System.Windows.Forms.TextBox()
    Me.txtConfigName = New System.Windows.Forms.TextBox()
    Me.bpl = New CDBNETCL.ButtonPanel()
    Me.cmdOK = New System.Windows.Forms.Button()
    Me.cmdAmend = New System.Windows.Forms.Button()
    Me.cmdNew = New System.Windows.Forms.Button()
    Me.cmdDelete = New System.Windows.Forms.Button()
    Me.cmdFind = New System.Windows.Forms.Button()
    Me.lblConfigDefault = New System.Windows.Forms.Label()
    Me.txtDefaultValue = New System.Windows.Forms.TextBox()
    CType(Me.splMain, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.splMain.Panel1.SuspendLayout()
    Me.splMain.Panel2.SuspendLayout()
    Me.splMain.SuspendLayout()
    CType(Me.splConfig, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.splConfig.Panel1.SuspendLayout()
    Me.splConfig.Panel2.SuspendLayout()
    Me.splConfig.SuspendLayout()
    Me.bpl.SuspendLayout()
    Me.SuspendLayout()
    '
    'splMain
    '
    Me.splMain.Dock = System.Windows.Forms.DockStyle.Fill
    Me.splMain.Location = New System.Drawing.Point(0, 0)
    Me.splMain.Name = "splMain"
    Me.splMain.Orientation = System.Windows.Forms.Orientation.Horizontal
    '
    'splMain.Panel1
    '
    Me.splMain.Panel1.Controls.Add(Me.splConfig)
    '
    'splMain.Panel2
    '
    Me.splMain.Panel2.Controls.Add(Me.lblConfigDefault)
    Me.splMain.Panel2.Controls.Add(Me.txtDefaultValue)
    Me.splMain.Panel2.Controls.Add(Me.lblNotes)
    Me.splMain.Panel2.Controls.Add(Me.lblConfigName)
    Me.splMain.Panel2.Controls.Add(Me.txtNotes)
    Me.splMain.Panel2.Controls.Add(Me.txtConfigName)
    Me.splMain.Size = New System.Drawing.Size(783, 555)
    Me.splMain.SplitterDistance = 379
    Me.splMain.TabIndex = 0
    '
    'splConfig
    '
    Me.splConfig.Dock = System.Windows.Forms.DockStyle.Fill
    Me.splConfig.Location = New System.Drawing.Point(0, 0)
    Me.splConfig.Name = "splConfig"
    '
    'splConfig.Panel1
    '
    Me.splConfig.Panel1.Controls.Add(Me.tvw)
    '
    'splConfig.Panel2
    '
    Me.splConfig.Panel2.Controls.Add(Me.dgr)
    Me.splConfig.Size = New System.Drawing.Size(783, 379)
    Me.splConfig.SplitterDistance = 259
    Me.splConfig.TabIndex = 0
    '
    'tvw
    '
    Me.tvw.BackColor = System.Drawing.Color.FromArgb(CType(CType(232, Byte), Integer), CType(CType(232, Byte), Integer), CType(CType(232, Byte), Integer))
    Me.tvw.Dock = System.Windows.Forms.DockStyle.Fill
    Me.tvw.FontHotTracking = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tvw.Location = New System.Drawing.Point(0, 0)
    Me.tvw.Name = "tvw"
    Me.tvw.Size = New System.Drawing.Size(259, 379)
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
    Me.dgr.Size = New System.Drawing.Size(520, 379)
    Me.dgr.SuppressHyperLinkFormat = False
    Me.dgr.TabIndex = 0
    '
    'lblNotes
    '
    Me.lblNotes.AutoSize = True
    Me.lblNotes.Location = New System.Drawing.Point(12, 88)
    Me.lblNotes.Name = "lblNotes"
    Me.lblNotes.Size = New System.Drawing.Size(38, 13)
    Me.lblNotes.TabIndex = 3
    Me.lblNotes.Text = "Notes:"
    '
    'lblConfigName
    '
    Me.lblConfigName.AutoSize = True
    Me.lblConfigName.Location = New System.Drawing.Point(12, 13)
    Me.lblConfigName.Name = "lblConfigName"
    Me.lblConfigName.Size = New System.Drawing.Size(71, 13)
    Me.lblConfigName.TabIndex = 2
    Me.lblConfigName.Text = "Config Name:"
    '
    'txtNotes
    '
    Me.txtNotes.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.txtNotes.Location = New System.Drawing.Point(153, 60)
    Me.txtNotes.Multiline = True
    Me.txtNotes.Name = "txtNotes"
    Me.txtNotes.ReadOnly = True
    Me.txtNotes.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
    Me.txtNotes.Size = New System.Drawing.Size(618, 67)
    Me.txtNotes.TabIndex = 1
    '
    'txtConfigName
    '
    Me.txtConfigName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.txtConfigName.Location = New System.Drawing.Point(153, 8)
    Me.txtConfigName.Name = "txtConfigName"
    Me.txtConfigName.ReadOnly = True
    Me.txtConfigName.Size = New System.Drawing.Size(618, 20)
    Me.txtConfigName.TabIndex = 0
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdOK)
    Me.bpl.Controls.Add(Me.cmdAmend)
    Me.bpl.Controls.Add(Me.cmdNew)
    Me.bpl.Controls.Add(Me.cmdDelete)
    Me.bpl.Controls.Add(Me.cmdFind)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bpl.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bpl.Location = New System.Drawing.Point(0, 516)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(783, 39)
    Me.bpl.TabIndex = 1
    '
    'cmdOK
    '
    Me.cmdOK.Location = New System.Drawing.Point(121, 6)
    Me.cmdOK.Name = "cmdOK"
    Me.cmdOK.Size = New System.Drawing.Size(96, 27)
    Me.cmdOK.TabIndex = 4
    Me.cmdOK.Text = "OK"
    Me.cmdOK.UseVisualStyleBackColor = True
    '
    'cmdAmend
    '
    Me.cmdAmend.Location = New System.Drawing.Point(232, 6)
    Me.cmdAmend.Name = "cmdAmend"
    Me.cmdAmend.Size = New System.Drawing.Size(96, 27)
    Me.cmdAmend.TabIndex = 3
    Me.cmdAmend.Text = "&Amend"
    Me.cmdAmend.UseVisualStyleBackColor = True
    '
    'cmdNew
    '
    Me.cmdNew.Location = New System.Drawing.Point(343, 6)
    Me.cmdNew.Name = "cmdNew"
    Me.cmdNew.Size = New System.Drawing.Size(96, 27)
    Me.cmdNew.TabIndex = 2
    Me.cmdNew.Text = "&New"
    Me.cmdNew.UseVisualStyleBackColor = True
    '
    'cmdDelete
    '
    Me.cmdDelete.Location = New System.Drawing.Point(454, 6)
    Me.cmdDelete.Name = "cmdDelete"
    Me.cmdDelete.Size = New System.Drawing.Size(96, 27)
    Me.cmdDelete.TabIndex = 1
    Me.cmdDelete.Text = "&Delete"
    Me.cmdDelete.UseVisualStyleBackColor = True
    '
    'cmdFind
    '
    Me.cmdFind.Location = New System.Drawing.Point(565, 6)
    Me.cmdFind.Name = "cmdFind"
    Me.cmdFind.Size = New System.Drawing.Size(96, 27)
    Me.cmdFind.TabIndex = 0
    Me.cmdFind.Text = "&Find"
    Me.cmdFind.UseVisualStyleBackColor = True
    '
    'lblConfigDefault
    '
    Me.lblConfigDefault.AutoSize = True
    Me.lblConfigDefault.Location = New System.Drawing.Point(12, 39)
    Me.lblConfigDefault.Name = "lblConfigDefault"
    Me.lblConfigDefault.Size = New System.Drawing.Size(74, 13)
    Me.lblConfigDefault.TabIndex = 5
    Me.lblConfigDefault.Text = "Default Value:"
    '
    'txtDefaultValue
    '
    Me.txtDefaultValue.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.txtDefaultValue.Location = New System.Drawing.Point(153, 34)
    Me.txtDefaultValue.Name = "txtDefaultValue"
    Me.txtDefaultValue.ReadOnly = True
    Me.txtDefaultValue.Size = New System.Drawing.Size(618, 20)
    Me.txtDefaultValue.TabIndex = 4
    '
    'frmConfigurationMaintenance
    '
    Me.ClientSize = New System.Drawing.Size(783, 555)
    Me.Controls.Add(Me.bpl)
    Me.Controls.Add(Me.splMain)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmConfigurationMaintenance"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    Me.Text = "System Configuration"
    Me.splMain.Panel1.ResumeLayout(False)
    Me.splMain.Panel2.ResumeLayout(False)
    Me.splMain.Panel2.PerformLayout()
    CType(Me.splMain, System.ComponentModel.ISupportInitialize).EndInit()
    Me.splMain.ResumeLayout(False)
    Me.splConfig.Panel1.ResumeLayout(False)
    Me.splConfig.Panel2.ResumeLayout(False)
    CType(Me.splConfig, System.ComponentModel.ISupportInitialize).EndInit()
    Me.splConfig.ResumeLayout(False)
    Me.bpl.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents splMain As System.Windows.Forms.SplitContainer
  Friend WithEvents splConfig As System.Windows.Forms.SplitContainer
  Friend WithEvents tvw As CDBNETCL.VistaTreeView
  Friend WithEvents lblNotes As System.Windows.Forms.Label
  Friend WithEvents lblConfigName As System.Windows.Forms.Label
  Friend WithEvents txtNotes As System.Windows.Forms.TextBox
  Friend WithEvents txtConfigName As System.Windows.Forms.TextBox
  Friend WithEvents dgr As CDBNETCL.DisplayGrid
  Friend WithEvents bpl As CDBNETCL.ButtonPanel
  Friend WithEvents cmdOK As System.Windows.Forms.Button
  Friend WithEvents cmdAmend As System.Windows.Forms.Button
  Friend WithEvents cmdNew As System.Windows.Forms.Button
  Friend WithEvents cmdDelete As System.Windows.Forms.Button
  Friend WithEvents cmdFind As System.Windows.Forms.Button
  Friend WithEvents lblConfigDefault As System.Windows.Forms.Label
  Friend WithEvents txtDefaultValue As System.Windows.Forms.TextBox

End Class
