<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCustomiseToolBar
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCustomiseToolBar))
    Me.cmdAdd = New System.Windows.Forms.Button()
    Me.cmdRemove = New System.Windows.Forms.Button()
    Me.cmdAddAll = New System.Windows.Forms.Button()
    Me.cmdRemoveAll = New System.Windows.Forms.Button()
    Me.ButtonPanel1 = New CDBNETCL.ButtonPanel()
    Me.cmdDelete = New System.Windows.Forms.Button()
    Me.cmdNew = New System.Windows.Forms.Button()
    Me.cmdEdit = New System.Windows.Forms.Button()
    Me.Button1 = New System.Windows.Forms.Button()
    Me.Button2 = New System.Windows.Forms.Button()
    Me.lblAvailable = New System.Windows.Forms.Label()
    Me.TableLayoutPanel = New System.Windows.Forms.TableLayoutPanel()
    Me.lstAvailable = New System.Windows.Forms.ListBox()
    Me.SplitHeader = New System.Windows.Forms.SplitContainer()
    Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
    Me.lvwSelected = New System.Windows.Forms.ListView()
    Me.ColumnHeader3 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
    Me.ColumnHeader4 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
    Me.Panel1 = New System.Windows.Forms.Panel()
    Me.bpl2 = New CDBNETCL.ButtonPanel()
    Me.btnAdd = New System.Windows.Forms.Button()
    Me.btnRemove = New System.Windows.Forms.Button()
    Me.btnAddAll = New System.Windows.Forms.Button()
    Me.btnRemoveAll = New System.Windows.Forms.Button()
    Me.lvwAvailable = New System.Windows.Forms.ListView()
    Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
    Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
    Me.pnl = New System.Windows.Forms.Panel()
    Me.chkLabelsBelow = New System.Windows.Forms.CheckBox()
    Me.txtLabel = New System.Windows.Forms.TextBox()
    Me.txtToolTip = New System.Windows.Forms.TextBox()
    Me.lblLabel = New System.Windows.Forms.Label()
    Me.lblToolTip = New System.Windows.Forms.Label()
    Me.bpl1 = New CDBNETCL.ButtonPanel()
    Me.cmdOk = New System.Windows.Forms.Button()
    Me.cmdApply = New System.Windows.Forms.Button()
    Me.cmdDefault = New System.Windows.Forms.Button()
    Me.cmdCancel = New System.Windows.Forms.Button()
    Me.ButtonPanel1.SuspendLayout()
    Me.TableLayoutPanel.SuspendLayout()
    CType(Me.SplitHeader, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SplitHeader.Panel1.SuspendLayout()
    Me.SplitHeader.Panel2.SuspendLayout()
    Me.SplitHeader.SuspendLayout()
    Me.TableLayoutPanel1.SuspendLayout()
    Me.Panel1.SuspendLayout()
    Me.bpl2.SuspendLayout()
    Me.pnl.SuspendLayout()
    Me.bpl1.SuspendLayout()
    Me.SuspendLayout()
    '
    'cmdAdd
    '
    Me.cmdAdd.Location = New System.Drawing.Point(8, 15)
    Me.cmdAdd.Name = "cmdAdd"
    Me.cmdAdd.Size = New System.Drawing.Size(96, 27)
    Me.cmdAdd.TabIndex = 0
    Me.cmdAdd.Text = "&Add >"
    Me.cmdAdd.UseVisualStyleBackColor = True
    '
    'cmdRemove
    '
    Me.cmdRemove.Location = New System.Drawing.Point(8, 57)
    Me.cmdRemove.Name = "cmdRemove"
    Me.cmdRemove.Size = New System.Drawing.Size(96, 27)
    Me.cmdRemove.TabIndex = 1
    Me.cmdRemove.Text = "< &Remove"
    Me.cmdRemove.UseVisualStyleBackColor = True
    '
    'cmdAddAll
    '
    Me.cmdAddAll.Location = New System.Drawing.Point(8, 99)
    Me.cmdAddAll.Name = "cmdAddAll"
    Me.cmdAddAll.Size = New System.Drawing.Size(96, 27)
    Me.cmdAddAll.TabIndex = 2
    Me.cmdAddAll.Text = "Add All &>>"
    Me.cmdAddAll.UseVisualStyleBackColor = True
    '
    'cmdRemoveAll
    '
    Me.cmdRemoveAll.Location = New System.Drawing.Point(8, 141)
    Me.cmdRemoveAll.Name = "cmdRemoveAll"
    Me.cmdRemoveAll.Size = New System.Drawing.Size(96, 27)
    Me.cmdRemoveAll.TabIndex = 3
    Me.cmdRemoveAll.Text = "&<< Remove All"
    Me.cmdRemoveAll.UseVisualStyleBackColor = True
    '
    'ButtonPanel1
    '
    Me.ButtonPanel1.Controls.Add(Me.cmdDelete)
    Me.ButtonPanel1.Controls.Add(Me.cmdNew)
    Me.ButtonPanel1.Controls.Add(Me.cmdEdit)
    Me.ButtonPanel1.Controls.Add(Me.Button1)
    Me.ButtonPanel1.Controls.Add(Me.Button2)
    Me.ButtonPanel1.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.ButtonPanel1.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.ButtonPanel1.Location = New System.Drawing.Point(0, 274)
    Me.ButtonPanel1.Name = "ButtonPanel1"
    Me.ButtonPanel1.Size = New System.Drawing.Size(528, 39)
    Me.ButtonPanel1.TabIndex = 10
    '
    'cmdDelete
    '
    Me.cmdDelete.Location = New System.Drawing.Point(0, 6)
    Me.cmdDelete.Name = "cmdDelete"
    Me.cmdDelete.Size = New System.Drawing.Size(96, 27)
    Me.cmdDelete.TabIndex = 4
    Me.cmdDelete.Text = "Delete"
    Me.cmdDelete.UseVisualStyleBackColor = True
    '
    'cmdNew
    '
    Me.cmdNew.Location = New System.Drawing.Point(108, 6)
    Me.cmdNew.Name = "cmdNew"
    Me.cmdNew.Size = New System.Drawing.Size(96, 27)
    Me.cmdNew.TabIndex = 3
    Me.cmdNew.Text = "New"
    Me.cmdNew.UseVisualStyleBackColor = True
    '
    'cmdEdit
    '
    Me.cmdEdit.Location = New System.Drawing.Point(216, 6)
    Me.cmdEdit.Name = "cmdEdit"
    Me.cmdEdit.Size = New System.Drawing.Size(96, 27)
    Me.cmdEdit.TabIndex = 2
    Me.cmdEdit.Text = "Edit"
    Me.cmdEdit.UseVisualStyleBackColor = True
    '
    'Button1
    '
    Me.Button1.DialogResult = System.Windows.Forms.DialogResult.OK
    Me.Button1.Location = New System.Drawing.Point(324, 6)
    Me.Button1.Name = "Button1"
    Me.Button1.Size = New System.Drawing.Size(96, 27)
    Me.Button1.TabIndex = 0
    Me.Button1.Text = "OK"
    Me.Button1.UseVisualStyleBackColor = True
    '
    'Button2
    '
    Me.Button2.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.Button2.Location = New System.Drawing.Point(432, 6)
    Me.Button2.Name = "Button2"
    Me.Button2.Size = New System.Drawing.Size(96, 27)
    Me.Button2.TabIndex = 1
    Me.Button2.Text = "Cancel"
    Me.Button2.UseVisualStyleBackColor = True
    '
    'lblAvailable
    '
    Me.lblAvailable.AutoSize = True
    Me.lblAvailable.Location = New System.Drawing.Point(3, 0)
    Me.lblAvailable.Name = "lblAvailable"
    Me.lblAvailable.Size = New System.Drawing.Size(32, 39)
    Me.lblAvailable.TabIndex = 4
    Me.lblAvailable.Text = "Available Items"
    '
    'TableLayoutPanel
    '
    Me.TableLayoutPanel.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.TableLayoutPanel.ColumnCount = 3
    Me.TableLayoutPanel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
    Me.TableLayoutPanel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 120.0!))
    Me.TableLayoutPanel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
    Me.TableLayoutPanel.Controls.Add(Me.lblAvailable, 0, 0)
    Me.TableLayoutPanel.Location = New System.Drawing.Point(0, 0)
    Me.TableLayoutPanel.Name = "TableLayoutPanel"
    Me.TableLayoutPanel.RowCount = 1
    Me.TableLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
    Me.TableLayoutPanel.Size = New System.Drawing.Size(200, 100)
    Me.TableLayoutPanel.TabIndex = 0
    '
    'lstAvailable
    '
    Me.lstAvailable.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.lstAvailable.FormattingEnabled = True
    Me.lstAvailable.Location = New System.Drawing.Point(3, 29)
    Me.lstAvailable.Name = "lstAvailable"
    Me.lstAvailable.Size = New System.Drawing.Size(198, 69)
    Me.lstAvailable.TabIndex = 0
    '
    'SplitHeader
    '
    Me.SplitHeader.Dock = System.Windows.Forms.DockStyle.Fill
    Me.SplitHeader.FixedPanel = System.Windows.Forms.FixedPanel.Panel2
    Me.SplitHeader.Location = New System.Drawing.Point(0, 0)
    Me.SplitHeader.Margin = New System.Windows.Forms.Padding(0)
    Me.SplitHeader.Name = "SplitHeader"
    Me.SplitHeader.Orientation = System.Windows.Forms.Orientation.Horizontal
    '
    'SplitHeader.Panel1
    '
    Me.SplitHeader.Panel1.Controls.Add(Me.TableLayoutPanel1)
    '
    'SplitHeader.Panel2
    '
    Me.SplitHeader.Panel2.Controls.Add(Me.pnl)
    Me.SplitHeader.Panel2.Controls.Add(Me.bpl1)
    Me.SplitHeader.Size = New System.Drawing.Size(572, 520)
    Me.SplitHeader.SplitterDistance = 381
    Me.SplitHeader.SplitterWidth = 1
    Me.SplitHeader.TabIndex = 11
    '
    'TableLayoutPanel1
    '
    Me.TableLayoutPanel1.ColumnCount = 3
    Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
    Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 120.0!))
    Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
    Me.TableLayoutPanel1.Controls.Add(Me.lvwSelected, 0, 0)
    Me.TableLayoutPanel1.Controls.Add(Me.Panel1, 0, 0)
    Me.TableLayoutPanel1.Controls.Add(Me.lvwAvailable, 0, 0)
    Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
    Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
    Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
    Me.TableLayoutPanel1.RowCount = 1
    Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
    Me.TableLayoutPanel1.Size = New System.Drawing.Size(572, 381)
    Me.TableLayoutPanel1.TabIndex = 11
    '
    'lvwSelected
    '
    Me.lvwSelected.AllowDrop = True
    Me.lvwSelected.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.lvwSelected.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader3, Me.ColumnHeader4})
    Me.lvwSelected.LabelEdit = True
    Me.lvwSelected.Location = New System.Drawing.Point(349, 3)
    Me.lvwSelected.MultiSelect = False
    Me.lvwSelected.Name = "lvwSelected"
    Me.lvwSelected.Size = New System.Drawing.Size(220, 375)
    Me.lvwSelected.TabIndex = 11
    Me.lvwSelected.UseCompatibleStateImageBehavior = False
    Me.lvwSelected.View = System.Windows.Forms.View.List
    '
    'Panel1
    '
    Me.Panel1.Controls.Add(Me.bpl2)
    Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
    Me.Panel1.Location = New System.Drawing.Point(229, 3)
    Me.Panel1.Name = "Panel1"
    Me.Panel1.Size = New System.Drawing.Size(114, 375)
    Me.Panel1.TabIndex = 10
    '
    'bpl2
    '
    Me.bpl2.Controls.Add(Me.btnAdd)
    Me.bpl2.Controls.Add(Me.btnRemove)
    Me.bpl2.Controls.Add(Me.btnAddAll)
    Me.bpl2.Controls.Add(Me.btnRemoveAll)
    Me.bpl2.Dock = System.Windows.Forms.DockStyle.Fill
    Me.bpl2.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsCenter
    Me.bpl2.Location = New System.Drawing.Point(0, 0)
    Me.bpl2.Name = "bpl2"
    Me.bpl2.Size = New System.Drawing.Size(112, 375)
    Me.bpl2.TabIndex = 9
    '
    'btnAdd
    '
    Me.btnAdd.Location = New System.Drawing.Point(8, 10)
    Me.btnAdd.Name = "btnAdd"
    Me.btnAdd.Size = New System.Drawing.Size(96, 27)
    Me.btnAdd.TabIndex = 1
    Me.btnAdd.Text = "&Add >"
    Me.btnAdd.UseVisualStyleBackColor = True
    '
    'btnRemove
    '
    Me.btnRemove.Location = New System.Drawing.Point(8, 47)
    Me.btnRemove.Name = "btnRemove"
    Me.btnRemove.Size = New System.Drawing.Size(96, 27)
    Me.btnRemove.TabIndex = 2
    Me.btnRemove.Text = "< &Remove"
    Me.btnRemove.UseVisualStyleBackColor = True
    '
    'btnAddAll
    '
    Me.btnAddAll.Location = New System.Drawing.Point(8, 84)
    Me.btnAddAll.Name = "btnAddAll"
    Me.btnAddAll.Size = New System.Drawing.Size(96, 27)
    Me.btnAddAll.TabIndex = 4
    Me.btnAddAll.Text = "Add All &>>"
    Me.btnAddAll.UseVisualStyleBackColor = True
    '
    'btnRemoveAll
    '
    Me.btnRemoveAll.Location = New System.Drawing.Point(8, 121)
    Me.btnRemoveAll.Name = "btnRemoveAll"
    Me.btnRemoveAll.Size = New System.Drawing.Size(96, 27)
    Me.btnRemoveAll.TabIndex = 3
    Me.btnRemoveAll.Text = "&<<Remove All"
    Me.btnRemoveAll.UseVisualStyleBackColor = True
    '
    'lvwAvailable
    '
    Me.lvwAvailable.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.lvwAvailable.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2})
    Me.lvwAvailable.LabelEdit = True
    Me.lvwAvailable.Location = New System.Drawing.Point(3, 3)
    Me.lvwAvailable.MultiSelect = False
    Me.lvwAvailable.Name = "lvwAvailable"
    Me.lvwAvailable.Size = New System.Drawing.Size(220, 375)
    Me.lvwAvailable.TabIndex = 1
    Me.lvwAvailable.UseCompatibleStateImageBehavior = False
    Me.lvwAvailable.View = System.Windows.Forms.View.List
    '
    'pnl
    '
    Me.pnl.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
    Me.pnl.Controls.Add(Me.chkLabelsBelow)
    Me.pnl.Controls.Add(Me.txtLabel)
    Me.pnl.Controls.Add(Me.txtToolTip)
    Me.pnl.Controls.Add(Me.lblLabel)
    Me.pnl.Controls.Add(Me.lblToolTip)
    Me.pnl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.pnl.Location = New System.Drawing.Point(0, -5)
    Me.pnl.Name = "pnl"
    Me.pnl.Size = New System.Drawing.Size(572, 104)
    Me.pnl.TabIndex = 8
    '
    'chkLabelsBelow
    '
    Me.chkLabelsBelow.AutoSize = True
    Me.chkLabelsBelow.Location = New System.Drawing.Point(15, 64)
    Me.chkLabelsBelow.Name = "chkLabelsBelow"
    Me.chkLabelsBelow.Size = New System.Drawing.Size(137, 17)
    Me.chkLabelsBelow.TabIndex = 4
    Me.chkLabelsBelow.Text = "Show labels below icon"
    Me.chkLabelsBelow.UseVisualStyleBackColor = True
    '
    'txtLabel
    '
    Me.txtLabel.Location = New System.Drawing.Point(100, 34)
    Me.txtLabel.Name = "txtLabel"
    Me.txtLabel.Size = New System.Drawing.Size(291, 20)
    Me.txtLabel.TabIndex = 3
    '
    'txtToolTip
    '
    Me.txtToolTip.Location = New System.Drawing.Point(99, 7)
    Me.txtToolTip.Name = "txtToolTip"
    Me.txtToolTip.Size = New System.Drawing.Size(291, 20)
    Me.txtToolTip.TabIndex = 2
    '
    'lblLabel
    '
    Me.lblLabel.AutoSize = True
    Me.lblLabel.Location = New System.Drawing.Point(11, 37)
    Me.lblLabel.Name = "lblLabel"
    Me.lblLabel.Size = New System.Drawing.Size(36, 13)
    Me.lblLabel.TabIndex = 1
    Me.lblLabel.Text = "Label:"
    '
    'lblToolTip
    '
    Me.lblToolTip.AutoSize = True
    Me.lblToolTip.Location = New System.Drawing.Point(12, 10)
    Me.lblToolTip.Name = "lblToolTip"
    Me.lblToolTip.Size = New System.Drawing.Size(46, 13)
    Me.lblToolTip.TabIndex = 0
    Me.lblToolTip.Text = "ToolTip:"
    '
    'bpl1
    '
    Me.bpl1.Controls.Add(Me.cmdOk)
    Me.bpl1.Controls.Add(Me.cmdApply)
    Me.bpl1.Controls.Add(Me.cmdDefault)
    Me.bpl1.Controls.Add(Me.cmdCancel)
    Me.bpl1.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bpl1.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bpl1.Location = New System.Drawing.Point(0, 99)
    Me.bpl1.Name = "bpl1"
    Me.bpl1.Size = New System.Drawing.Size(572, 39)
    Me.bpl1.TabIndex = 7
    '
    'cmdOk
    '
    Me.cmdOk.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK
    Me.cmdOk.Location = New System.Drawing.Point(71, 6)
    Me.cmdOk.Name = "cmdOk"
    Me.cmdOk.Size = New System.Drawing.Size(96, 27)
    Me.cmdOk.TabIndex = 10
    Me.cmdOk.Text = "&OK"
    '
    'cmdApply
    '
    Me.cmdApply.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdApply.Location = New System.Drawing.Point(182, 6)
    Me.cmdApply.Name = "cmdApply"
    Me.cmdApply.Size = New System.Drawing.Size(96, 27)
    Me.cmdApply.TabIndex = 8
    Me.cmdApply.Text = "&Apply"
    '
    'cmdDefault
    '
    Me.cmdDefault.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdDefault.Location = New System.Drawing.Point(293, 6)
    Me.cmdDefault.Name = "cmdDefault"
    Me.cmdDefault.Size = New System.Drawing.Size(96, 27)
    Me.cmdDefault.TabIndex = 9
    Me.cmdDefault.Text = "&Defaults"
    '
    'cmdCancel
    '
    Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdCancel.Location = New System.Drawing.Point(404, 6)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(96, 27)
    Me.cmdCancel.TabIndex = 7
    Me.cmdCancel.Text = "Cancel"
    '
    'frmCustomiseToolBar
    '
    Me.AcceptButton = Me.cmdOk
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.cmdCancel
    Me.ClientSize = New System.Drawing.Size(572, 520)
    Me.Controls.Add(Me.SplitHeader)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmCustomiseToolBar"
    Me.ShowInTaskbar = False
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.Text = "Customise ToolBar"
    Me.ButtonPanel1.ResumeLayout(False)
    Me.TableLayoutPanel.ResumeLayout(False)
    Me.TableLayoutPanel.PerformLayout()
    Me.SplitHeader.Panel1.ResumeLayout(False)
    Me.SplitHeader.Panel2.ResumeLayout(False)
    CType(Me.SplitHeader, System.ComponentModel.ISupportInitialize).EndInit()
    Me.SplitHeader.ResumeLayout(False)
    Me.TableLayoutPanel1.ResumeLayout(False)
    Me.Panel1.ResumeLayout(False)
    Me.bpl2.ResumeLayout(False)
    Me.pnl.ResumeLayout(False)
    Me.pnl.PerformLayout()
    Me.bpl1.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub
  'Friend WithEvents bpl2 As CDBNETCL.ButtonPanel
  Friend WithEvents cmdAdd As System.Windows.Forms.Button
  Friend WithEvents cmdRemove As System.Windows.Forms.Button
  Friend WithEvents cmdAddAll As System.Windows.Forms.Button
  Friend WithEvents cmdRemoveAll As System.Windows.Forms.Button
  Friend WithEvents ButtonPanel1 As CDBNETCL.ButtonPanel
  Friend WithEvents cmdDelete As System.Windows.Forms.Button
  Friend WithEvents cmdNew As System.Windows.Forms.Button
  Friend WithEvents cmdEdit As System.Windows.Forms.Button
  Friend WithEvents Button1 As System.Windows.Forms.Button
  Friend WithEvents Button2 As System.Windows.Forms.Button
  Friend WithEvents lblAvailable As System.Windows.Forms.Label
  Friend WithEvents TableLayoutPanel As System.Windows.Forms.TableLayoutPanel
  Friend WithEvents lstAvailable As System.Windows.Forms.ListBox
  Friend WithEvents SplitHeader As System.Windows.Forms.SplitContainer
  Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
  Friend WithEvents lvwSelected As System.Windows.Forms.ListView
  Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
  Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
  Friend WithEvents Panel1 As System.Windows.Forms.Panel
  Friend WithEvents bpl2 As CDBNETCL.ButtonPanel
  Friend WithEvents btnAdd As System.Windows.Forms.Button
  Friend WithEvents btnRemoveAll As System.Windows.Forms.Button
  Friend WithEvents btnRemove As System.Windows.Forms.Button
  Friend WithEvents btnAddAll As System.Windows.Forms.Button
  Friend WithEvents lvwAvailable As System.Windows.Forms.ListView
  Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
  Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
  Friend WithEvents pnl As System.Windows.Forms.Panel
  Friend WithEvents chkLabelsBelow As System.Windows.Forms.CheckBox
  Friend WithEvents txtLabel As System.Windows.Forms.TextBox
  Friend WithEvents txtToolTip As System.Windows.Forms.TextBox
  Friend WithEvents lblLabel As System.Windows.Forms.Label
  Friend WithEvents lblToolTip As System.Windows.Forms.Label
  Friend WithEvents bpl1 As CDBNETCL.ButtonPanel
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents cmdApply As System.Windows.Forms.Button
  Friend WithEvents cmdDefault As System.Windows.Forms.Button
  Friend WithEvents cmdOk As System.Windows.Forms.Button
End Class
