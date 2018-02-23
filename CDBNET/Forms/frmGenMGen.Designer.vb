<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmGenMGen
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmGenMGen))
    Me.Panel1 = New System.Windows.Forms.Panel()
    Me.ButtonPanel1 = New CDBNETCL.ButtonPanel()
    Me.cmdOK = New System.Windows.Forms.Button()
    Me.cmdCancel = New System.Windows.Forms.Button()
    Me.cmdRefine = New System.Windows.Forms.Button()
    Me.cmdReset = New System.Windows.Forms.Button()
    Me.cmdClear = New System.Windows.Forms.Button()
    Me.cmdSaveList = New System.Windows.Forms.Button()
    Me.cmdMerge = New System.Windows.Forms.Button()
    Me.cmdSaveCriteria = New System.Windows.Forms.Button()
    Me.cmdPrint = New System.Windows.Forms.Button()
    Me.cmdView = New System.Windows.Forms.Button()
    Me.cmdOmit = New System.Windows.Forms.Button()
    Me.cmdCount = New System.Windows.Forms.Button()
    Me.cmdReport = New System.Windows.Forms.Button()
    Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
    Me.pnlSOCancellation = New System.Windows.Forms.Panel()
    Me.eplSOCancellation = New CDBNETCL.EditPanel()
    Me.pnlGrid = New System.Windows.Forms.Panel()
    Me.drg = New CDBNETCL.DisplayGrid()
    Me.pnlEpl = New System.Windows.Forms.Panel()
    Me.eplSelectionTester = New CDBNETCL.EditPanel()
    Me.eplMisc = New CDBNETCL.EditPanel()
    Me.eplStandard = New CDBNETCL.EditPanel()
    Me.Panel3 = New System.Windows.Forms.Panel()
    Me.pnlAddress = New CDBNETCL.PanelEx()
    Me.lblOrganisation = New System.Windows.Forms.Label()
    Me.lblAddress = New System.Windows.Forms.Label()
    Me.lblAddressText = New System.Windows.Forms.Label()
    Me.lblOrganisationLabel = New System.Windows.Forms.Label()
    Me.cmdFindAddress = New System.Windows.Forms.Button()
    Me.Panel1.SuspendLayout()
    Me.ButtonPanel1.SuspendLayout()
    CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SplitContainer1.Panel1.SuspendLayout()
    Me.SplitContainer1.Panel2.SuspendLayout()
    Me.SplitContainer1.SuspendLayout()
    Me.pnlSOCancellation.SuspendLayout()
    Me.pnlGrid.SuspendLayout()
    Me.pnlEpl.SuspendLayout()
    Me.Panel3.SuspendLayout()
    Me.pnlAddress.SuspendLayout()
    Me.SuspendLayout()
    '
    'Panel1
    '
    Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
    Me.Panel1.Controls.Add(Me.ButtonPanel1)
    Me.Panel1.Dock = System.Windows.Forms.DockStyle.Right
    Me.Panel1.Location = New System.Drawing.Point(729, 0)
    Me.Panel1.Name = "Panel1"
    Me.Panel1.Size = New System.Drawing.Size(115, 631)
    Me.Panel1.TabIndex = 0
    '
    'ButtonPanel1
    '
    Me.ButtonPanel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
    Me.ButtonPanel1.Controls.Add(Me.cmdOK)
    Me.ButtonPanel1.Controls.Add(Me.cmdCancel)
    Me.ButtonPanel1.Controls.Add(Me.cmdRefine)
    Me.ButtonPanel1.Controls.Add(Me.cmdReset)
    Me.ButtonPanel1.Controls.Add(Me.cmdClear)
    Me.ButtonPanel1.Controls.Add(Me.cmdSaveList)
    Me.ButtonPanel1.Controls.Add(Me.cmdMerge)
    Me.ButtonPanel1.Controls.Add(Me.cmdSaveCriteria)
    Me.ButtonPanel1.Controls.Add(Me.cmdPrint)
    Me.ButtonPanel1.Controls.Add(Me.cmdView)
    Me.ButtonPanel1.Controls.Add(Me.cmdOmit)
    Me.ButtonPanel1.Controls.Add(Me.cmdCount)
    Me.ButtonPanel1.Controls.Add(Me.cmdReport)
    Me.ButtonPanel1.Dock = System.Windows.Forms.DockStyle.Fill
    Me.ButtonPanel1.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsCenter
    Me.ButtonPanel1.Location = New System.Drawing.Point(0, 0)
    Me.ButtonPanel1.Name = "ButtonPanel1"
    Me.ButtonPanel1.Size = New System.Drawing.Size(112, 629)
    Me.ButtonPanel1.TabIndex = 0
    '
    'cmdOK
    '
    Me.cmdOK.Location = New System.Drawing.Point(8, 10)
    Me.cmdOK.Name = "cmdOK"
    Me.cmdOK.Size = New System.Drawing.Size(96, 27)
    Me.cmdOK.TabIndex = 1
    Me.cmdOK.Text = "OK"
    Me.cmdOK.UseVisualStyleBackColor = True
    '
    'cmdCancel
    '
    Me.cmdCancel.Location = New System.Drawing.Point(8, 47)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(96, 27)
    Me.cmdCancel.TabIndex = 2
    Me.cmdCancel.Text = "Cancel"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'cmdRefine
    '
    Me.cmdRefine.Location = New System.Drawing.Point(8, 84)
    Me.cmdRefine.Name = "cmdRefine"
    Me.cmdRefine.Size = New System.Drawing.Size(96, 27)
    Me.cmdRefine.TabIndex = 3
    Me.cmdRefine.Text = "Re&fine"
    Me.cmdRefine.UseVisualStyleBackColor = True
    '
    'cmdReset
    '
    Me.cmdReset.Location = New System.Drawing.Point(8, 121)
    Me.cmdReset.Name = "cmdReset"
    Me.cmdReset.Size = New System.Drawing.Size(96, 27)
    Me.cmdReset.TabIndex = 4
    Me.cmdReset.Text = "&Reset"
    Me.cmdReset.UseVisualStyleBackColor = True
    '
    'cmdClear
    '
    Me.cmdClear.Location = New System.Drawing.Point(8, 158)
    Me.cmdClear.Name = "cmdClear"
    Me.cmdClear.Size = New System.Drawing.Size(96, 27)
    Me.cmdClear.TabIndex = 5
    Me.cmdClear.Text = "&Clear"
    Me.cmdClear.UseVisualStyleBackColor = True
    '
    'cmdSaveList
    '
    Me.cmdSaveList.Location = New System.Drawing.Point(8, 195)
    Me.cmdSaveList.Name = "cmdSaveList"
    Me.cmdSaveList.Size = New System.Drawing.Size(96, 27)
    Me.cmdSaveList.TabIndex = 6
    Me.cmdSaveList.Text = "&Save List"
    Me.cmdSaveList.UseVisualStyleBackColor = True
    '
    'cmdMerge
    '
    Me.cmdMerge.Location = New System.Drawing.Point(8, 232)
    Me.cmdMerge.Name = "cmdMerge"
    Me.cmdMerge.Size = New System.Drawing.Size(96, 27)
    Me.cmdMerge.TabIndex = 7
    Me.cmdMerge.Text = "&Add To List"
    Me.cmdMerge.UseVisualStyleBackColor = True
    '
    'cmdSaveCriteria
    '
    Me.cmdSaveCriteria.Location = New System.Drawing.Point(8, 269)
    Me.cmdSaveCriteria.Name = "cmdSaveCriteria"
    Me.cmdSaveCriteria.Size = New System.Drawing.Size(96, 27)
    Me.cmdSaveCriteria.TabIndex = 8
    Me.cmdSaveCriteria.Text = "Sa&ve Criteria"
    Me.cmdSaveCriteria.UseVisualStyleBackColor = True
    '
    'cmdPrint
    '
    Me.cmdPrint.Location = New System.Drawing.Point(8, 306)
    Me.cmdPrint.Name = "cmdPrint"
    Me.cmdPrint.Size = New System.Drawing.Size(96, 27)
    Me.cmdPrint.TabIndex = 9
    Me.cmdPrint.Text = "&Print"
    Me.cmdPrint.UseVisualStyleBackColor = True
    '
    'cmdView
    '
    Me.cmdView.Location = New System.Drawing.Point(8, 343)
    Me.cmdView.Name = "cmdView"
    Me.cmdView.Size = New System.Drawing.Size(96, 27)
    Me.cmdView.TabIndex = 10
    Me.cmdView.Text = "&View"
    Me.cmdView.UseVisualStyleBackColor = True
    '
    'cmdOmit
    '
    Me.cmdOmit.Location = New System.Drawing.Point(8, 380)
    Me.cmdOmit.Name = "cmdOmit"
    Me.cmdOmit.Size = New System.Drawing.Size(96, 27)
    Me.cmdOmit.TabIndex = 11
    Me.cmdOmit.Text = "&Omit"
    Me.cmdOmit.UseVisualStyleBackColor = True
    '
    'cmdCount
    '
    Me.cmdCount.Location = New System.Drawing.Point(8, 417)
    Me.cmdCount.Name = "cmdCount"
    Me.cmdCount.Size = New System.Drawing.Size(96, 27)
    Me.cmdCount.TabIndex = 13
    Me.cmdCount.Text = "Count"
    Me.cmdCount.UseVisualStyleBackColor = True
    '
    'cmdReport
    '
    Me.cmdReport.Location = New System.Drawing.Point(8, 454)
    Me.cmdReport.Name = "cmdReport"
    Me.cmdReport.Size = New System.Drawing.Size(96, 27)
    Me.cmdReport.TabIndex = 12
    Me.cmdReport.Text = "&Report"
    Me.cmdReport.UseVisualStyleBackColor = True
    Me.cmdReport.Visible = False
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
    Me.SplitContainer1.Panel1.Controls.Add(Me.pnlSOCancellation)
    Me.SplitContainer1.Panel1.Controls.Add(Me.pnlGrid)
    Me.SplitContainer1.Panel1.Controls.Add(Me.pnlEpl)
    '
    'SplitContainer1.Panel2
    '
    Me.SplitContainer1.Panel2.Controls.Add(Me.Panel3)
    Me.SplitContainer1.Size = New System.Drawing.Size(729, 631)
    Me.SplitContainer1.SplitterDistance = 499
    Me.SplitContainer1.TabIndex = 2
    '
    'pnlSOCancellation
    '
    Me.pnlSOCancellation.Controls.Add(Me.eplSOCancellation)
    Me.pnlSOCancellation.Dock = System.Windows.Forms.DockStyle.Top
    Me.pnlSOCancellation.Location = New System.Drawing.Point(0, 393)
    Me.pnlSOCancellation.Name = "pnlSOCancellation"
    Me.pnlSOCancellation.Size = New System.Drawing.Size(729, 393)
    Me.pnlSOCancellation.TabIndex = 9
    '
    'eplSOCancellation
    '
    Me.eplSOCancellation.AddressChanged = False
    Me.eplSOCancellation.BackColor = System.Drawing.Color.Transparent
    Me.eplSOCancellation.DataChanged = False
    Me.eplSOCancellation.DefaultSaveFolder = ""
    Me.eplSOCancellation.Location = New System.Drawing.Point(0, 0)
    Me.eplSOCancellation.Name = "eplSOCancellation"
    Me.eplSOCancellation.Recipients = Nothing
    Me.eplSOCancellation.Size = New System.Drawing.Size(729, 0)
    Me.eplSOCancellation.SuppressDrawing = False
    Me.eplSOCancellation.TabIndex = 8
    Me.eplSOCancellation.TabSelectedIndex = 0
    Me.eplSOCancellation.Visible = False
    '
    'pnlGrid
    '
    Me.pnlGrid.Controls.Add(Me.drg)
    Me.pnlGrid.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.pnlGrid.Location = New System.Drawing.Point(0, 399)
    Me.pnlGrid.Name = "pnlGrid"
    Me.pnlGrid.Size = New System.Drawing.Size(729, 100)
    Me.pnlGrid.TabIndex = 1
    '
    'drg
    '
    Me.drg.AccessibleName = "Display Grid"
    Me.drg.ActiveColumn = 0
    Me.drg.AllowColumnResize = True
    Me.drg.AllowSorting = True
    Me.drg.AutoSetHeight = False
    Me.drg.AutoSetRowHeight = False
    Me.drg.DisplayTitle = Nothing
    Me.drg.Dock = System.Windows.Forms.DockStyle.Fill
    Me.drg.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!)
    Me.drg.Location = New System.Drawing.Point(0, 0)
    Me.drg.MaintenanceDesc = Nothing
    Me.drg.MaxGridRows = 8
    Me.drg.MultipleSelect = False
    Me.drg.Name = "drg"
    Me.drg.RowCount = 10
    Me.drg.ShowIfEmpty = False
    Me.drg.Size = New System.Drawing.Size(729, 100)
    Me.drg.SuppressHyperLinkFormat = False
    Me.drg.TabIndex = 0
    '
    'pnlEpl
    '
    Me.pnlEpl.AutoScroll = True
    Me.pnlEpl.Controls.Add(Me.eplSelectionTester)
    Me.pnlEpl.Controls.Add(Me.eplMisc)
    Me.pnlEpl.Controls.Add(Me.eplStandard)
    Me.pnlEpl.Dock = System.Windows.Forms.DockStyle.Top
    Me.pnlEpl.Location = New System.Drawing.Point(0, 0)
    Me.pnlEpl.Name = "pnlEpl"
    Me.pnlEpl.Size = New System.Drawing.Size(729, 393)
    Me.pnlEpl.TabIndex = 0
    '
    'eplSelectionTester
    '
    Me.eplSelectionTester.AddressChanged = False
    Me.eplSelectionTester.BackColor = System.Drawing.Color.Transparent
    Me.eplSelectionTester.DataChanged = False
    Me.eplSelectionTester.DefaultSaveFolder = ""
    Me.eplSelectionTester.Location = New System.Drawing.Point(0, 0)
    Me.eplSelectionTester.Name = "eplSelectionTester"
    Me.eplSelectionTester.Recipients = Nothing
    Me.eplSelectionTester.Size = New System.Drawing.Size(729, 103)
    Me.eplSelectionTester.SuppressDrawing = False
    Me.eplSelectionTester.TabIndex = 7
    Me.eplSelectionTester.TabSelectedIndex = 0
    '
    'eplMisc
    '
    Me.eplMisc.AddressChanged = False
    Me.eplMisc.BackColor = System.Drawing.Color.Transparent
    Me.eplMisc.DataChanged = False
    Me.eplMisc.DefaultSaveFolder = ""
    Me.eplMisc.Location = New System.Drawing.Point(0, 253)
    Me.eplMisc.Name = "eplMisc"
    Me.eplMisc.Recipients = Nothing
    Me.eplMisc.Size = New System.Drawing.Size(729, 137)
    Me.eplMisc.SuppressDrawing = False
    Me.eplMisc.TabIndex = 6
    Me.eplMisc.TabSelectedIndex = 0
    '
    'eplStandard
    '
    Me.eplStandard.AddressChanged = False
    Me.eplStandard.BackColor = System.Drawing.Color.Transparent
    Me.eplStandard.DataChanged = False
    Me.eplStandard.DefaultSaveFolder = ""
    Me.eplStandard.Location = New System.Drawing.Point(0, 0)
    Me.eplStandard.Name = "eplStandard"
    Me.eplStandard.Recipients = Nothing
    Me.eplStandard.Size = New System.Drawing.Size(729, 254)
    Me.eplStandard.SuppressDrawing = False
    Me.eplStandard.TabIndex = 5
    Me.eplStandard.TabSelectedIndex = 0
    '
    'Panel3
    '
    Me.Panel3.Controls.Add(Me.pnlAddress)
    Me.Panel3.Location = New System.Drawing.Point(0, 0)
    Me.Panel3.Name = "Panel3"
    Me.Panel3.Size = New System.Drawing.Size(729, 100)
    Me.Panel3.TabIndex = 0
    '
    'pnlAddress
    '
    Me.pnlAddress.BackColor = System.Drawing.Color.Transparent
    Me.pnlAddress.Controls.Add(Me.lblOrganisation)
    Me.pnlAddress.Controls.Add(Me.lblAddress)
    Me.pnlAddress.Controls.Add(Me.lblAddressText)
    Me.pnlAddress.Controls.Add(Me.lblOrganisationLabel)
    Me.pnlAddress.Controls.Add(Me.cmdFindAddress)
    Me.pnlAddress.Dock = System.Windows.Forms.DockStyle.Top
    Me.pnlAddress.Location = New System.Drawing.Point(0, 0)
    Me.pnlAddress.Name = "pnlAddress"
    Me.pnlAddress.Size = New System.Drawing.Size(729, 130)
    Me.pnlAddress.TabIndex = 10
    '
    'lblOrganisation
    '
    Me.lblOrganisation.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
    Me.lblOrganisation.Location = New System.Drawing.Point(153, 5)
    Me.lblOrganisation.Name = "lblOrganisation"
    Me.lblOrganisation.Size = New System.Drawing.Size(338, 22)
    Me.lblOrganisation.TabIndex = 6
    '
    'lblAddress
    '
    Me.lblAddress.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
    Me.lblAddress.Location = New System.Drawing.Point(153, 40)
    Me.lblAddress.Name = "lblAddress"
    Me.lblAddress.Size = New System.Drawing.Size(338, 60)
    Me.lblAddress.TabIndex = 5
    '
    'lblAddressText
    '
    Me.lblAddressText.AutoSize = True
    Me.lblAddressText.Location = New System.Drawing.Point(14, 41)
    Me.lblAddressText.Name = "lblAddressText"
    Me.lblAddressText.Size = New System.Drawing.Size(48, 13)
    Me.lblAddressText.TabIndex = 4
    Me.lblAddressText.Text = "Address:"
    '
    'lblOrganisationLabel
    '
    Me.lblOrganisationLabel.AutoSize = True
    Me.lblOrganisationLabel.Location = New System.Drawing.Point(14, 7)
    Me.lblOrganisationLabel.Name = "lblOrganisationLabel"
    Me.lblOrganisationLabel.Size = New System.Drawing.Size(69, 13)
    Me.lblOrganisationLabel.TabIndex = 3
    Me.lblOrganisationLabel.Text = "Organisation:"
    '
    'cmdFindAddress
    '
    Me.cmdFindAddress.Location = New System.Drawing.Point(497, 41)
    Me.cmdFindAddress.Name = "cmdFindAddress"
    Me.cmdFindAddress.Size = New System.Drawing.Size(25, 23)
    Me.cmdFindAddress.TabIndex = 2
    Me.cmdFindAddress.Text = "?"
    Me.cmdFindAddress.UseVisualStyleBackColor = True
    '
    'frmGenMGen
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(844, 631)
    Me.Controls.Add(Me.SplitContainer1)
    Me.Controls.Add(Me.Panel1)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmGenMGen"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    Me.Text = "Selection Manager - Generate"
    Me.Panel1.ResumeLayout(False)
    Me.ButtonPanel1.ResumeLayout(False)
    Me.SplitContainer1.Panel1.ResumeLayout(False)
    Me.SplitContainer1.Panel2.ResumeLayout(False)
    CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
    Me.SplitContainer1.ResumeLayout(False)
    Me.pnlSOCancellation.ResumeLayout(False)
    Me.pnlGrid.ResumeLayout(False)
    Me.pnlEpl.ResumeLayout(False)
    Me.Panel3.ResumeLayout(False)
    Me.pnlAddress.ResumeLayout(False)
    Me.pnlAddress.PerformLayout()
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents cmdOK As System.Windows.Forms.Button
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents cmdRefine As System.Windows.Forms.Button
  Friend WithEvents cmdReset As System.Windows.Forms.Button
  Friend WithEvents ButtonPanel1 As CDBNETCL.ButtonPanel
  Friend WithEvents cmdMerge As System.Windows.Forms.Button
  Friend WithEvents cmdSaveList As System.Windows.Forms.Button
  Friend WithEvents cmdClear As System.Windows.Forms.Button
  Friend WithEvents cmdReport As System.Windows.Forms.Button
  Friend WithEvents cmdOmit As System.Windows.Forms.Button
  Friend WithEvents cmdView As System.Windows.Forms.Button
  Friend WithEvents cmdPrint As System.Windows.Forms.Button
  Friend WithEvents cmdSaveCriteria As System.Windows.Forms.Button
  Friend WithEvents Panel1 As System.Windows.Forms.Panel
  Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
  Friend WithEvents Panel3 As System.Windows.Forms.Panel
  Friend WithEvents pnlAddress As CDBNETCL.PanelEx
  Friend WithEvents lblOrganisation As System.Windows.Forms.Label
  Friend WithEvents lblAddress As System.Windows.Forms.Label
  Friend WithEvents lblAddressText As System.Windows.Forms.Label
  Friend WithEvents lblOrganisationLabel As System.Windows.Forms.Label
  Friend WithEvents cmdFindAddress As System.Windows.Forms.Button
  Friend WithEvents pnlGrid As System.Windows.Forms.Panel
  Friend WithEvents drg As CDBNETCL.DisplayGrid
  Friend WithEvents pnlEpl As System.Windows.Forms.Panel
  Friend WithEvents eplStandard As CDBNETCL.EditPanel
  Friend WithEvents eplMisc As CDBNETCL.EditPanel
  Friend WithEvents eplSelectionTester As CDBNETCL.EditPanel
  Friend WithEvents pnlSOCancellation As System.Windows.Forms.Panel
  Friend WithEvents eplSOCancellation As CDBNETCL.EditPanel
  Friend WithEvents cmdCount As System.Windows.Forms.Button
End Class
