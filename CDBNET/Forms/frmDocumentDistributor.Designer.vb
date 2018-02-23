<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDocumentDistributor
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDocumentDistributor))
    Me.splt = New System.Windows.Forms.SplitContainer()
    Me.tab = New CDBNETCL.TabControl()
    Me.tbpType = New System.Windows.Forms.TabPage()
    Me.pnlType = New System.Windows.Forms.Panel()
    Me.cboType = New System.Windows.Forms.ComboBox()
    Me.tbpPostPoint = New System.Windows.Forms.TabPage()
    Me.pnlPostPoint = New System.Windows.Forms.Panel()
    Me.cboPostPoint = New System.Windows.Forms.ComboBox()
    Me.tbpRecipient = New System.Windows.Forms.TabPage()
    Me.pnlRecipient = New System.Windows.Forms.Panel()
    Me.cboRecipient = New System.Windows.Forms.ComboBox()
    Me.grpBox = New System.Windows.Forms.GroupBox()
    Me.spltImageViewer = New System.Windows.Forms.SplitContainer()
    Me.pnlPictureBox = New System.Windows.Forms.Panel()
    Me.pctDocument = New System.Windows.Forms.PictureBox()
    Me.ZoomSlider = New System.Windows.Forms.TrackBar()
    Me.spltZoom = New System.Windows.Forms.SplitContainer()
    Me.cboZoom = New System.Windows.Forms.ComboBox()
    Me.bpl = New CDBNETCL.ButtonPanel()
    Me.cmdRotate = New System.Windows.Forms.Button()
    Me.cmdUpdate = New System.Windows.Forms.Button()
    Me.cmdDelete = New System.Windows.Forms.Button()
    Me.cmdClose = New System.Windows.Forms.Button()
    Me.pnlFinder = New System.Windows.Forms.Panel()
    Me.lblAddress = New System.Windows.Forms.Label()
    Me.lblSender = New System.Windows.Forms.Label()
    Me.cmdContactFinder = New System.Windows.Forms.Button()
    Me.cmdOrganisationFinder = New System.Windows.Forms.Button()
    Me.cboAddresses = New System.Windows.Forms.ComboBox()
    Me.txtName = New System.Windows.Forms.TextBox()
    CType(Me.splt, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.splt.Panel1.SuspendLayout()
    Me.splt.Panel2.SuspendLayout()
    Me.splt.SuspendLayout()
    Me.tab.SuspendLayout()
    Me.tbpType.SuspendLayout()
    Me.tbpPostPoint.SuspendLayout()
    Me.tbpRecipient.SuspendLayout()
    Me.grpBox.SuspendLayout()
    CType(Me.spltImageViewer, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.spltImageViewer.Panel1.SuspendLayout()
    Me.spltImageViewer.Panel2.SuspendLayout()
    Me.spltImageViewer.SuspendLayout()
    Me.pnlPictureBox.SuspendLayout()
    CType(Me.pctDocument, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.ZoomSlider, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.spltZoom, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.spltZoom.Panel1.SuspendLayout()
    Me.spltZoom.Panel2.SuspendLayout()
    Me.spltZoom.SuspendLayout()
    Me.bpl.SuspendLayout()
    Me.pnlFinder.SuspendLayout()
    Me.SuspendLayout()
    '
    'splt
    '
    Me.splt.Dock = System.Windows.Forms.DockStyle.Fill
    Me.splt.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
    Me.splt.IsSplitterFixed = True
    Me.splt.Location = New System.Drawing.Point(0, 0)
    Me.splt.Name = "splt"
    '
    'splt.Panel1
    '
    Me.splt.Panel1.Controls.Add(Me.tab)
    '
    'splt.Panel2
    '
    Me.splt.Panel2.Controls.Add(Me.grpBox)
    Me.splt.Panel2.Controls.Add(Me.pnlFinder)
    Me.splt.Size = New System.Drawing.Size(855, 501)
    Me.splt.SplitterDistance = 255
    Me.splt.TabIndex = 0
    '
    'tab
    '
    Me.tab.Controls.Add(Me.tbpType)
    Me.tab.Controls.Add(Me.tbpPostPoint)
    Me.tab.Controls.Add(Me.tbpRecipient)
    Me.tab.Dock = System.Windows.Forms.DockStyle.Fill
    Me.tab.ItemSize = New System.Drawing.Size(74, 22)
    Me.tab.Location = New System.Drawing.Point(0, 0)
    Me.tab.Name = "tab"
    Me.tab.SelectedIndex = 0
    Me.tab.Size = New System.Drawing.Size(255, 501)
    Me.tab.SizeMode = System.Windows.Forms.TabSizeMode.Fixed
    Me.tab.TabIndex = 1
    '
    'tbpType
    '
    Me.tbpType.Controls.Add(Me.pnlType)
    Me.tbpType.Controls.Add(Me.cboType)
    Me.tbpType.Location = New System.Drawing.Point(4, 26)
    Me.tbpType.Name = "tbpType"
    Me.tbpType.Padding = New System.Windows.Forms.Padding(3)
    Me.tbpType.Size = New System.Drawing.Size(247, 471)
    Me.tbpType.TabIndex = 0
    Me.tbpType.Text = "Type"
    Me.tbpType.UseVisualStyleBackColor = True
    '
    'pnlType
    '
    Me.pnlType.AutoScroll = True
    Me.pnlType.Dock = System.Windows.Forms.DockStyle.Fill
    Me.pnlType.Location = New System.Drawing.Point(3, 24)
    Me.pnlType.Name = "pnlType"
    Me.pnlType.Size = New System.Drawing.Size(241, 444)
    Me.pnlType.TabIndex = 2
    '
    'cboType
    '
    Me.cboType.Dock = System.Windows.Forms.DockStyle.Top
    Me.cboType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cboType.FormattingEnabled = True
    Me.cboType.ImeMode = System.Windows.Forms.ImeMode.Off
    Me.cboType.Location = New System.Drawing.Point(3, 3)
    Me.cboType.Name = "cboType"
    Me.cboType.Size = New System.Drawing.Size(241, 21)
    Me.cboType.TabIndex = 1
    '
    'tbpPostPoint
    '
    Me.tbpPostPoint.Controls.Add(Me.pnlPostPoint)
    Me.tbpPostPoint.Controls.Add(Me.cboPostPoint)
    Me.tbpPostPoint.Location = New System.Drawing.Point(4, 26)
    Me.tbpPostPoint.Name = "tbpPostPoint"
    Me.tbpPostPoint.Padding = New System.Windows.Forms.Padding(3)
    Me.tbpPostPoint.Size = New System.Drawing.Size(247, 471)
    Me.tbpPostPoint.TabIndex = 1
    Me.tbpPostPoint.Text = "PostPoint"
    Me.tbpPostPoint.UseVisualStyleBackColor = True
    '
    'pnlPostPoint
    '
    Me.pnlPostPoint.AutoScroll = True
    Me.pnlPostPoint.Dock = System.Windows.Forms.DockStyle.Fill
    Me.pnlPostPoint.Location = New System.Drawing.Point(3, 24)
    Me.pnlPostPoint.Name = "pnlPostPoint"
    Me.pnlPostPoint.Size = New System.Drawing.Size(241, 444)
    Me.pnlPostPoint.TabIndex = 3
    '
    'cboPostPoint
    '
    Me.cboPostPoint.Dock = System.Windows.Forms.DockStyle.Top
    Me.cboPostPoint.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cboPostPoint.FormattingEnabled = True
    Me.cboPostPoint.Location = New System.Drawing.Point(3, 3)
    Me.cboPostPoint.Name = "cboPostPoint"
    Me.cboPostPoint.Size = New System.Drawing.Size(241, 21)
    Me.cboPostPoint.TabIndex = 2
    '
    'tbpRecipient
    '
    Me.tbpRecipient.Controls.Add(Me.pnlRecipient)
    Me.tbpRecipient.Controls.Add(Me.cboRecipient)
    Me.tbpRecipient.Location = New System.Drawing.Point(4, 26)
    Me.tbpRecipient.Name = "tbpRecipient"
    Me.tbpRecipient.Padding = New System.Windows.Forms.Padding(3)
    Me.tbpRecipient.Size = New System.Drawing.Size(247, 471)
    Me.tbpRecipient.TabIndex = 2
    Me.tbpRecipient.Text = "Recipient"
    Me.tbpRecipient.UseVisualStyleBackColor = True
    '
    'pnlRecipient
    '
    Me.pnlRecipient.AutoScroll = True
    Me.pnlRecipient.Dock = System.Windows.Forms.DockStyle.Fill
    Me.pnlRecipient.Location = New System.Drawing.Point(3, 24)
    Me.pnlRecipient.Name = "pnlRecipient"
    Me.pnlRecipient.Size = New System.Drawing.Size(241, 444)
    Me.pnlRecipient.TabIndex = 5
    '
    'cboRecipient
    '
    Me.cboRecipient.Dock = System.Windows.Forms.DockStyle.Top
    Me.cboRecipient.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cboRecipient.FormattingEnabled = True
    Me.cboRecipient.Location = New System.Drawing.Point(3, 3)
    Me.cboRecipient.Name = "cboRecipient"
    Me.cboRecipient.Size = New System.Drawing.Size(241, 21)
    Me.cboRecipient.TabIndex = 4
    '
    'grpBox
    '
    Me.grpBox.Controls.Add(Me.spltImageViewer)
    Me.grpBox.Dock = System.Windows.Forms.DockStyle.Fill
    Me.grpBox.Location = New System.Drawing.Point(0, 81)
    Me.grpBox.Name = "grpBox"
    Me.grpBox.Size = New System.Drawing.Size(596, 420)
    Me.grpBox.TabIndex = 6
    Me.grpBox.TabStop = False
    Me.grpBox.Text = "Document Number:"
    '
    'spltImageViewer
    '
    Me.spltImageViewer.Dock = System.Windows.Forms.DockStyle.Fill
    Me.spltImageViewer.FixedPanel = System.Windows.Forms.FixedPanel.Panel2
    Me.spltImageViewer.Location = New System.Drawing.Point(3, 16)
    Me.spltImageViewer.Name = "spltImageViewer"
    '
    'spltImageViewer.Panel1
    '
    Me.spltImageViewer.Panel1.Controls.Add(Me.pnlPictureBox)
    Me.spltImageViewer.Panel1.Controls.Add(Me.ZoomSlider)
    '
    'spltImageViewer.Panel2
    '
    Me.spltImageViewer.Panel2.Controls.Add(Me.spltZoom)
    Me.spltImageViewer.Panel2MinSize = 112
    Me.spltImageViewer.Size = New System.Drawing.Size(590, 401)
    Me.spltImageViewer.SplitterDistance = 474
    Me.spltImageViewer.TabIndex = 7
    '
    'pnlPictureBox
    '
    Me.pnlPictureBox.AutoScroll = True
    Me.pnlPictureBox.Controls.Add(Me.pctDocument)
    Me.pnlPictureBox.Dock = System.Windows.Forms.DockStyle.Fill
    Me.pnlPictureBox.Location = New System.Drawing.Point(0, 0)
    Me.pnlPictureBox.Name = "pnlPictureBox"
    Me.pnlPictureBox.Size = New System.Drawing.Size(474, 356)
    Me.pnlPictureBox.TabIndex = 5
    '
    'pctDocument
    '
    Me.pctDocument.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
    Me.pctDocument.Location = New System.Drawing.Point(3, 3)
    Me.pctDocument.Name = "pctDocument"
    Me.pctDocument.Size = New System.Drawing.Size(483, 316)
    Me.pctDocument.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
    Me.pctDocument.TabIndex = 2
    Me.pctDocument.TabStop = False
    '
    'ZoomSlider
    '
    Me.ZoomSlider.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.ZoomSlider.Location = New System.Drawing.Point(0, 356)
    Me.ZoomSlider.Name = "ZoomSlider"
    Me.ZoomSlider.Size = New System.Drawing.Size(474, 45)
    Me.ZoomSlider.TabIndex = 4
    '
    'spltZoom
    '
    Me.spltZoom.Dock = System.Windows.Forms.DockStyle.Fill
    Me.spltZoom.Location = New System.Drawing.Point(0, 0)
    Me.spltZoom.Name = "spltZoom"
    Me.spltZoom.Orientation = System.Windows.Forms.Orientation.Horizontal
    '
    'spltZoom.Panel1
    '
    Me.spltZoom.Panel1.Controls.Add(Me.cboZoom)
    '
    'spltZoom.Panel2
    '
    Me.spltZoom.Panel2.Controls.Add(Me.bpl)
    Me.spltZoom.Size = New System.Drawing.Size(112, 401)
    Me.spltZoom.SplitterDistance = 37
    Me.spltZoom.TabIndex = 0
    '
    'cboZoom
    '
    Me.cboZoom.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cboZoom.FormattingEnabled = True
    Me.cboZoom.Items.AddRange(New Object() {"10%", "20%", "30%", "40%", "50%", "60%", "70%", "80%", "90%", "100%", "150%", "200%", "250%", "300%", "350%", "400%", "450%", "500%"})
    Me.cboZoom.Location = New System.Drawing.Point(0, 10)
    Me.cboZoom.Name = "cboZoom"
    Me.cboZoom.Size = New System.Drawing.Size(107, 21)
    Me.cboZoom.TabIndex = 2
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdRotate)
    Me.bpl.Controls.Add(Me.cmdUpdate)
    Me.bpl.Controls.Add(Me.cmdDelete)
    Me.bpl.Controls.Add(Me.cmdClose)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Fill
    Me.bpl.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsCenter
    Me.bpl.Location = New System.Drawing.Point(0, 0)
    Me.bpl.Margin = New System.Windows.Forms.Padding(0)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(112, 360)
    Me.bpl.TabIndex = 0
    '
    'cmdRotate
    '
    Me.cmdRotate.Location = New System.Drawing.Point(8, 10)
    Me.cmdRotate.Name = "cmdRotate"
    Me.cmdRotate.Size = New System.Drawing.Size(96, 27)
    Me.cmdRotate.TabIndex = 5
    Me.cmdRotate.Text = "Rotate"
    Me.cmdRotate.UseVisualStyleBackColor = True
    '
    'cmdUpdate
    '
    Me.cmdUpdate.Location = New System.Drawing.Point(8, 47)
    Me.cmdUpdate.Name = "cmdUpdate"
    Me.cmdUpdate.Size = New System.Drawing.Size(96, 27)
    Me.cmdUpdate.TabIndex = 7
    Me.cmdUpdate.Text = "Update"
    Me.cmdUpdate.UseVisualStyleBackColor = True
    '
    'cmdDelete
    '
    Me.cmdDelete.Location = New System.Drawing.Point(8, 84)
    Me.cmdDelete.Name = "cmdDelete"
    Me.cmdDelete.Size = New System.Drawing.Size(96, 27)
    Me.cmdDelete.TabIndex = 8
    Me.cmdDelete.Text = "Delete"
    Me.cmdDelete.UseVisualStyleBackColor = True
    '
    'cmdClose
    '
    Me.cmdClose.Location = New System.Drawing.Point(8, 121)
    Me.cmdClose.Name = "cmdClose"
    Me.cmdClose.Size = New System.Drawing.Size(96, 27)
    Me.cmdClose.TabIndex = 6
    Me.cmdClose.Text = "Close"
    Me.cmdClose.UseVisualStyleBackColor = True
    '
    'pnlFinder
    '
    Me.pnlFinder.Controls.Add(Me.lblAddress)
    Me.pnlFinder.Controls.Add(Me.lblSender)
    Me.pnlFinder.Controls.Add(Me.cmdContactFinder)
    Me.pnlFinder.Controls.Add(Me.cmdOrganisationFinder)
    Me.pnlFinder.Controls.Add(Me.cboAddresses)
    Me.pnlFinder.Controls.Add(Me.txtName)
    Me.pnlFinder.Dock = System.Windows.Forms.DockStyle.Top
    Me.pnlFinder.Location = New System.Drawing.Point(0, 0)
    Me.pnlFinder.Name = "pnlFinder"
    Me.pnlFinder.Size = New System.Drawing.Size(596, 81)
    Me.pnlFinder.TabIndex = 8
    '
    'lblAddress
    '
    Me.lblAddress.AutoSize = True
    Me.lblAddress.Location = New System.Drawing.Point(17, 44)
    Me.lblAddress.Name = "lblAddress"
    Me.lblAddress.Size = New System.Drawing.Size(51, 13)
    Me.lblAddress.TabIndex = 7
    Me.lblAddress.Text = "Address :"
    '
    'lblSender
    '
    Me.lblSender.AutoSize = True
    Me.lblSender.Location = New System.Drawing.Point(17, 13)
    Me.lblSender.Name = "lblSender"
    Me.lblSender.Size = New System.Drawing.Size(47, 13)
    Me.lblSender.TabIndex = 6
    Me.lblSender.Text = "Sender :"
    '
    'cmdContactFinder
    '
    Me.cmdContactFinder.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdContactFinder.Location = New System.Drawing.Point(481, 8)
    Me.cmdContactFinder.Name = "cmdContactFinder"
    Me.cmdContactFinder.Size = New System.Drawing.Size(50, 44)
    Me.cmdContactFinder.TabIndex = 5
    Me.cmdContactFinder.UseVisualStyleBackColor = True
    '
    'cmdOrganisationFinder
    '
    Me.cmdOrganisationFinder.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdOrganisationFinder.Location = New System.Drawing.Point(534, 8)
    Me.cmdOrganisationFinder.Name = "cmdOrganisationFinder"
    Me.cmdOrganisationFinder.Size = New System.Drawing.Size(50, 44)
    Me.cmdOrganisationFinder.TabIndex = 4
    Me.cmdOrganisationFinder.UseVisualStyleBackColor = True
    '
    'cboAddresses
    '
    Me.cboAddresses.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cboAddresses.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cboAddresses.FormattingEnabled = True
    Me.cboAddresses.Location = New System.Drawing.Point(89, 37)
    Me.cboAddresses.Name = "cboAddresses"
    Me.cboAddresses.Size = New System.Drawing.Size(370, 21)
    Me.cboAddresses.TabIndex = 3
    '
    'txtName
    '
    Me.txtName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.txtName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
    Me.txtName.Location = New System.Drawing.Point(89, 8)
    Me.txtName.Name = "txtName"
    Me.txtName.ReadOnly = True
    Me.txtName.Size = New System.Drawing.Size(370, 20)
    Me.txtName.TabIndex = 2
    '
    'frmDocumentDistributor
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(855, 501)
    Me.Controls.Add(Me.splt)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmDocumentDistributor"
    Me.ShowInTaskbar = False
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    Me.Text = "Document Distributor"
    Me.splt.Panel1.ResumeLayout(False)
    Me.splt.Panel2.ResumeLayout(False)
    CType(Me.splt, System.ComponentModel.ISupportInitialize).EndInit()
    Me.splt.ResumeLayout(False)
    Me.tab.ResumeLayout(False)
    Me.tbpType.ResumeLayout(False)
    Me.tbpPostPoint.ResumeLayout(False)
    Me.tbpRecipient.ResumeLayout(False)
    Me.grpBox.ResumeLayout(False)
    Me.spltImageViewer.Panel1.ResumeLayout(False)
    Me.spltImageViewer.Panel1.PerformLayout()
    Me.spltImageViewer.Panel2.ResumeLayout(False)
    CType(Me.spltImageViewer, System.ComponentModel.ISupportInitialize).EndInit()
    Me.spltImageViewer.ResumeLayout(False)
    Me.pnlPictureBox.ResumeLayout(False)
    Me.pnlPictureBox.PerformLayout()
    CType(Me.pctDocument, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.ZoomSlider, System.ComponentModel.ISupportInitialize).EndInit()
    Me.spltZoom.Panel1.ResumeLayout(False)
    Me.spltZoom.Panel2.ResumeLayout(False)
    CType(Me.spltZoom, System.ComponentModel.ISupportInitialize).EndInit()
    Me.spltZoom.ResumeLayout(False)
    Me.bpl.ResumeLayout(False)
    Me.pnlFinder.ResumeLayout(False)
    Me.pnlFinder.PerformLayout()
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents splt As System.Windows.Forms.SplitContainer
  Friend WithEvents pctDocument As System.Windows.Forms.PictureBox
  Friend WithEvents ZoomSlider As System.Windows.Forms.TrackBar
  Friend WithEvents grpBox As System.Windows.Forms.GroupBox
  Friend WithEvents spltImageViewer As System.Windows.Forms.SplitContainer
  Friend WithEvents pnlFinder As System.Windows.Forms.Panel
  Friend WithEvents pnlPictureBox As System.Windows.Forms.Panel
  Friend WithEvents txtName As System.Windows.Forms.TextBox
  Friend WithEvents cboAddresses As System.Windows.Forms.ComboBox
  Friend WithEvents cmdOrganisationFinder As System.Windows.Forms.Button
  Friend WithEvents cmdContactFinder As System.Windows.Forms.Button
  Friend WithEvents lblAddress As System.Windows.Forms.Label
  Friend WithEvents lblSender As System.Windows.Forms.Label
  Friend WithEvents tab As CDBNETCL.TabControl
  Friend WithEvents tbpType As System.Windows.Forms.TabPage
  Friend WithEvents tbpPostPoint As System.Windows.Forms.TabPage
  Friend WithEvents tbpRecipient As System.Windows.Forms.TabPage
  Friend WithEvents pnlType As System.Windows.Forms.Panel
  Friend WithEvents cboType As System.Windows.Forms.ComboBox
  Friend WithEvents pnlPostPoint As System.Windows.Forms.Panel
  Friend WithEvents cboPostPoint As System.Windows.Forms.ComboBox
  Friend WithEvents pnlRecipient As System.Windows.Forms.Panel
  Friend WithEvents cboRecipient As System.Windows.Forms.ComboBox
  Friend WithEvents cboZoom As System.Windows.Forms.ComboBox
  Friend WithEvents bpl As CDBNETCL.ButtonPanel
  Friend WithEvents cmdRotate As System.Windows.Forms.Button
  Friend WithEvents cmdUpdate As System.Windows.Forms.Button
  Friend WithEvents cmdDelete As System.Windows.Forms.Button
  Friend WithEvents cmdClose As System.Windows.Forms.Button
  Friend WithEvents spltZoom As System.Windows.Forms.SplitContainer
End Class
