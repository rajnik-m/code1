<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPreferences
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmPreferences))
    Me.tab = New System.Windows.Forms.TabControl()
    Me.tabNotification = New System.Windows.Forms.TabPage()
    Me.lblTaskPollingInterval = New System.Windows.Forms.Label()
    Me.txtTaskPollingInterval = New System.Windows.Forms.TextBox()
    Me.chkNotifyMeetings = New System.Windows.Forms.CheckBox()
    Me.lblPollingInterval = New System.Windows.Forms.Label()
    Me.txtPollingInterval = New System.Windows.Forms.TextBox()
    Me.chkNotifyDeadlines = New System.Windows.Forms.CheckBox()
    Me.chkNotifyDocuments = New System.Windows.Forms.CheckBox()
    Me.chkNotifyActions = New System.Windows.Forms.CheckBox()
    Me.tabDisplay = New System.Windows.Forms.TabPage()
    Me.lblSchemes = New System.Windows.Forms.Label()
    Me.cboSchemes = New System.Windows.Forms.ComboBox()
    Me.chkPlainEditPanel = New System.Windows.Forms.CheckBox()
    Me.cmdBackgroundImage = New System.Windows.Forms.Button()
    Me.lblBackgroundImageLayout = New System.Windows.Forms.Label()
    Me.lblBackgroundImage = New System.Windows.Forms.Label()
    Me.cboBackgroundImageLayout = New System.Windows.Forms.ComboBox()
    Me.txtBackgroundImage = New System.Windows.Forms.TextBox()
    Me.tabGeneral = New System.Windows.Forms.TabPage()
    Me.chkFinderResultsMsgBox = New System.Windows.Forms.CheckBox()
    Me.chkHideHistoricNetwork = New System.Windows.Forms.CheckBox()
    Me.chkErrorsAsMsgbox = New System.Windows.Forms.CheckBox()
    Me.chkDisplayDashboardAtLogin = New System.Windows.Forms.CheckBox()
    Me.chkTabIntoHeaderPanel = New System.Windows.Forms.CheckBox()
    Me.chkTabIntoDisplayPanel = New System.Windows.Forms.CheckBox()
    Me.lblWebServicesTimeout = New System.Windows.Forms.Label()
    Me.txtWebServicesTimeout = New System.Windows.Forms.TextBox()
    Me.lblHistoryDays = New System.Windows.Forms.Label()
    Me.txtHistoryDays = New System.Windows.Forms.TextBox()
    Me.tabConfirmation = New System.Windows.Forms.TabPage()
    Me.chkConfirmCancel = New System.Windows.Forms.CheckBox()
    Me.chkConfirmUpdate = New System.Windows.Forms.CheckBox()
    Me.chkConfirmInsert = New System.Windows.Forms.CheckBox()
    Me.chkConfirmDelete = New System.Windows.Forms.CheckBox()
    Me.cmd = New System.Windows.Forms.OpenFileDialog()
    Me.cdlg = New System.Windows.Forms.ColorDialog()
    Me.cmdOK = New System.Windows.Forms.Button()
    Me.cmdCancel = New System.Windows.Forms.Button()
    Me.bpl = New CDBNETCL.ButtonPanel()
    Me.cmdDesign = New System.Windows.Forms.Button()
    Me.cmdSaveAs = New System.Windows.Forms.Button()
    Me.cmdApply = New System.Windows.Forms.Button()
    Me.cmdDefaults = New System.Windows.Forms.Button()
    Me.tim = New System.Windows.Forms.Timer(Me.components)
    Me.tab.SuspendLayout()
    Me.tabNotification.SuspendLayout()
    Me.tabDisplay.SuspendLayout()
    Me.tabGeneral.SuspendLayout()
    Me.tabConfirmation.SuspendLayout()
    Me.bpl.SuspendLayout()
    Me.SuspendLayout()
    '
    'tab
    '
    Me.tab.Controls.Add(Me.tabNotification)
    Me.tab.Controls.Add(Me.tabDisplay)
    Me.tab.Controls.Add(Me.tabGeneral)
    Me.tab.Controls.Add(Me.tabConfirmation)
    Me.tab.Dock = System.Windows.Forms.DockStyle.Fill
    Me.tab.Location = New System.Drawing.Point(0, 0)
    Me.tab.Name = "tab"
    Me.tab.SelectedIndex = 0
    Me.tab.Size = New System.Drawing.Size(802, 329)
    Me.tab.TabIndex = 0
    '
    'tabNotification
    '
    Me.tabNotification.Controls.Add(Me.lblTaskPollingInterval)
    Me.tabNotification.Controls.Add(Me.txtTaskPollingInterval)
    Me.tabNotification.Controls.Add(Me.chkNotifyMeetings)
    Me.tabNotification.Controls.Add(Me.lblPollingInterval)
    Me.tabNotification.Controls.Add(Me.txtPollingInterval)
    Me.tabNotification.Controls.Add(Me.chkNotifyDeadlines)
    Me.tabNotification.Controls.Add(Me.chkNotifyDocuments)
    Me.tabNotification.Controls.Add(Me.chkNotifyActions)
    Me.tabNotification.Location = New System.Drawing.Point(4, 22)
    Me.tabNotification.Name = "tabNotification"
    Me.tabNotification.Padding = New System.Windows.Forms.Padding(3)
    Me.tabNotification.Size = New System.Drawing.Size(794, 303)
    Me.tabNotification.TabIndex = 0
    Me.tabNotification.Text = "Notification"
    Me.tabNotification.UseVisualStyleBackColor = True
    '
    'lblTaskPollingInterval
    '
    Me.lblTaskPollingInterval.AutoSize = True
    Me.lblTaskPollingInterval.Location = New System.Drawing.Point(16, 193)
    Me.lblTaskPollingInterval.Name = "lblTaskPollingInterval"
    Me.lblTaskPollingInterval.Size = New System.Drawing.Size(134, 13)
    Me.lblTaskPollingInterval.TabIndex = 6
    Me.lblTaskPollingInterval.Text = "Task Polling Interval (secs)"
    '
    'txtTaskPollingInterval
    '
    Me.txtTaskPollingInterval.Location = New System.Drawing.Point(201, 190)
    Me.txtTaskPollingInterval.Name = "txtTaskPollingInterval"
    Me.txtTaskPollingInterval.Size = New System.Drawing.Size(47, 20)
    Me.txtTaskPollingInterval.TabIndex = 7
    '
    'chkNotifyMeetings
    '
    Me.chkNotifyMeetings.AutoSize = True
    Me.chkNotifyMeetings.Location = New System.Drawing.Point(19, 126)
    Me.chkNotifyMeetings.Name = "chkNotifyMeetings"
    Me.chkNotifyMeetings.Size = New System.Drawing.Size(99, 17)
    Me.chkNotifyMeetings.TabIndex = 5
    Me.chkNotifyMeetings.Text = "Notify Meetings"
    Me.chkNotifyMeetings.UseVisualStyleBackColor = True
    '
    'lblPollingInterval
    '
    Me.lblPollingInterval.AutoSize = True
    Me.lblPollingInterval.Location = New System.Drawing.Point(16, 161)
    Me.lblPollingInterval.Name = "lblPollingInterval"
    Me.lblPollingInterval.Size = New System.Drawing.Size(106, 13)
    Me.lblPollingInterval.TabIndex = 3
    Me.lblPollingInterval.Text = "Polling Interval (mins)"
    '
    'txtPollingInterval
    '
    Me.txtPollingInterval.Location = New System.Drawing.Point(201, 158)
    Me.txtPollingInterval.Name = "txtPollingInterval"
    Me.txtPollingInterval.Size = New System.Drawing.Size(47, 20)
    Me.txtPollingInterval.TabIndex = 4
    '
    'chkNotifyDeadlines
    '
    Me.chkNotifyDeadlines.AutoSize = True
    Me.chkNotifyDeadlines.Location = New System.Drawing.Point(19, 88)
    Me.chkNotifyDeadlines.Name = "chkNotifyDeadlines"
    Me.chkNotifyDeadlines.Size = New System.Drawing.Size(160, 17)
    Me.chkNotifyDeadlines.TabIndex = 2
    Me.chkNotifyDeadlines.Text = "Notify Actions Past &Deadline"
    Me.chkNotifyDeadlines.UseVisualStyleBackColor = True
    '
    'chkNotifyDocuments
    '
    Me.chkNotifyDocuments.AutoSize = True
    Me.chkNotifyDocuments.Location = New System.Drawing.Point(19, 51)
    Me.chkNotifyDocuments.Name = "chkNotifyDocuments"
    Me.chkNotifyDocuments.Size = New System.Drawing.Size(159, 17)
    Me.chkNotifyDocuments.TabIndex = 1
    Me.chkNotifyDocuments.Text = "Notify &Documents Received"
    Me.chkNotifyDocuments.UseVisualStyleBackColor = True
    '
    'chkNotifyActions
    '
    Me.chkNotifyActions.AutoSize = True
    Me.chkNotifyActions.Location = New System.Drawing.Point(19, 15)
    Me.chkNotifyActions.Name = "chkNotifyActions"
    Me.chkNotifyActions.Size = New System.Drawing.Size(91, 17)
    Me.chkNotifyActions.TabIndex = 0
    Me.chkNotifyActions.Text = "Notify &Actions"
    Me.chkNotifyActions.UseVisualStyleBackColor = True
    '
    'tabDisplay
    '
    Me.tabDisplay.Controls.Add(Me.lblSchemes)
    Me.tabDisplay.Controls.Add(Me.cboSchemes)
    Me.tabDisplay.Controls.Add(Me.chkPlainEditPanel)
    Me.tabDisplay.Controls.Add(Me.cmdBackgroundImage)
    Me.tabDisplay.Controls.Add(Me.lblBackgroundImageLayout)
    Me.tabDisplay.Controls.Add(Me.lblBackgroundImage)
    Me.tabDisplay.Controls.Add(Me.cboBackgroundImageLayout)
    Me.tabDisplay.Controls.Add(Me.txtBackgroundImage)
    Me.tabDisplay.Location = New System.Drawing.Point(4, 22)
    Me.tabDisplay.Name = "tabDisplay"
    Me.tabDisplay.Padding = New System.Windows.Forms.Padding(3)
    Me.tabDisplay.Size = New System.Drawing.Size(794, 303)
    Me.tabDisplay.TabIndex = 1
    Me.tabDisplay.Text = "Display"
    Me.tabDisplay.UseVisualStyleBackColor = True
    '
    'lblSchemes
    '
    Me.lblSchemes.AutoSize = True
    Me.lblSchemes.Location = New System.Drawing.Point(8, 157)
    Me.lblSchemes.Name = "lblSchemes"
    Me.lblSchemes.Size = New System.Drawing.Size(51, 13)
    Me.lblSchemes.TabIndex = 6
    Me.lblSchemes.Text = "Schemes"
    '
    'cboSchemes
    '
    Me.cboSchemes.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cboSchemes.FormattingEnabled = True
    Me.cboSchemes.Location = New System.Drawing.Point(127, 154)
    Me.cboSchemes.Name = "cboSchemes"
    Me.cboSchemes.Size = New System.Drawing.Size(229, 21)
    Me.cboSchemes.TabIndex = 7
    '
    'chkPlainEditPanel
    '
    Me.chkPlainEditPanel.AutoSize = True
    Me.chkPlainEditPanel.Location = New System.Drawing.Point(11, 118)
    Me.chkPlainEditPanel.Name = "chkPlainEditPanel"
    Me.chkPlainEditPanel.Size = New System.Drawing.Size(219, 17)
    Me.chkPlainEditPanel.TabIndex = 5
    Me.chkPlainEditPanel.Text = "Plain Edit Panel background and borders"
    Me.chkPlainEditPanel.UseVisualStyleBackColor = True
    '
    'cmdBackgroundImage
    '
    Me.cmdBackgroundImage.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdBackgroundImage.Location = New System.Drawing.Point(462, 39)
    Me.cmdBackgroundImage.Name = "cmdBackgroundImage"
    Me.cmdBackgroundImage.Size = New System.Drawing.Size(22, 22)
    Me.cmdBackgroundImage.TabIndex = 2
    Me.cmdBackgroundImage.Text = "?"
    Me.cmdBackgroundImage.UseVisualStyleBackColor = True
    '
    'lblBackgroundImageLayout
    '
    Me.lblBackgroundImageLayout.AutoSize = True
    Me.lblBackgroundImageLayout.Location = New System.Drawing.Point(8, 75)
    Me.lblBackgroundImageLayout.Name = "lblBackgroundImageLayout"
    Me.lblBackgroundImageLayout.Size = New System.Drawing.Size(132, 13)
    Me.lblBackgroundImageLayout.TabIndex = 3
    Me.lblBackgroundImageLayout.Text = "Background Image Layout"
    '
    'lblBackgroundImage
    '
    Me.lblBackgroundImage.AutoSize = True
    Me.lblBackgroundImage.Location = New System.Drawing.Point(8, 15)
    Me.lblBackgroundImage.Name = "lblBackgroundImage"
    Me.lblBackgroundImage.Size = New System.Drawing.Size(97, 13)
    Me.lblBackgroundImage.TabIndex = 0
    Me.lblBackgroundImage.Text = "Background Image"
    '
    'cboBackgroundImageLayout
    '
    Me.cboBackgroundImageLayout.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cboBackgroundImageLayout.FormattingEnabled = True
    Me.cboBackgroundImageLayout.Location = New System.Drawing.Point(222, 72)
    Me.cboBackgroundImageLayout.Name = "cboBackgroundImageLayout"
    Me.cboBackgroundImageLayout.Size = New System.Drawing.Size(134, 21)
    Me.cboBackgroundImageLayout.TabIndex = 4
    '
    'txtBackgroundImage
    '
    Me.txtBackgroundImage.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.txtBackgroundImage.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
    Me.txtBackgroundImage.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.FileSystem
    Me.txtBackgroundImage.Location = New System.Drawing.Point(11, 39)
    Me.txtBackgroundImage.Name = "txtBackgroundImage"
    Me.txtBackgroundImage.Size = New System.Drawing.Size(445, 20)
    Me.txtBackgroundImage.TabIndex = 1
    '
    'tabGeneral
    '
    Me.tabGeneral.Controls.Add(Me.chkFinderResultsMsgBox)
    Me.tabGeneral.Controls.Add(Me.chkHideHistoricNetwork)
    Me.tabGeneral.Controls.Add(Me.chkErrorsAsMsgbox)
    Me.tabGeneral.Controls.Add(Me.chkDisplayDashboardAtLogin)
    Me.tabGeneral.Controls.Add(Me.chkTabIntoHeaderPanel)
    Me.tabGeneral.Controls.Add(Me.chkTabIntoDisplayPanel)
    Me.tabGeneral.Controls.Add(Me.lblWebServicesTimeout)
    Me.tabGeneral.Controls.Add(Me.txtWebServicesTimeout)
    Me.tabGeneral.Controls.Add(Me.lblHistoryDays)
    Me.tabGeneral.Controls.Add(Me.txtHistoryDays)
    Me.tabGeneral.Location = New System.Drawing.Point(4, 22)
    Me.tabGeneral.Name = "tabGeneral"
    Me.tabGeneral.Padding = New System.Windows.Forms.Padding(3)
    Me.tabGeneral.Size = New System.Drawing.Size(794, 303)
    Me.tabGeneral.TabIndex = 2
    Me.tabGeneral.Text = "General"
    Me.tabGeneral.UseVisualStyleBackColor = True
    '
    'chkFinderResultsMsgBox
    '
    Me.chkFinderResultsMsgBox.AutoSize = True
    Me.chkFinderResultsMsgBox.Location = New System.Drawing.Point(11, 263)
    Me.chkFinderResultsMsgBox.Name = "chkFinderResultsMsgBox"
    Me.chkFinderResultsMsgBox.Size = New System.Drawing.Size(319, 17)
    Me.chkFinderResultsMsgBox.TabIndex = 9
    Me.chkFinderResultsMsgBox.Text = "Show number of records returned for finder and &QBE searches"
    Me.chkFinderResultsMsgBox.UseVisualStyleBackColor = True
    '
    'chkHideHistoricNetwork
    '
    Me.chkHideHistoricNetwork.AutoSize = True
    Me.chkHideHistoricNetwork.Location = New System.Drawing.Point(11, 229)
    Me.chkHideHistoricNetwork.Name = "chkHideHistoricNetwork"
    Me.chkHideHistoricNetwork.Size = New System.Drawing.Size(187, 17)
    Me.chkHideHistoricNetwork.TabIndex = 8
    Me.chkHideHistoricNetwork.Text = "Hide historic data in &Network view"
    Me.chkHideHistoricNetwork.UseVisualStyleBackColor = True
    '
    'chkErrorsAsMsgbox
    '
    Me.chkErrorsAsMsgbox.AutoSize = True
    Me.chkErrorsAsMsgbox.Location = New System.Drawing.Point(11, 195)
    Me.chkErrorsAsMsgbox.Name = "chkErrorsAsMsgbox"
    Me.chkErrorsAsMsgbox.Size = New System.Drawing.Size(192, 17)
    Me.chkErrorsAsMsgbox.TabIndex = 7
    Me.chkErrorsAsMsgbox.Text = "Show &editing errors in message box"
    Me.chkErrorsAsMsgbox.UseVisualStyleBackColor = True
    '
    'chkDisplayDashboardAtLogin
    '
    Me.chkDisplayDashboardAtLogin.AutoSize = True
    Me.chkDisplayDashboardAtLogin.Location = New System.Drawing.Point(11, 158)
    Me.chkDisplayDashboardAtLogin.Name = "chkDisplayDashboardAtLogin"
    Me.chkDisplayDashboardAtLogin.Size = New System.Drawing.Size(156, 17)
    Me.chkDisplayDashboardAtLogin.TabIndex = 6
    Me.chkDisplayDashboardAtLogin.Text = "Display Dashboard at &Login"
    Me.chkDisplayDashboardAtLogin.UseVisualStyleBackColor = True
    '
    'chkTabIntoHeaderPanel
    '
    Me.chkTabIntoHeaderPanel.AutoSize = True
    Me.chkTabIntoHeaderPanel.Location = New System.Drawing.Point(11, 121)
    Me.chkTabIntoHeaderPanel.Name = "chkTabIntoHeaderPanel"
    Me.chkTabIntoHeaderPanel.Size = New System.Drawing.Size(161, 17)
    Me.chkTabIntoHeaderPanel.TabIndex = 5
    Me.chkTabIntoHeaderPanel.Text = "Tab into &Header Panel Items"
    Me.chkTabIntoHeaderPanel.UseVisualStyleBackColor = True
    '
    'chkTabIntoDisplayPanel
    '
    Me.chkTabIntoDisplayPanel.AutoSize = True
    Me.chkTabIntoDisplayPanel.Location = New System.Drawing.Point(11, 84)
    Me.chkTabIntoDisplayPanel.Name = "chkTabIntoDisplayPanel"
    Me.chkTabIntoDisplayPanel.Size = New System.Drawing.Size(160, 17)
    Me.chkTabIntoDisplayPanel.TabIndex = 4
    Me.chkTabIntoDisplayPanel.Text = "Tab into &Display Panel Items"
    Me.chkTabIntoDisplayPanel.UseVisualStyleBackColor = True
    '
    'lblWebServicesTimeout
    '
    Me.lblWebServicesTimeout.AutoSize = True
    Me.lblWebServicesTimeout.Location = New System.Drawing.Point(8, 46)
    Me.lblWebServicesTimeout.Name = "lblWebServicesTimeout"
    Me.lblWebServicesTimeout.Size = New System.Drawing.Size(164, 13)
    Me.lblWebServicesTimeout.TabIndex = 2
    Me.lblWebServicesTimeout.Text = "Web Services Timeout (seconds)"
    '
    'txtWebServicesTimeout
    '
    Me.txtWebServicesTimeout.Location = New System.Drawing.Point(257, 43)
    Me.txtWebServicesTimeout.MaxLength = 4
    Me.txtWebServicesTimeout.Name = "txtWebServicesTimeout"
    Me.txtWebServicesTimeout.Size = New System.Drawing.Size(47, 20)
    Me.txtWebServicesTimeout.TabIndex = 3
    '
    'lblHistoryDays
    '
    Me.lblHistoryDays.AutoSize = True
    Me.lblHistoryDays.Location = New System.Drawing.Point(8, 16)
    Me.lblHistoryDays.Name = "lblHistoryDays"
    Me.lblHistoryDays.Size = New System.Drawing.Size(130, 13)
    Me.lblHistoryDays.TabIndex = 0
    Me.lblHistoryDays.Text = "Days to keep history items"
    '
    'txtHistoryDays
    '
    Me.txtHistoryDays.Location = New System.Drawing.Point(257, 13)
    Me.txtHistoryDays.MaxLength = 3
    Me.txtHistoryDays.Name = "txtHistoryDays"
    Me.txtHistoryDays.Size = New System.Drawing.Size(47, 20)
    Me.txtHistoryDays.TabIndex = 1
    '
    'tabConfirmation
    '
    Me.tabConfirmation.Controls.Add(Me.chkConfirmCancel)
    Me.tabConfirmation.Controls.Add(Me.chkConfirmUpdate)
    Me.tabConfirmation.Controls.Add(Me.chkConfirmInsert)
    Me.tabConfirmation.Controls.Add(Me.chkConfirmDelete)
    Me.tabConfirmation.Location = New System.Drawing.Point(4, 22)
    Me.tabConfirmation.Name = "tabConfirmation"
    Me.tabConfirmation.Size = New System.Drawing.Size(794, 303)
    Me.tabConfirmation.TabIndex = 5
    Me.tabConfirmation.Text = "Confirmation"
    Me.tabConfirmation.UseVisualStyleBackColor = True
    '
    'chkConfirmCancel
    '
    Me.chkConfirmCancel.AutoSize = True
    Me.chkConfirmCancel.Location = New System.Drawing.Point(18, 123)
    Me.chkConfirmCancel.Name = "chkConfirmCancel"
    Me.chkConfirmCancel.Size = New System.Drawing.Size(97, 17)
    Me.chkConfirmCancel.TabIndex = 3
    Me.chkConfirmCancel.Text = "Confirm &Cancel"
    Me.chkConfirmCancel.UseVisualStyleBackColor = True
    '
    'chkConfirmUpdate
    '
    Me.chkConfirmUpdate.AutoSize = True
    Me.chkConfirmUpdate.Location = New System.Drawing.Point(18, 86)
    Me.chkConfirmUpdate.Name = "chkConfirmUpdate"
    Me.chkConfirmUpdate.Size = New System.Drawing.Size(99, 17)
    Me.chkConfirmUpdate.TabIndex = 2
    Me.chkConfirmUpdate.Text = "Confirm &Update"
    Me.chkConfirmUpdate.UseVisualStyleBackColor = True
    '
    'chkConfirmInsert
    '
    Me.chkConfirmInsert.AutoSize = True
    Me.chkConfirmInsert.Location = New System.Drawing.Point(18, 49)
    Me.chkConfirmInsert.Name = "chkConfirmInsert"
    Me.chkConfirmInsert.Size = New System.Drawing.Size(90, 17)
    Me.chkConfirmInsert.TabIndex = 1
    Me.chkConfirmInsert.Text = "Confirm &Insert"
    Me.chkConfirmInsert.UseVisualStyleBackColor = True
    '
    'chkConfirmDelete
    '
    Me.chkConfirmDelete.AutoSize = True
    Me.chkConfirmDelete.Location = New System.Drawing.Point(18, 13)
    Me.chkConfirmDelete.Name = "chkConfirmDelete"
    Me.chkConfirmDelete.Size = New System.Drawing.Size(95, 17)
    Me.chkConfirmDelete.TabIndex = 0
    Me.chkConfirmDelete.Text = "Confirm &Delete"
    Me.chkConfirmDelete.UseVisualStyleBackColor = True
    '
    'cmdOK
    '
    Me.cmdOK.Location = New System.Drawing.Point(186, 6)
    Me.cmdOK.Name = "cmdOK"
    Me.cmdOK.Size = New System.Drawing.Size(96, 27)
    Me.cmdOK.TabIndex = 0
    Me.cmdOK.Text = "OK"
    '
    'cmdCancel
    '
    Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdCancel.Location = New System.Drawing.Point(630, 6)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(96, 27)
    Me.cmdCancel.TabIndex = 4
    Me.cmdCancel.Text = "Cancel"
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdDesign)
    Me.bpl.Controls.Add(Me.cmdOK)
    Me.bpl.Controls.Add(Me.cmdSaveAs)
    Me.bpl.Controls.Add(Me.cmdApply)
    Me.bpl.Controls.Add(Me.cmdDefaults)
    Me.bpl.Controls.Add(Me.cmdCancel)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bpl.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bpl.Location = New System.Drawing.Point(0, 329)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(802, 39)
    Me.bpl.TabIndex = 2
    '
    'cmdDesign
    '
    Me.cmdDesign.Location = New System.Drawing.Point(75, 6)
    Me.cmdDesign.Name = "cmdDesign"
    Me.cmdDesign.Size = New System.Drawing.Size(96, 27)
    Me.cmdDesign.TabIndex = 5
    Me.cmdDesign.Text = "&Design..."
    '
    'cmdSaveAs
    '
    Me.cmdSaveAs.Location = New System.Drawing.Point(297, 6)
    Me.cmdSaveAs.Name = "cmdSaveAs"
    Me.cmdSaveAs.Size = New System.Drawing.Size(96, 27)
    Me.cmdSaveAs.TabIndex = 1
    Me.cmdSaveAs.Text = "&Save As"
    '
    'cmdApply
    '
    Me.cmdApply.Location = New System.Drawing.Point(408, 6)
    Me.cmdApply.Name = "cmdApply"
    Me.cmdApply.Size = New System.Drawing.Size(96, 27)
    Me.cmdApply.TabIndex = 2
    Me.cmdApply.Text = "&Apply"
    '
    'cmdDefaults
    '
    Me.cmdDefaults.Location = New System.Drawing.Point(519, 6)
    Me.cmdDefaults.Name = "cmdDefaults"
    Me.cmdDefaults.Size = New System.Drawing.Size(96, 27)
    Me.cmdDefaults.TabIndex = 3
    Me.cmdDefaults.Text = "Defaults"
    '
    'tim
    '
    '
    'frmPreferences
    '
    Me.AcceptButton = Me.cmdOK
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.cmdCancel
    Me.ClientSize = New System.Drawing.Size(802, 368)
    Me.Controls.Add(Me.tab)
    Me.Controls.Add(Me.bpl)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmPreferences"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.Text = "User Preferences"
    Me.tab.ResumeLayout(False)
    Me.tabNotification.ResumeLayout(False)
    Me.tabNotification.PerformLayout()
    Me.tabDisplay.ResumeLayout(False)
    Me.tabDisplay.PerformLayout()
    Me.tabGeneral.ResumeLayout(False)
    Me.tabGeneral.PerformLayout()
    Me.tabConfirmation.ResumeLayout(False)
    Me.tabConfirmation.PerformLayout()
    Me.bpl.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents bpl As CDBNETCL.ButtonPanel
  Friend WithEvents cmdOK As System.Windows.Forms.Button
  Friend WithEvents cmdDefaults As System.Windows.Forms.Button
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents tab As System.Windows.Forms.TabControl
  Friend WithEvents tabNotification As System.Windows.Forms.TabPage
  Friend WithEvents tabDisplay As System.Windows.Forms.TabPage
  Friend WithEvents chkNotifyDocuments As System.Windows.Forms.CheckBox
  Friend WithEvents chkNotifyActions As System.Windows.Forms.CheckBox
  Friend WithEvents lblPollingInterval As System.Windows.Forms.Label
  Friend WithEvents txtPollingInterval As System.Windows.Forms.TextBox
  Friend WithEvents chkNotifyDeadlines As System.Windows.Forms.CheckBox
  Friend WithEvents lblBackgroundImageLayout As System.Windows.Forms.Label
  Friend WithEvents lblBackgroundImage As System.Windows.Forms.Label
  Friend WithEvents cboBackgroundImageLayout As System.Windows.Forms.ComboBox
  Friend WithEvents txtBackgroundImage As System.Windows.Forms.TextBox
  Friend WithEvents cmdBackgroundImage As System.Windows.Forms.Button
  Friend WithEvents cmd As System.Windows.Forms.OpenFileDialog
  Friend WithEvents tabGeneral As System.Windows.Forms.TabPage
  Friend WithEvents lblHistoryDays As System.Windows.Forms.Label
  Friend WithEvents txtHistoryDays As System.Windows.Forms.TextBox
  Friend WithEvents lblWebServicesTimeout As System.Windows.Forms.Label
  Friend WithEvents txtWebServicesTimeout As System.Windows.Forms.TextBox
  Friend WithEvents chkPlainEditPanel As System.Windows.Forms.CheckBox
  Friend WithEvents cdlg As System.Windows.Forms.ColorDialog
  Friend WithEvents cmdApply As System.Windows.Forms.Button
  Friend WithEvents cmdSaveAs As System.Windows.Forms.Button
  Friend WithEvents lblSchemes As System.Windows.Forms.Label
  Friend WithEvents cboSchemes As System.Windows.Forms.ComboBox
  Friend WithEvents tim As System.Windows.Forms.Timer
  Friend WithEvents chkTabIntoHeaderPanel As System.Windows.Forms.CheckBox
  Friend WithEvents chkTabIntoDisplayPanel As System.Windows.Forms.CheckBox
  Friend WithEvents chkDisplayDashboardAtLogin As System.Windows.Forms.CheckBox
  Friend WithEvents tabConfirmation As System.Windows.Forms.TabPage
  Friend WithEvents chkConfirmCancel As System.Windows.Forms.CheckBox
  Friend WithEvents chkConfirmUpdate As System.Windows.Forms.CheckBox
  Friend WithEvents chkConfirmInsert As System.Windows.Forms.CheckBox
  Friend WithEvents chkConfirmDelete As System.Windows.Forms.CheckBox
  Friend WithEvents chkErrorsAsMsgbox As System.Windows.Forms.CheckBox
  Friend WithEvents chkNotifyMeetings As System.Windows.Forms.CheckBox
  Friend WithEvents lblTaskPollingInterval As System.Windows.Forms.Label
  Friend WithEvents txtTaskPollingInterval As System.Windows.Forms.TextBox
  Friend WithEvents chkHideHistoricNetwork As System.Windows.Forms.CheckBox
  Friend WithEvents chkFinderResultsMsgBox As System.Windows.Forms.CheckBox
  Friend WithEvents cmdDesign As System.Windows.Forms.Button
End Class
