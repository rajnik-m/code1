<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDocumentation
  Inherits System.Windows.Forms.Form

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
    Me.cmdGo = New System.Windows.Forms.Button()
    Me.sts = New System.Windows.Forms.StatusStrip()
    Me.ssl = New System.Windows.Forms.ToolStripStatusLabel()
    Me.chkBoth = New System.Windows.Forms.CheckBox()
    Me.chkDifferencesOnly = New System.Windows.Forms.CheckBox()
    Me.chkBuildVersion = New System.Windows.Forms.CheckBox()
    Me.rtb = New System.Windows.Forms.RichTextBox()
    Me.chkWebServices = New System.Windows.Forms.CheckBox()
    Me.chkJira = New System.Windows.Forms.CheckBox()
    Me.txtJira = New System.Windows.Forms.TextBox()
    Me.cmdBrowse = New System.Windows.Forms.Button()
    Me.ofdJira = New System.Windows.Forms.OpenFileDialog()
    Me.chkWSSetupFile = New System.Windows.Forms.CheckBox()
    Me.cmdExams = New System.Windows.Forms.Button()
    Me.sts.SuspendLayout()
    Me.SuspendLayout()
    '
    'cmdGo
    '
    Me.cmdGo.Location = New System.Drawing.Point(183, 155)
    Me.cmdGo.Margin = New System.Windows.Forms.Padding(2)
    Me.cmdGo.Name = "cmdGo"
    Me.cmdGo.Size = New System.Drawing.Size(136, 30)
    Me.cmdGo.TabIndex = 0
    Me.cmdGo.Text = "OK"
    Me.cmdGo.UseVisualStyleBackColor = True
    '
    'sts
    '
    Me.sts.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ssl})
    Me.sts.Location = New System.Drawing.Point(0, 191)
    Me.sts.Name = "sts"
    Me.sts.Padding = New System.Windows.Forms.Padding(1, 0, 10, 0)
    Me.sts.Size = New System.Drawing.Size(498, 22)
    Me.sts.TabIndex = 1
    Me.sts.Text = "Test"
    '
    'ssl
    '
    Me.ssl.Name = "ssl"
    Me.ssl.Size = New System.Drawing.Size(0, 17)
    '
    'chkBoth
    '
    Me.chkBoth.AutoSize = True
    Me.chkBoth.Location = New System.Drawing.Point(20, 32)
    Me.chkBoth.Margin = New System.Windows.Forms.Padding(2)
    Me.chkBoth.Name = "chkBoth"
    Me.chkBoth.Size = New System.Drawing.Size(267, 17)
    Me.chkBoth.TabIndex = 2
    Me.chkBoth.Text = "Build Web Service Documentation and Differences"
    Me.chkBoth.UseVisualStyleBackColor = True
    '
    'chkDifferencesOnly
    '
    Me.chkDifferencesOnly.AutoSize = True
    Me.chkDifferencesOnly.Location = New System.Drawing.Point(20, 54)
    Me.chkDifferencesOnly.Margin = New System.Windows.Forms.Padding(2)
    Me.chkDifferencesOnly.Name = "chkDifferencesOnly"
    Me.chkDifferencesOnly.Size = New System.Drawing.Size(130, 17)
    Me.chkDifferencesOnly.TabIndex = 3
    Me.chkDifferencesOnly.Text = "Build Differences Only"
    Me.chkDifferencesOnly.UseVisualStyleBackColor = True
    '
    'chkBuildVersion
    '
    Me.chkBuildVersion.AutoSize = True
    Me.chkBuildVersion.Checked = True
    Me.chkBuildVersion.CheckState = System.Windows.Forms.CheckState.Checked
    Me.chkBuildVersion.Location = New System.Drawing.Point(20, 98)
    Me.chkBuildVersion.Margin = New System.Windows.Forms.Padding(2)
    Me.chkBuildVersion.Name = "chkBuildVersion"
    Me.chkBuildVersion.Size = New System.Drawing.Size(162, 17)
    Me.chkBuildVersion.TabIndex = 4
    Me.chkBuildVersion.Text = "Build Version Documentation"
    Me.chkBuildVersion.UseVisualStyleBackColor = True
    '
    'rtb
    '
    Me.rtb.Location = New System.Drawing.Point(231, 102)
    Me.rtb.Margin = New System.Windows.Forms.Padding(2)
    Me.rtb.Name = "rtb"
    Me.rtb.Size = New System.Drawing.Size(194, 15)
    Me.rtb.TabIndex = 5
    Me.rtb.Text = ""
    Me.rtb.Visible = False
    '
    'chkWebServices
    '
    Me.chkWebServices.AutoSize = True
    Me.chkWebServices.Location = New System.Drawing.Point(20, 10)
    Me.chkWebServices.Margin = New System.Windows.Forms.Padding(2)
    Me.chkWebServices.Name = "chkWebServices"
    Me.chkWebServices.Size = New System.Drawing.Size(189, 17)
    Me.chkWebServices.TabIndex = 6
    Me.chkWebServices.Text = "Build Web Service Documentation"
    Me.chkWebServices.UseVisualStyleBackColor = True
    '
    'chkJira
    '
    Me.chkJira.AutoSize = True
    Me.chkJira.Location = New System.Drawing.Point(20, 120)
    Me.chkJira.Margin = New System.Windows.Forms.Padding(2)
    Me.chkJira.Name = "chkJira"
    Me.chkJira.Size = New System.Drawing.Size(107, 17)
    Me.chkJira.TabIndex = 7
    Me.chkJira.Text = "Process Jira data"
    Me.chkJira.UseVisualStyleBackColor = True
    '
    'txtJira
    '
    Me.txtJira.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.FileSystem
    Me.txtJira.Location = New System.Drawing.Point(142, 122)
    Me.txtJira.Name = "txtJira"
    Me.txtJira.Size = New System.Drawing.Size(283, 20)
    Me.txtJira.TabIndex = 8
    '
    'cmdBrowse
    '
    Me.cmdBrowse.Location = New System.Drawing.Point(431, 121)
    Me.cmdBrowse.Name = "cmdBrowse"
    Me.cmdBrowse.Size = New System.Drawing.Size(57, 20)
    Me.cmdBrowse.TabIndex = 9
    Me.cmdBrowse.Text = "Browse"
    Me.cmdBrowse.UseVisualStyleBackColor = True
    '
    'ofdJira
    '
    Me.ofdJira.FileName = "OpenFileDialog1"
    '
    'chkWSSetupFile
    '
    Me.chkWSSetupFile.AutoSize = True
    Me.chkWSSetupFile.Location = New System.Drawing.Point(20, 76)
    Me.chkWSSetupFile.Margin = New System.Windows.Forms.Padding(2)
    Me.chkWSSetupFile.Name = "chkWSSetupFile"
    Me.chkWSSetupFile.Size = New System.Drawing.Size(177, 17)
    Me.chkWSSetupFile.TabIndex = 10
    Me.chkWSSetupFile.Text = "Check Web Services Setup File"
    Me.chkWSSetupFile.UseVisualStyleBackColor = True
    '
    'cmdExams
    '
    Me.cmdExams.Location = New System.Drawing.Point(337, 155)
    Me.cmdExams.Margin = New System.Windows.Forms.Padding(2)
    Me.cmdExams.Name = "cmdExams"
    Me.cmdExams.Size = New System.Drawing.Size(136, 30)
    Me.cmdExams.TabIndex = 11
    Me.cmdExams.Text = "Exams!"
    Me.cmdExams.UseVisualStyleBackColor = True
    Me.cmdExams.Visible = False
    '
    'frmDocumentation
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(498, 213)
    Me.Controls.Add(Me.cmdExams)
    Me.Controls.Add(Me.chkWSSetupFile)
    Me.Controls.Add(Me.cmdBrowse)
    Me.Controls.Add(Me.txtJira)
    Me.Controls.Add(Me.chkJira)
    Me.Controls.Add(Me.chkWebServices)
    Me.Controls.Add(Me.rtb)
    Me.Controls.Add(Me.chkBuildVersion)
    Me.Controls.Add(Me.chkDifferencesOnly)
    Me.Controls.Add(Me.chkBoth)
    Me.Controls.Add(Me.sts)
    Me.Controls.Add(Me.cmdGo)
    Me.Margin = New System.Windows.Forms.Padding(2)
    Me.Name = "frmDocumentation"
    Me.Text = "Care Documentation Creation"
    Me.sts.ResumeLayout(False)
    Me.sts.PerformLayout()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents cmdGo As System.Windows.Forms.Button
  Friend WithEvents sts As System.Windows.Forms.StatusStrip
  Friend WithEvents ssl As System.Windows.Forms.ToolStripStatusLabel
  Friend WithEvents chkBoth As System.Windows.Forms.CheckBox
  Friend WithEvents chkDifferencesOnly As System.Windows.Forms.CheckBox
  Friend WithEvents chkBuildVersion As System.Windows.Forms.CheckBox
  Friend WithEvents rtb As System.Windows.Forms.RichTextBox
  Friend WithEvents chkWebServices As System.Windows.Forms.CheckBox
  Friend WithEvents chkJira As System.Windows.Forms.CheckBox
  Friend WithEvents txtJira As System.Windows.Forms.TextBox
  Friend WithEvents cmdBrowse As System.Windows.Forms.Button
  Friend WithEvents ofdJira As System.Windows.Forms.OpenFileDialog
  Friend WithEvents chkWSSetupFile As System.Windows.Forms.CheckBox
  Friend WithEvents cmdExams As System.Windows.Forms.Button

End Class
