<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTaskInfo
  Inherits PersistentForm

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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmTaskInfo))
    Me.Panel1 = New System.Windows.Forms.Panel()
    Me.lblUpgrade = New System.Windows.Forms.Label()
    Me.lblNote = New System.Windows.Forms.Label()
    Me.lblJobNumber = New System.Windows.Forms.Label()
    Me.lblTaskJobNumber = New System.Windows.Forms.Label()
    Me.lblStatus = New System.Windows.Forms.Label()
    Me.lblTaskStatus = New System.Windows.Forms.Label()
    Me.lblTask = New System.Windows.Forms.Label()
    Me.bpl = New CDBNETCL.ButtonPanel()
    Me.cmdClose = New System.Windows.Forms.Button()
    Me.Panel1.SuspendLayout()
    Me.bpl.SuspendLayout()
    Me.SuspendLayout()
    '
    'Panel1
    '
    Me.Panel1.Controls.Add(Me.lblUpgrade)
    Me.Panel1.Controls.Add(Me.lblNote)
    Me.Panel1.Controls.Add(Me.lblJobNumber)
    Me.Panel1.Controls.Add(Me.lblTaskJobNumber)
    Me.Panel1.Controls.Add(Me.lblStatus)
    Me.Panel1.Controls.Add(Me.lblTaskStatus)
    Me.Panel1.Controls.Add(Me.lblTask)
    Me.Panel1.Controls.Add(Me.bpl)
    Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
    Me.Panel1.Location = New System.Drawing.Point(0, 0)
    Me.Panel1.Name = "Panel1"
    Me.Panel1.Size = New System.Drawing.Size(541, 177)
    Me.Panel1.TabIndex = 0
    '
    'lblUpgrade
    '
    Me.lblUpgrade.AutoSize = True
    Me.lblUpgrade.Location = New System.Drawing.Point(14, 111)
    Me.lblUpgrade.Name = "lblUpgrade"
    Me.lblUpgrade.Size = New System.Drawing.Size(404, 13)
    Me.lblUpgrade.TabIndex = 7
    Me.lblUpgrade.Text = "You should wait for the upgrade process to finish before performing any other act" & _
    "ivity"
    '
    'lblNote
    '
    Me.lblNote.AutoSize = True
    Me.lblNote.Location = New System.Drawing.Point(14, 34)
    Me.lblNote.Name = "lblNote"
    Me.lblNote.Size = New System.Drawing.Size(354, 13)
    Me.lblNote.TabIndex = 6
    Me.lblNote.Text = "NOTE: Closing this window will not stop the process running on the server"
    '
    'lblJobNumber
    '
    Me.lblJobNumber.AutoSize = True
    Me.lblJobNumber.Location = New System.Drawing.Point(87, 59)
    Me.lblJobNumber.Name = "lblJobNumber"
    Me.lblJobNumber.Size = New System.Drawing.Size(130, 13)
    Me.lblJobNumber.TabIndex = 5
    Me.lblJobNumber.Text = "Searching for Job Number"
    '
    'lblTaskJobNumber
    '
    Me.lblTaskJobNumber.AutoSize = True
    Me.lblTaskJobNumber.Location = New System.Drawing.Point(14, 59)
    Me.lblTaskJobNumber.Name = "lblTaskJobNumber"
    Me.lblTaskJobNumber.Size = New System.Drawing.Size(67, 13)
    Me.lblTaskJobNumber.TabIndex = 4
    Me.lblTaskJobNumber.Text = "Job Number:"
    '
    'lblStatus
    '
    Me.lblStatus.AutoSize = True
    Me.lblStatus.Location = New System.Drawing.Point(87, 82)
    Me.lblStatus.Name = "lblStatus"
    Me.lblStatus.Size = New System.Drawing.Size(130, 13)
    Me.lblStatus.TabIndex = 3
    Me.lblStatus.Text = "Searching for Task Status"
    '
    'lblTaskStatus
    '
    Me.lblTaskStatus.AutoSize = True
    Me.lblTaskStatus.Location = New System.Drawing.Point(14, 82)
    Me.lblTaskStatus.Name = "lblTaskStatus"
    Me.lblTaskStatus.Size = New System.Drawing.Size(67, 13)
    Me.lblTaskStatus.TabIndex = 2
    Me.lblTaskStatus.Text = "Task Status:"
    '
    'lblTask
    '
    Me.lblTask.AutoSize = True
    Me.lblTask.Location = New System.Drawing.Point(14, 9)
    Me.lblTask.Name = "lblTask"
    Me.lblTask.Size = New System.Drawing.Size(301, 13)
    Me.lblTask.TabIndex = 1
    Me.lblTask.Text = "The {0} task has been set to run asynchronously on the server"
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdClose)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bpl.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bpl.Location = New System.Drawing.Point(0, 138)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(541, 39)
    Me.bpl.TabIndex = 0
    '
    'cmdClose
    '
    Me.cmdClose.DialogResult = System.Windows.Forms.DialogResult.OK
    Me.cmdClose.Location = New System.Drawing.Point(222, 6)
    Me.cmdClose.Name = "cmdClose"
    Me.cmdClose.Size = New System.Drawing.Size(96, 27)
    Me.cmdClose.TabIndex = 0
    Me.cmdClose.Text = "Close"
    Me.cmdClose.UseVisualStyleBackColor = True
    '
    'frmTaskInfo
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(541, 177)
    Me.Controls.Add(Me.Panel1)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmTaskInfo"
    Me.Text = "frmTaskStatus"
    Me.Panel1.ResumeLayout(False)
    Me.Panel1.PerformLayout()
    Me.bpl.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents Panel1 As System.Windows.Forms.Panel
  Friend WithEvents bpl As CDBNETCL.ButtonPanel
  Friend WithEvents cmdClose As System.Windows.Forms.Button
  Friend WithEvents lblTask As System.Windows.Forms.Label
  Friend WithEvents lblStatus As System.Windows.Forms.Label
  Friend WithEvents lblTaskStatus As System.Windows.Forms.Label
  Friend WithEvents lblJobNumber As System.Windows.Forms.Label
  Friend WithEvents lblTaskJobNumber As System.Windows.Forms.Label
  Friend WithEvents lblUpgrade As System.Windows.Forms.Label
  Friend WithEvents lblNote As System.Windows.Forms.Label
End Class
