<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmGenMMerge
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmGenMMerge))
    Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
    Me.lblDesc = New System.Windows.Forms.Label()
    Me.gbpLabels = New System.Windows.Forms.GroupBox()
    Me.Label2 = New System.Windows.Forms.Label()
    Me.Label1 = New System.Windows.Forms.Label()
    Me.lblC = New System.Windows.Forms.Label()
    Me.lblB = New System.Windows.Forms.Label()
    Me.lblA = New System.Windows.Forms.Label()
    Me.lbl3 = New System.Windows.Forms.Label()
    Me.lbl2 = New System.Windows.Forms.Label()
    Me.lbl1 = New System.Windows.Forms.Label()
    Me.Panel2 = New System.Windows.Forms.Panel()
    Me.Panel1 = New System.Windows.Forms.Panel()
    Me.gbx = New System.Windows.Forms.GroupBox()
    Me.optAAndC = New System.Windows.Forms.RadioButton()
    Me.optCOnly = New System.Windows.Forms.RadioButton()
    Me.optBOnly = New System.Windows.Forms.RadioButton()
    Me.optAOnly = New System.Windows.Forms.RadioButton()
    Me.optMergerAll = New System.Windows.Forms.RadioButton()
    Me.ButtonPanel1 = New CDBNETCL.ButtonPanel()
    Me.cmdOK = New System.Windows.Forms.Button()
    Me.cmdCancel = New System.Windows.Forms.Button()
    CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SplitContainer1.Panel1.SuspendLayout()
    Me.SplitContainer1.Panel2.SuspendLayout()
    Me.SplitContainer1.SuspendLayout()
    Me.gbpLabels.SuspendLayout()
    Me.gbx.SuspendLayout()
    Me.ButtonPanel1.SuspendLayout()
    Me.SuspendLayout()
    '
    'SplitContainer1
    '
    Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
    Me.SplitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel2
    Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
    Me.SplitContainer1.Name = "SplitContainer1"
    Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
    '
    'SplitContainer1.Panel1
    '
    Me.SplitContainer1.Panel1.Controls.Add(Me.lblDesc)
    Me.SplitContainer1.Panel1.Controls.Add(Me.gbpLabels)
    Me.SplitContainer1.Panel1.Controls.Add(Me.gbx)
    '
    'SplitContainer1.Panel2
    '
    Me.SplitContainer1.Panel2.Controls.Add(Me.ButtonPanel1)
    Me.SplitContainer1.Size = New System.Drawing.Size(603, 471)
    Me.SplitContainer1.SplitterDistance = 412
    Me.SplitContainer1.TabIndex = 0
    '
    'lblDesc
    '
    Me.lblDesc.Location = New System.Drawing.Point(3, 372)
    Me.lblDesc.Name = "lblDesc"
    Me.lblDesc.Size = New System.Drawing.Size(556, 43)
    Me.lblDesc.TabIndex = 2
    '
    'gbpLabels
    '
    Me.gbpLabels.Controls.Add(Me.Label2)
    Me.gbpLabels.Controls.Add(Me.Label1)
    Me.gbpLabels.Controls.Add(Me.lblC)
    Me.gbpLabels.Controls.Add(Me.lblB)
    Me.gbpLabels.Controls.Add(Me.lblA)
    Me.gbpLabels.Controls.Add(Me.lbl3)
    Me.gbpLabels.Controls.Add(Me.lbl2)
    Me.gbpLabels.Controls.Add(Me.lbl1)
    Me.gbpLabels.Controls.Add(Me.Panel2)
    Me.gbpLabels.Controls.Add(Me.Panel1)
    Me.gbpLabels.Location = New System.Drawing.Point(13, 13)
    Me.gbpLabels.Name = "gbpLabels"
    Me.gbpLabels.Size = New System.Drawing.Size(419, 355)
    Me.gbpLabels.TabIndex = 1
    Me.gbpLabels.TabStop = False
    Me.gbpLabels.Text = " Merged Lists"
    '
    'Label2
    '
    Me.Label2.Location = New System.Drawing.Point(9, 112)
    Me.Label2.Name = "Label2"
    Me.Label2.Size = New System.Drawing.Size(63, 53)
    Me.Label2.TabIndex = 13
    Me.Label2.Tag = ""
    Me.Label2.Text = "Current List"
    Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopCenter
    '
    'Label1
    '
    Me.Label1.Location = New System.Drawing.Point(335, 198)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(72, 56)
    Me.Label1.TabIndex = 12
    Me.Label1.Tag = ""
    Me.Label1.Text = "List From New Criteria"
    Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
    '
    'lblC
    '
    Me.lblC.BackColor = System.Drawing.Color.White
    Me.lblC.Location = New System.Drawing.Point(119, 259)
    Me.lblC.Name = "lblC"
    Me.lblC.Size = New System.Drawing.Size(170, 43)
    Me.lblC.TabIndex = 9
    Me.lblC.Text = "C. Records only selected by the new criteria"
    Me.lblC.TextAlign = System.Drawing.ContentAlignment.TopCenter
    '
    'lblB
    '
    Me.lblB.BackColor = System.Drawing.Color.White
    Me.lblB.Location = New System.Drawing.Point(119, 156)
    Me.lblB.Name = "lblB"
    Me.lblB.Size = New System.Drawing.Size(170, 43)
    Me.lblB.TabIndex = 8
    Me.lblB.Text = "B. Records common to both lists"
    Me.lblB.TextAlign = System.Drawing.ContentAlignment.TopCenter
    '
    'lblA
    '
    Me.lblA.BackColor = System.Drawing.Color.White
    Me.lblA.Location = New System.Drawing.Point(119, 61)
    Me.lblA.Name = "lblA"
    Me.lblA.Size = New System.Drawing.Size(170, 43)
    Me.lblA.TabIndex = 7
    Me.lblA.Text = "A. Records Only Found in Current List"
    Me.lblA.TextAlign = System.Drawing.ContentAlignment.TopCenter
    '
    'lbl3
    '
    Me.lbl3.BackColor = System.Drawing.Color.White
    Me.lbl3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
    Me.lbl3.Font = New System.Drawing.Font("Microsoft Sans Serif", 3.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.lbl3.Location = New System.Drawing.Point(93, 233)
    Me.lbl3.Name = "lbl3"
    Me.lbl3.Size = New System.Drawing.Size(223, 103)
    Me.lbl3.TabIndex = 6
    Me.lbl3.Text = resources.GetString("lbl3.Text")
    '
    'lbl2
    '
    Me.lbl2.BackColor = System.Drawing.Color.White
    Me.lbl2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
    Me.lbl2.Font = New System.Drawing.Font("Microsoft Sans Serif", 3.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.lbl2.Location = New System.Drawing.Point(93, 130)
    Me.lbl2.Name = "lbl2"
    Me.lbl2.Size = New System.Drawing.Size(223, 103)
    Me.lbl2.TabIndex = 5
    Me.lbl2.Text = resources.GetString("lbl2.Text")
    '
    'lbl1
    '
    Me.lbl1.BackColor = System.Drawing.Color.White
    Me.lbl1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
    Me.lbl1.Font = New System.Drawing.Font("Microsoft Sans Serif", 3.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.lbl1.Location = New System.Drawing.Point(93, 27)
    Me.lbl1.Name = "lbl1"
    Me.lbl1.Size = New System.Drawing.Size(223, 103)
    Me.lbl1.TabIndex = 4
    Me.lbl1.Text = resources.GetString("lbl1.Text")
    '
    'Panel2
    '
    Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
    Me.Panel2.Location = New System.Drawing.Point(37, 27)
    Me.Panel2.Name = "Panel2"
    Me.Panel2.Size = New System.Drawing.Size(279, 206)
    Me.Panel2.TabIndex = 15
    '
    'Panel1
    '
    Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
    Me.Panel1.Location = New System.Drawing.Point(93, 130)
    Me.Panel1.Name = "Panel1"
    Me.Panel1.Size = New System.Drawing.Size(279, 206)
    Me.Panel1.TabIndex = 14
    '
    'gbx
    '
    Me.gbx.Controls.Add(Me.optAAndC)
    Me.gbx.Controls.Add(Me.optCOnly)
    Me.gbx.Controls.Add(Me.optBOnly)
    Me.gbx.Controls.Add(Me.optAOnly)
    Me.gbx.Controls.Add(Me.optMergerAll)
    Me.gbx.Location = New System.Drawing.Point(450, 182)
    Me.gbx.Name = "gbx"
    Me.gbx.Size = New System.Drawing.Size(148, 187)
    Me.gbx.TabIndex = 0
    Me.gbx.TabStop = False
    Me.gbx.Text = "Selection Choices"
    '
    'optAAndC
    '
    Me.optAAndC.AutoSize = True
    Me.optAAndC.Location = New System.Drawing.Point(18, 154)
    Me.optAAndC.Name = "optAAndC"
    Me.optAAndC.Size = New System.Drawing.Size(63, 17)
    Me.optAAndC.TabIndex = 4
    Me.optAAndC.TabStop = True
    Me.optAAndC.Text = "A and C"
    Me.optAAndC.UseVisualStyleBackColor = True
    '
    'optCOnly
    '
    Me.optCOnly.AutoSize = True
    Me.optCOnly.Location = New System.Drawing.Point(18, 123)
    Me.optCOnly.Name = "optCOnly"
    Me.optCOnly.Size = New System.Drawing.Size(56, 17)
    Me.optCOnly.TabIndex = 3
    Me.optCOnly.TabStop = True
    Me.optCOnly.Text = "C Only"
    Me.optCOnly.UseVisualStyleBackColor = True
    '
    'optBOnly
    '
    Me.optBOnly.AutoSize = True
    Me.optBOnly.Location = New System.Drawing.Point(18, 88)
    Me.optBOnly.Name = "optBOnly"
    Me.optBOnly.Size = New System.Drawing.Size(56, 17)
    Me.optBOnly.TabIndex = 2
    Me.optBOnly.TabStop = True
    Me.optBOnly.Text = "B Only"
    Me.optBOnly.UseVisualStyleBackColor = True
    '
    'optAOnly
    '
    Me.optAOnly.AutoSize = True
    Me.optAOnly.Location = New System.Drawing.Point(18, 55)
    Me.optAOnly.Name = "optAOnly"
    Me.optAOnly.Size = New System.Drawing.Size(56, 17)
    Me.optAOnly.TabIndex = 1
    Me.optAOnly.TabStop = True
    Me.optAOnly.Text = "A Only"
    Me.optAOnly.UseVisualStyleBackColor = True
    '
    'optMergerAll
    '
    Me.optMergerAll.AutoSize = True
    Me.optMergerAll.Location = New System.Drawing.Point(18, 20)
    Me.optMergerAll.Name = "optMergerAll"
    Me.optMergerAll.Size = New System.Drawing.Size(36, 17)
    Me.optMergerAll.TabIndex = 0
    Me.optMergerAll.TabStop = True
    Me.optMergerAll.Text = "All"
    Me.optMergerAll.UseVisualStyleBackColor = True
    '
    'ButtonPanel1
    '
    Me.ButtonPanel1.Controls.Add(Me.cmdOK)
    Me.ButtonPanel1.Controls.Add(Me.cmdCancel)
    Me.ButtonPanel1.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.ButtonPanel1.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.ButtonPanel1.Location = New System.Drawing.Point(0, 16)
    Me.ButtonPanel1.Name = "ButtonPanel1"
    Me.ButtonPanel1.Size = New System.Drawing.Size(603, 39)
    Me.ButtonPanel1.TabIndex = 0
    '
    'cmdOK
    '
    Me.cmdOK.Location = New System.Drawing.Point(198, 6)
    Me.cmdOK.Name = "cmdOK"
    Me.cmdOK.Size = New System.Drawing.Size(96, 27)
    Me.cmdOK.TabIndex = 0
    Me.cmdOK.Text = "&OK"
    Me.cmdOK.UseVisualStyleBackColor = True
    '
    'cmdCancel
    '
    Me.cmdCancel.Location = New System.Drawing.Point(309, 6)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(96, 27)
    Me.cmdCancel.TabIndex = 1
    Me.cmdCancel.Text = "Cancel"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'frmGenMMerge
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(603, 471)
    Me.Controls.Add(Me.SplitContainer1)
    Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.MaximizeBox = False
    Me.Name = "frmGenMMerge"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    Me.Text = "frmGenMMerge"
    Me.SplitContainer1.Panel1.ResumeLayout(False)
    Me.SplitContainer1.Panel2.ResumeLayout(False)
    CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
    Me.SplitContainer1.ResumeLayout(False)
    Me.gbpLabels.ResumeLayout(False)
    Me.gbx.ResumeLayout(False)
    Me.gbx.PerformLayout()
    Me.ButtonPanel1.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
  Friend WithEvents ButtonPanel1 As CDBNETCL.ButtonPanel
  Friend WithEvents cmdOK As System.Windows.Forms.Button
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents gbx As System.Windows.Forms.GroupBox
  Friend WithEvents optMergerAll As System.Windows.Forms.RadioButton
  Friend WithEvents optAOnly As System.Windows.Forms.RadioButton
  Friend WithEvents optAAndC As System.Windows.Forms.RadioButton
  Friend WithEvents optCOnly As System.Windows.Forms.RadioButton
  Friend WithEvents optBOnly As System.Windows.Forms.RadioButton
  Friend WithEvents gbpLabels As System.Windows.Forms.GroupBox
  Friend WithEvents lbl3 As System.Windows.Forms.Label
  Friend WithEvents lbl2 As System.Windows.Forms.Label
  Friend WithEvents lbl1 As System.Windows.Forms.Label
  Friend WithEvents lblDesc As System.Windows.Forms.Label
  Friend WithEvents lblC As System.Windows.Forms.Label
  Friend WithEvents lblB As System.Windows.Forms.Label
  Friend WithEvents lblA As System.Windows.Forms.Label
  Friend WithEvents Label2 As System.Windows.Forms.Label
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents Panel1 As System.Windows.Forms.Panel
  Friend WithEvents Panel2 As System.Windows.Forms.Panel
End Class
