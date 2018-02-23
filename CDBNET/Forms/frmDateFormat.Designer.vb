<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDateFormat
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDateFormat))
    Me.grpFormat = New System.Windows.Forms.GroupBox()
    Me.lblDateFormat = New System.Windows.Forms.Label()
    Me.optAS400Format = New System.Windows.Forms.RadioButton()
    Me.optDefFormat = New System.Windows.Forms.RadioButton()
    Me.grpOrder = New System.Windows.Forms.GroupBox()
    Me.optOrderYMD = New System.Windows.Forms.RadioButton()
    Me.optOrderMDY = New System.Windows.Forms.RadioButton()
    Me.optOrderDMY = New System.Windows.Forms.RadioButton()
    Me.grpYear = New System.Windows.Forms.GroupBox()
    Me.txtSeperator = New System.Windows.Forms.TextBox()
    Me.lblSeperator = New System.Windows.Forms.Label()
    Me.optYearYYYY = New System.Windows.Forms.RadioButton()
    Me.optYearYY = New System.Windows.Forms.RadioButton()
    Me.bpl = New CDBNETCL.ButtonPanel()
    Me.cmdOK = New System.Windows.Forms.Button()
    Me.cmdCancel = New System.Windows.Forms.Button()
    Me.grpMMYY = New System.Windows.Forms.GroupBox()
    Me.optYM = New System.Windows.Forms.RadioButton()
    Me.optMY = New System.Windows.Forms.RadioButton()
    Me.grpFormat.SuspendLayout()
    Me.grpOrder.SuspendLayout()
    Me.grpYear.SuspendLayout()
    Me.bpl.SuspendLayout()
    Me.grpMMYY.SuspendLayout()
    Me.SuspendLayout()
    '
    'grpFormat
    '
    Me.grpFormat.Controls.Add(Me.lblDateFormat)
    Me.grpFormat.Controls.Add(Me.optAS400Format)
    Me.grpFormat.Controls.Add(Me.optDefFormat)
    Me.grpFormat.Location = New System.Drawing.Point(8, 2)
    Me.grpFormat.Name = "grpFormat"
    Me.grpFormat.Size = New System.Drawing.Size(332, 89)
    Me.grpFormat.TabIndex = 0
    Me.grpFormat.TabStop = False
    '
    'lblDateFormat
    '
    Me.lblDateFormat.AutoSize = True
    Me.lblDateFormat.Location = New System.Drawing.Point(116, 55)
    Me.lblDateFormat.Name = "lblDateFormat"
    Me.lblDateFormat.Size = New System.Drawing.Size(65, 13)
    Me.lblDateFormat.TabIndex = 2
    Me.lblDateFormat.Text = "Date Format"
    '
    'optAS400Format
    '
    Me.optAS400Format.AutoSize = True
    Me.optAS400Format.Location = New System.Drawing.Point(207, 21)
    Me.optAS400Format.Name = "optAS400Format"
    Me.optAS400Format.Size = New System.Drawing.Size(92, 17)
    Me.optAS400Format.TabIndex = 1
    Me.optAS400Format.Text = "AS400 Format"
    Me.optAS400Format.UseVisualStyleBackColor = True
    '
    'optDefFormat
    '
    Me.optDefFormat.AutoSize = True
    Me.optDefFormat.Checked = True
    Me.optDefFormat.Location = New System.Drawing.Point(20, 21)
    Me.optDefFormat.Name = "optDefFormat"
    Me.optDefFormat.Size = New System.Drawing.Size(91, 17)
    Me.optDefFormat.TabIndex = 0
    Me.optDefFormat.TabStop = True
    Me.optDefFormat.Text = "Define Format"
    Me.optDefFormat.UseVisualStyleBackColor = True
    '
    'grpOrder
    '
    Me.grpOrder.Controls.Add(Me.optOrderYMD)
    Me.grpOrder.Controls.Add(Me.optOrderMDY)
    Me.grpOrder.Controls.Add(Me.optOrderDMY)
    Me.grpOrder.Location = New System.Drawing.Point(8, 92)
    Me.grpOrder.Name = "grpOrder"
    Me.grpOrder.Size = New System.Drawing.Size(332, 58)
    Me.grpOrder.TabIndex = 1
    Me.grpOrder.TabStop = False
    '
    'optOrderYMD
    '
    Me.optOrderYMD.AutoSize = True
    Me.optOrderYMD.Location = New System.Drawing.Point(207, 21)
    Me.optOrderYMD.Name = "optOrderYMD"
    Me.optOrderYMD.Size = New System.Drawing.Size(49, 17)
    Me.optOrderYMD.TabIndex = 3
    Me.optOrderYMD.Text = "YMD"
    Me.optOrderYMD.UseVisualStyleBackColor = True
    '
    'optOrderMDY
    '
    Me.optOrderMDY.AutoSize = True
    Me.optOrderMDY.Location = New System.Drawing.Point(119, 21)
    Me.optOrderMDY.Name = "optOrderMDY"
    Me.optOrderMDY.Size = New System.Drawing.Size(49, 17)
    Me.optOrderMDY.TabIndex = 2
    Me.optOrderMDY.Text = "MDY"
    Me.optOrderMDY.UseVisualStyleBackColor = True
    '
    'optOrderDMY
    '
    Me.optOrderDMY.AutoSize = True
    Me.optOrderDMY.Checked = True
    Me.optOrderDMY.Location = New System.Drawing.Point(20, 21)
    Me.optOrderDMY.Name = "optOrderDMY"
    Me.optOrderDMY.Size = New System.Drawing.Size(49, 17)
    Me.optOrderDMY.TabIndex = 1
    Me.optOrderDMY.TabStop = True
    Me.optOrderDMY.Text = "DMY"
    Me.optOrderDMY.UseVisualStyleBackColor = True
    '
    'grpYear
    '
    Me.grpYear.Controls.Add(Me.txtSeperator)
    Me.grpYear.Controls.Add(Me.lblSeperator)
    Me.grpYear.Controls.Add(Me.optYearYYYY)
    Me.grpYear.Controls.Add(Me.optYearYY)
    Me.grpYear.Location = New System.Drawing.Point(8, 151)
    Me.grpYear.Name = "grpYear"
    Me.grpYear.Size = New System.Drawing.Size(332, 94)
    Me.grpYear.TabIndex = 1
    Me.grpYear.TabStop = False
    '
    'txtSeperator
    '
    Me.txtSeperator.Location = New System.Drawing.Point(113, 58)
    Me.txtSeperator.MaxLength = 1
    Me.txtSeperator.Name = "txtSeperator"
    Me.txtSeperator.Size = New System.Drawing.Size(21, 20)
    Me.txtSeperator.TabIndex = 5
    Me.txtSeperator.Text = "/"
    '
    'lblSeperator
    '
    Me.lblSeperator.AutoSize = True
    Me.lblSeperator.Location = New System.Drawing.Point(13, 63)
    Me.lblSeperator.Name = "lblSeperator"
    Me.lblSeperator.Size = New System.Drawing.Size(56, 13)
    Me.lblSeperator.TabIndex = 4
    Me.lblSeperator.Text = "Separator:"
    '
    'optYearYYYY
    '
    Me.optYearYYYY.AutoSize = True
    Me.optYearYYYY.Location = New System.Drawing.Point(115, 21)
    Me.optYearYYYY.Name = "optYearYYYY"
    Me.optYearYYYY.Size = New System.Drawing.Size(53, 17)
    Me.optYearYYYY.TabIndex = 3
    Me.optYearYYYY.Text = "YYYY"
    Me.optYearYYYY.UseVisualStyleBackColor = True
    '
    'optYearYY
    '
    Me.optYearYY.AutoSize = True
    Me.optYearYY.Checked = True
    Me.optYearYY.Location = New System.Drawing.Point(16, 21)
    Me.optYearYY.Name = "optYearYY"
    Me.optYearYY.Size = New System.Drawing.Size(39, 17)
    Me.optYearYY.TabIndex = 2
    Me.optYearYY.TabStop = True
    Me.optYearYY.Text = "YY"
    Me.optYearYY.UseVisualStyleBackColor = True
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdOK)
    Me.bpl.Controls.Add(Me.cmdCancel)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bpl.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bpl.Location = New System.Drawing.Point(0, 254)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(347, 39)
    Me.bpl.TabIndex = 2
    '
    'cmdOK
    '
    Me.cmdOK.DialogResult = System.Windows.Forms.DialogResult.OK
    Me.cmdOK.Location = New System.Drawing.Point(70, 6)
    Me.cmdOK.Name = "cmdOK"
    Me.cmdOK.Size = New System.Drawing.Size(96, 27)
    Me.cmdOK.TabIndex = 1
    Me.cmdOK.Text = "OK"
    Me.cmdOK.UseVisualStyleBackColor = True
    '
    'cmdCancel
    '
    Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdCancel.Location = New System.Drawing.Point(181, 6)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(96, 27)
    Me.cmdCancel.TabIndex = 0
    Me.cmdCancel.Text = "Cancel"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'grpMMYY
    '
    Me.grpMMYY.Controls.Add(Me.optYM)
    Me.grpMMYY.Controls.Add(Me.optMY)
    Me.grpMMYY.Location = New System.Drawing.Point(8, 92)
    Me.grpMMYY.Name = "grpMMYY"
    Me.grpMMYY.Size = New System.Drawing.Size(332, 58)
    Me.grpMMYY.TabIndex = 4
    Me.grpMMYY.TabStop = False
    '
    'optYM
    '
    Me.optYM.AutoSize = True
    Me.optYM.Location = New System.Drawing.Point(119, 21)
    Me.optYM.Name = "optYM"
    Me.optYM.Size = New System.Drawing.Size(41, 17)
    Me.optYM.TabIndex = 2
    Me.optYM.Text = "YM"
    Me.optYM.UseVisualStyleBackColor = True
    '
    'optMY
    '
    Me.optMY.AutoSize = True
    Me.optMY.Location = New System.Drawing.Point(20, 21)
    Me.optMY.Name = "optMY"
    Me.optMY.Size = New System.Drawing.Size(41, 17)
    Me.optMY.TabIndex = 1
    Me.optMY.Text = "MY"
    Me.optMY.UseVisualStyleBackColor = True
    '
    'frmDateFormat
    '
    Me.AcceptButton = Me.cmdOK
    Me.CancelButton = Me.cmdCancel
    Me.ClientSize = New System.Drawing.Size(347, 293)
    Me.Controls.Add(Me.bpl)
    Me.Controls.Add(Me.grpYear)
    Me.Controls.Add(Me.grpOrder)
    Me.Controls.Add(Me.grpFormat)
    Me.Controls.Add(Me.grpMMYY)
    Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.MaximizeBox = False
    Me.Name = "frmDateFormat"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    Me.Text = "Select Date Format"
    Me.grpFormat.ResumeLayout(False)
    Me.grpFormat.PerformLayout()
    Me.grpOrder.ResumeLayout(False)
    Me.grpOrder.PerformLayout()
    Me.grpYear.ResumeLayout(False)
    Me.grpYear.PerformLayout()
    Me.bpl.ResumeLayout(False)
    Me.grpMMYY.ResumeLayout(False)
    Me.grpMMYY.PerformLayout()
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents grpFormat As System.Windows.Forms.GroupBox
  Friend WithEvents grpOrder As System.Windows.Forms.GroupBox
  Friend WithEvents grpYear As System.Windows.Forms.GroupBox
  Friend WithEvents lblDateFormat As System.Windows.Forms.Label
  Friend WithEvents optAS400Format As System.Windows.Forms.RadioButton
  Friend WithEvents optDefFormat As System.Windows.Forms.RadioButton
  Friend WithEvents optOrderYMD As System.Windows.Forms.RadioButton
  Friend WithEvents optOrderMDY As System.Windows.Forms.RadioButton
  Friend WithEvents optOrderDMY As System.Windows.Forms.RadioButton
  Friend WithEvents txtSeperator As System.Windows.Forms.TextBox
  Friend WithEvents lblSeperator As System.Windows.Forms.Label
  Friend WithEvents optYearYYYY As System.Windows.Forms.RadioButton
  Friend WithEvents optYearYY As System.Windows.Forms.RadioButton
  Friend WithEvents bpl As CDBNETCL.ButtonPanel
  Friend WithEvents cmdOK As System.Windows.Forms.Button
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents grpMMYY As System.Windows.Forms.GroupBox
  Friend WithEvents optYM As System.Windows.Forms.RadioButton
  Friend WithEvents optMY As System.Windows.Forms.RadioButton

End Class
