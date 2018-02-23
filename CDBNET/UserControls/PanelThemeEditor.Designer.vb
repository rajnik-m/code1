<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PanelThemeEditor
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
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
    Me.cboCurvature = New System.Windows.Forms.ComboBox()
    Me.lblCurvature = New System.Windows.Forms.Label()
    Me.cboCurveMode = New System.Windows.Forms.ComboBox()
    Me.lblCurveMode = New System.Windows.Forms.Label()
    Me.cboBorderStyle = New System.Windows.Forms.ComboBox()
    Me.lblBorderStyle = New System.Windows.Forms.Label()
    Me.lblGradientMode = New System.Windows.Forms.Label()
    Me.cboGradientMode = New System.Windows.Forms.ComboBox()
    Me.lblColor1 = New System.Windows.Forms.Label()
    Me.lblColor2 = New System.Windows.Forms.Label()
    Me.cs2 = New CDBNETCL.ColorSelector()
    Me.cs1 = New CDBNETCL.ColorSelector()
    Me.SuspendLayout()
    '
    'cboCurvature
    '
    Me.cboCurvature.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cboCurvature.FormattingEnabled = True
    Me.cboCurvature.Location = New System.Drawing.Point(459, 63)
    Me.cboCurvature.Margin = New System.Windows.Forms.Padding(2)
    Me.cboCurvature.Name = "cboCurvature"
    Me.cboCurvature.Size = New System.Drawing.Size(98, 21)
    Me.cboCurvature.TabIndex = 9
    '
    'lblCurvature
    '
    Me.lblCurvature.AutoSize = True
    Me.lblCurvature.Location = New System.Drawing.Point(352, 66)
    Me.lblCurvature.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
    Me.lblCurvature.Name = "lblCurvature"
    Me.lblCurvature.Size = New System.Drawing.Size(53, 13)
    Me.lblCurvature.TabIndex = 8
    Me.lblCurvature.Text = "Curvature"
    '
    'cboCurveMode
    '
    Me.cboCurveMode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cboCurveMode.FormattingEnabled = True
    Me.cboCurveMode.Location = New System.Drawing.Point(98, 63)
    Me.cboCurveMode.Margin = New System.Windows.Forms.Padding(2)
    Me.cboCurveMode.Name = "cboCurveMode"
    Me.cboCurveMode.Size = New System.Drawing.Size(98, 21)
    Me.cboCurveMode.TabIndex = 7
    '
    'lblCurveMode
    '
    Me.lblCurveMode.AutoSize = True
    Me.lblCurveMode.Location = New System.Drawing.Point(4, 63)
    Me.lblCurveMode.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
    Me.lblCurveMode.Name = "lblCurveMode"
    Me.lblCurveMode.Size = New System.Drawing.Size(65, 13)
    Me.lblCurveMode.TabIndex = 6
    Me.lblCurveMode.Text = "Curve Mode"
    '
    'cboBorderStyle
    '
    Me.cboBorderStyle.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cboBorderStyle.FormattingEnabled = True
    Me.cboBorderStyle.Location = New System.Drawing.Point(98, 34)
    Me.cboBorderStyle.Margin = New System.Windows.Forms.Padding(2)
    Me.cboBorderStyle.Name = "cboBorderStyle"
    Me.cboBorderStyle.Size = New System.Drawing.Size(98, 21)
    Me.cboBorderStyle.TabIndex = 3
    '
    'lblBorderStyle
    '
    Me.lblBorderStyle.AutoSize = True
    Me.lblBorderStyle.Location = New System.Drawing.Point(4, 34)
    Me.lblBorderStyle.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
    Me.lblBorderStyle.Name = "lblBorderStyle"
    Me.lblBorderStyle.Size = New System.Drawing.Size(64, 13)
    Me.lblBorderStyle.TabIndex = 2
    Me.lblBorderStyle.Text = "Border Style"
    '
    'lblGradientMode
    '
    Me.lblGradientMode.AutoSize = True
    Me.lblGradientMode.Location = New System.Drawing.Point(352, 37)
    Me.lblGradientMode.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
    Me.lblGradientMode.Name = "lblGradientMode"
    Me.lblGradientMode.Size = New System.Drawing.Size(77, 13)
    Me.lblGradientMode.TabIndex = 4
    Me.lblGradientMode.Text = "Gradient Mode"
    '
    'cboGradientMode
    '
    Me.cboGradientMode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cboGradientMode.FormattingEnabled = True
    Me.cboGradientMode.Location = New System.Drawing.Point(459, 34)
    Me.cboGradientMode.Margin = New System.Windows.Forms.Padding(2)
    Me.cboGradientMode.Name = "cboGradientMode"
    Me.cboGradientMode.Size = New System.Drawing.Size(98, 21)
    Me.cboGradientMode.TabIndex = 5
    '
    'lblColor1
    '
    Me.lblColor1.AutoSize = True
    Me.lblColor1.Location = New System.Drawing.Point(4, 5)
    Me.lblColor1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
    Me.lblColor1.Name = "lblColor1"
    Me.lblColor1.Size = New System.Drawing.Size(40, 13)
    Me.lblColor1.TabIndex = 11
    Me.lblColor1.Text = "Color 1"
    '
    'lblColor2
    '
    Me.lblColor2.AutoSize = True
    Me.lblColor2.Location = New System.Drawing.Point(352, 5)
    Me.lblColor2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
    Me.lblColor2.Name = "lblColor2"
    Me.lblColor2.Size = New System.Drawing.Size(40, 13)
    Me.lblColor2.TabIndex = 12
    Me.lblColor2.Text = "Color 2"
    '
    'cs2
    '
    Me.cs2.Color = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
    Me.cs2.ColorDialog = Nothing
    Me.cs2.Location = New System.Drawing.Point(459, 5)
    Me.cs2.Name = "cs2"
    Me.cs2.RGBValue = "16777215"
    Me.cs2.Size = New System.Drawing.Size(203, 22)
    Me.cs2.SupportsTransparentColor = False
    Me.cs2.TabIndex = 1
    '
    'cs1
    '
    Me.cs1.Color = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
    Me.cs1.ColorDialog = Nothing
    Me.cs1.Location = New System.Drawing.Point(98, 5)
    Me.cs1.Name = "cs1"
    Me.cs1.RGBValue = "16777215"
    Me.cs1.Size = New System.Drawing.Size(203, 22)
    Me.cs1.SupportsTransparentColor = False
    Me.cs1.TabIndex = 0
    '
    'PanelThemeEditor
    '
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
    Me.Controls.Add(Me.cs1)
    Me.Controls.Add(Me.cs2)
    Me.Controls.Add(Me.lblColor2)
    Me.Controls.Add(Me.lblColor1)
    Me.Controls.Add(Me.cboCurvature)
    Me.Controls.Add(Me.lblCurvature)
    Me.Controls.Add(Me.cboCurveMode)
    Me.Controls.Add(Me.lblCurveMode)
    Me.Controls.Add(Me.cboBorderStyle)
    Me.Controls.Add(Me.lblBorderStyle)
    Me.Controls.Add(Me.lblGradientMode)
    Me.Controls.Add(Me.cboGradientMode)
    Me.Margin = New System.Windows.Forms.Padding(2)
    Me.Name = "PanelThemeEditor"
    Me.Size = New System.Drawing.Size(671, 89)
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents cboCurvature As System.Windows.Forms.ComboBox
  Friend WithEvents lblCurvature As System.Windows.Forms.Label
  Friend WithEvents cboCurveMode As System.Windows.Forms.ComboBox
  Friend WithEvents lblCurveMode As System.Windows.Forms.Label
  Friend WithEvents cboBorderStyle As System.Windows.Forms.ComboBox
  Friend WithEvents lblBorderStyle As System.Windows.Forms.Label
  Friend WithEvents lblGradientMode As System.Windows.Forms.Label
  Friend WithEvents cboGradientMode As System.Windows.Forms.ComboBox
  Friend WithEvents lblColor1 As System.Windows.Forms.Label
  Friend WithEvents lblColor2 As System.Windows.Forms.Label
  Friend WithEvents cs2 As CDBNETCL.ColorSelector
  Friend WithEvents cs1 As CDBNETCL.ColorSelector

End Class
