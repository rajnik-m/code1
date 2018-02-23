<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmActionSchedule
  Inherits MaintenanceParentForm

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
    Dim TipAppearance1 As FarPoint.Win.Spread.TipAppearance = New FarPoint.Win.Spread.TipAppearance
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmActionSchedule))
    Me.bpl = New CDBNETCL.ButtonPanel
    Me.cmdClose = New System.Windows.Forms.Button
    Me.tab = New CDBNETCL.TabControl
    Me.tbpLeft = New System.Windows.Forms.TabPage
    Me.tbpDay = New System.Windows.Forms.TabPage
    Me.tbpWeek = New System.Windows.Forms.TabPage
    Me.tbpMonth = New System.Windows.Forms.TabPage
    Me.tbpQuarter = New System.Windows.Forms.TabPage
    Me.tbpYear = New System.Windows.Forms.TabPage
    Me.tbpRight = New System.Windows.Forms.TabPage
    Me.vas = New FarPoint.Win.Spread.FpSpread
    Me.vas_Sheet1 = New FarPoint.Win.Spread.SheetView
    Me.bpl.SuspendLayout()
    Me.tab.SuspendLayout()
    CType(Me.vas, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.vas_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SuspendLayout()
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdClose)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bpl.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bpl.Location = New System.Drawing.Point(0, 361)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(800, 39)
    Me.bpl.TabIndex = 0
    '
    'cmdClose
    '
    Me.cmdClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdClose.Location = New System.Drawing.Point(352, 6)
    Me.cmdClose.Name = "cmdClose"
    Me.cmdClose.Size = New System.Drawing.Size(96, 27)
    Me.cmdClose.TabIndex = 0
    Me.cmdClose.Text = "Close"
    Me.cmdClose.UseVisualStyleBackColor = True
    '
    'tab
    '
    Me.tab.Controls.Add(Me.tbpLeft)
    Me.tab.Controls.Add(Me.tbpDay)
    Me.tab.Controls.Add(Me.tbpWeek)
    Me.tab.Controls.Add(Me.tbpMonth)
    Me.tab.Controls.Add(Me.tbpQuarter)
    Me.tab.Controls.Add(Me.tbpYear)
    Me.tab.Controls.Add(Me.tbpRight)
    Me.tab.Dock = System.Windows.Forms.DockStyle.Top
    Me.tab.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!)
    Me.tab.ItemSize = New System.Drawing.Size(78, 21)
    Me.tab.Location = New System.Drawing.Point(0, 0)
    Me.tab.Name = "tab"
    Me.tab.SelectedIndex = 0
    Me.tab.Size = New System.Drawing.Size(800, 30)
    Me.tab.SizeMode = System.Windows.Forms.TabSizeMode.Fixed
    Me.tab.TabIndex = 1
    '
    'tbpLeft
    '
    Me.tbpLeft.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbpLeft.Location = New System.Drawing.Point(4, 25)
    Me.tbpLeft.Name = "tbpLeft"
    Me.tbpLeft.Size = New System.Drawing.Size(792, 1)
    Me.tbpLeft.TabIndex = 5
    Me.tbpLeft.Text = "<<"
    Me.tbpLeft.UseVisualStyleBackColor = True
    '
    'tbpDay
    '
    Me.tbpDay.Location = New System.Drawing.Point(4, 26)
    Me.tbpDay.Name = "tbpDay"
    Me.tbpDay.Padding = New System.Windows.Forms.Padding(3)
    Me.tbpDay.Size = New System.Drawing.Size(192, 70)
    Me.tbpDay.TabIndex = 0
    Me.tbpDay.Text = "Day"
    Me.tbpDay.UseVisualStyleBackColor = True
    '
    'tbpWeek
    '
    Me.tbpWeek.Location = New System.Drawing.Point(4, 26)
    Me.tbpWeek.Name = "tbpWeek"
    Me.tbpWeek.Padding = New System.Windows.Forms.Padding(3)
    Me.tbpWeek.Size = New System.Drawing.Size(192, 70)
    Me.tbpWeek.TabIndex = 1
    Me.tbpWeek.Text = "Week"
    Me.tbpWeek.UseVisualStyleBackColor = True
    '
    'tbpMonth
    '
    Me.tbpMonth.Location = New System.Drawing.Point(4, 26)
    Me.tbpMonth.Name = "tbpMonth"
    Me.tbpMonth.Size = New System.Drawing.Size(192, 70)
    Me.tbpMonth.TabIndex = 2
    Me.tbpMonth.Text = "Month"
    Me.tbpMonth.UseVisualStyleBackColor = True
    '
    'tbpQuarter
    '
    Me.tbpQuarter.Location = New System.Drawing.Point(4, 26)
    Me.tbpQuarter.Name = "tbpQuarter"
    Me.tbpQuarter.Size = New System.Drawing.Size(192, 70)
    Me.tbpQuarter.TabIndex = 3
    Me.tbpQuarter.Text = "Quarter"
    Me.tbpQuarter.UseVisualStyleBackColor = True
    '
    'tbpYear
    '
    Me.tbpYear.Location = New System.Drawing.Point(4, 26)
    Me.tbpYear.Name = "tbpYear"
    Me.tbpYear.Size = New System.Drawing.Size(192, 70)
    Me.tbpYear.TabIndex = 4
    Me.tbpYear.Text = "Year"
    Me.tbpYear.UseVisualStyleBackColor = True
    '
    'tbpRight
    '
    Me.tbpRight.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbpRight.Location = New System.Drawing.Point(4, 26)
    Me.tbpRight.Name = "tbpRight"
    Me.tbpRight.Size = New System.Drawing.Size(192, 70)
    Me.tbpRight.TabIndex = 6
    Me.tbpRight.Text = ">>"
    Me.tbpRight.UseVisualStyleBackColor = True
    '
    'vas
    '
    Me.vas.BackColor = System.Drawing.SystemColors.Control
    Me.vas.Dock = System.Windows.Forms.DockStyle.Fill
    Me.vas.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded
    Me.vas.Location = New System.Drawing.Point(0, 30)
    Me.vas.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
    Me.vas.Name = "vas"
    Me.vas.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.vas_Sheet1})
    Me.vas.Size = New System.Drawing.Size(800, 331)
    Me.vas.TabIndex = 2
    TipAppearance1.BackColor = System.Drawing.SystemColors.Info
    TipAppearance1.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    TipAppearance1.ForeColor = System.Drawing.SystemColors.InfoText
    Me.vas.TextTipAppearance = TipAppearance1
    Me.vas.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded
    '
    'vas_Sheet1
    '
    Me.vas_Sheet1.Reset()
    'Formulas and custom names must be loaded with R1C1 reference style
    Me.vas_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
    Me.vas_Sheet1.ColumnCount = 6
    Me.vas_Sheet1.RowCount = 48
    Me.vas_Sheet1.AutoGenerateColumns = False
    Me.vas_Sheet1.SelectionPolicy = FarPoint.Win.Spread.Model.SelectionPolicy.[Single]
    Me.vas_Sheet1.SheetName = "Sheet1"
    Me.vas_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
    '
    'frmActionSchedule
    '
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
    Me.CancelButton = Me.cmdClose
    Me.ClientSize = New System.Drawing.Size(800, 400)
    Me.Controls.Add(Me.vas)
    Me.Controls.Add(Me.tab)
    Me.Controls.Add(Me.bpl)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmActionSchedule"
    Me.Text = "Action Schedule"
    Me.bpl.ResumeLayout(False)
    Me.tab.ResumeLayout(False)
    CType(Me.vas, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.vas_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents bpl As CDBNETCL.ButtonPanel
  Friend WithEvents cmdClose As System.Windows.Forms.Button
  Friend WithEvents tab As CDBNETCL.TabControl
  Friend WithEvents tbpDay As System.Windows.Forms.TabPage
  Friend WithEvents tbpWeek As System.Windows.Forms.TabPage
  Friend WithEvents vas As FarPoint.Win.Spread.FpSpread
  Friend WithEvents vas_Sheet1 As FarPoint.Win.Spread.SheetView
  Friend WithEvents tbpMonth As System.Windows.Forms.TabPage
  Friend WithEvents tbpQuarter As System.Windows.Forms.TabPage
  Friend WithEvents tbpYear As System.Windows.Forms.TabPage
  Friend WithEvents tbpLeft As System.Windows.Forms.TabPage
  Friend WithEvents tbpRight As System.Windows.Forms.TabPage
End Class
