<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class RelationshipDataSheet
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
    Dim TipAppearance1 As FarPoint.Win.Spread.TipAppearance = New FarPoint.Win.Spread.TipAppearance
    Me.vas = New FarPoint.Win.Spread.FpSpread
    Me.vas_Sheet1 = New FarPoint.Win.Spread.SheetView
    CType(Me.vas, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.vas_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SuspendLayout()
    '
    'vas
    '
    Me.vas.BackColor = System.Drawing.SystemColors.Control
    Me.vas.Dock = System.Windows.Forms.DockStyle.Fill
    Me.vas.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded
    Me.vas.Location = New System.Drawing.Point(0, 0)
    Me.vas.Name = "vas"
    Me.vas.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.vas_Sheet1})
    Me.vas.Size = New System.Drawing.Size(150, 150)
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
    Me.vas_Sheet1.ColumnCount = 14
    Me.vas_Sheet1.RowCount = 1
    Me.vas_Sheet1.ColumnHeader.AutoText = FarPoint.Win.Spread.HeaderAutoText.Blank
    Me.vas_Sheet1.ColumnHeader.Visible = False
    Me.vas_Sheet1.RowHeader.AutoText = FarPoint.Win.Spread.HeaderAutoText.Blank
    Me.vas_Sheet1.RowHeader.Visible = False
    Me.vas_Sheet1.SelectionPolicy = FarPoint.Win.Spread.Model.SelectionPolicy.[Single]
    Me.vas_Sheet1.SheetName = "Sheet"
    Me.vas_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
    '
    'RelationshipDataSheet
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.Controls.Add(Me.vas)
    Me.Name = "RelationshipDataSheet"
    CType(Me.vas, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.vas_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents vas As FarPoint.Win.Spread.FpSpread
  Friend WithEvents vas_Sheet1 As FarPoint.Win.Spread.SheetView

End Class
