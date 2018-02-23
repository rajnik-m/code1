<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDataSheet
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
    Me.components = New System.ComponentModel.Container
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDataSheet))
    Me.erp = New System.Windows.Forms.ErrorProvider(Me.components)
    Me.cmdOK = New System.Windows.Forms.Button
    Me.cmdCancel = New System.Windows.Forms.Button
    Me.pnlDataSheet = New CDBNETCL.PanelEx
    Me.rds = New CDBNET.RelationshipDataSheet
    Me.ads = New CDBNET.ActivityDataSheet
    Me.pnlSource = New CDBNETCL.PanelEx
    Me.txtSource = New CDBNETCL.TextLookupBox
    Me.lblSource = New CDBNET.LabelEx
    Me.bpl = New CDBNETCL.ButtonPanel
    CType(Me.erp, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.pnlDataSheet.SuspendLayout()
    Me.pnlSource.SuspendLayout()
    Me.bpl.SuspendLayout()
    Me.SuspendLayout()
    '
    'erp
    '
    Me.erp.ContainerControl = Me
    '
    'cmdOK
    '
    Me.cmdOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdOK.Location = New System.Drawing.Point(248, 6)
    Me.cmdOK.Name = "cmdOK"
    Me.cmdOK.Size = New System.Drawing.Size(94, 27)
    Me.cmdOK.TabIndex = 0
    Me.cmdOK.Text = "OK"
    '
    'cmdCancel
    '
    Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdCancel.Location = New System.Drawing.Point(357, 6)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(94, 27)
    Me.cmdCancel.TabIndex = 1
    Me.cmdCancel.Text = "Cancel"
    '
    'pnlDataSheet
    '
    Me.pnlDataSheet.BackColor = System.Drawing.Color.Transparent
    Me.pnlDataSheet.Controls.Add(Me.rds)
    Me.pnlDataSheet.Controls.Add(Me.ads)
    Me.pnlDataSheet.Dock = System.Windows.Forms.DockStyle.Fill
    Me.pnlDataSheet.Location = New System.Drawing.Point(0, 44)
    Me.pnlDataSheet.Name = "pnlDataSheet"
    Me.pnlDataSheet.Padding = New System.Windows.Forms.Padding(8)
    Me.pnlDataSheet.Size = New System.Drawing.Size(700, 487)
    Me.pnlDataSheet.TabIndex = 1
    '
    'rds
    '
    Me.rds.Dock = System.Windows.Forms.DockStyle.Top
    Me.rds.Location = New System.Drawing.Point(8, 182)
    Me.rds.Name = "rds"
    Me.rds.Size = New System.Drawing.Size(684, 277)
    Me.rds.TabIndex = 1
    '
    'ads
    '
    Me.ads.BackColor = System.Drawing.Color.Transparent
    Me.ads.Dock = System.Windows.Forms.DockStyle.Top
    Me.ads.Location = New System.Drawing.Point(8, 8)
    Me.ads.Name = "ads"
    Me.ads.Size = New System.Drawing.Size(684, 174)
    Me.ads.TabIndex = 0
    '
    'pnlSource
    '
    Me.pnlSource.BackColor = System.Drawing.Color.Transparent
    Me.pnlSource.Controls.Add(Me.txtSource)
    Me.pnlSource.Controls.Add(Me.lblSource)
    Me.pnlSource.Dock = System.Windows.Forms.DockStyle.Top
    Me.pnlSource.Location = New System.Drawing.Point(0, 0)
    Me.pnlSource.Name = "pnlSource"
    Me.pnlSource.Size = New System.Drawing.Size(700, 44)
    Me.pnlSource.TabIndex = 0
    '
    'txtSource
    '
    Me.txtSource.BackColor = System.Drawing.Color.Transparent
    Me.txtSource.Location = New System.Drawing.Point(95, 11)
    Me.txtSource.MaxLength = 32767
    Me.txtSource.Name = "txtSource"
    Me.txtSource.OriginalText = Nothing
    Me.txtSource.Size = New System.Drawing.Size(496, 24)
    Me.txtSource.TabIndex = 2
    Me.txtSource.TotalWidth = 408
    '
    'lblSource
    '
    Me.lblSource.AutoSize = True
    Me.lblSource.BackColor = System.Drawing.Color.Transparent
    Me.lblSource.Location = New System.Drawing.Point(26, 14)
    Me.lblSource.Name = "lblSource"
    Me.lblSource.Size = New System.Drawing.Size(53, 17)
    Me.lblSource.TabIndex = 1
    Me.lblSource.Text = "Source"
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdOK)
    Me.bpl.Controls.Add(Me.cmdCancel)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bpl.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bpl.Location = New System.Drawing.Point(0, 531)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(700, 39)
    Me.bpl.TabIndex = 2
    '
    'frmDataSheet
    '
    Me.AcceptButton = Me.cmdOK
    Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.cmdCancel
    Me.ClientSize = New System.Drawing.Size(700, 570)
    Me.Controls.Add(Me.pnlDataSheet)
    Me.Controls.Add(Me.pnlSource)
    Me.Controls.Add(Me.bpl)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmDataSheet"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.Text = "Data Sheet"
    CType(Me.erp, System.ComponentModel.ISupportInitialize).EndInit()
    Me.pnlDataSheet.ResumeLayout(False)
    Me.pnlSource.ResumeLayout(False)
    Me.pnlSource.PerformLayout()
    Me.bpl.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents bpl As CDBNETCL.ButtonPanel
  Friend WithEvents pnlSource As CDBNETCL.PanelEx
  Friend WithEvents lblSource As CDBNET.LabelEx
  Friend WithEvents pnlDataSheet As CDBNETCL.PanelEx
  Friend WithEvents cmdOK As System.Windows.Forms.Button
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents ads As CDBNET.ActivityDataSheet
  Friend WithEvents rds As CDBNET.RelationshipDataSheet
  Friend WithEvents txtSource As CDBNETCL.TextLookupBox
  Friend WithEvents erp As System.Windows.Forms.ErrorProvider
End Class
