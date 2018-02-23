<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTableEntry
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmTableEntry))
    Me.bpl = New CDBNETCL.ButtonPanel()
    Me.cmdHTMLEditor = New System.Windows.Forms.Button()
    Me.cmdAddMore = New System.Windows.Forms.Button()
    Me.cmdOK = New System.Windows.Forms.Button()
    Me.cmdCancel = New System.Windows.Forms.Button()
    Me.spc = New System.Windows.Forms.SplitContainer()
    Me.epl = New CDBNETCL.EditPanel()
    Me.txtNotes = New System.Windows.Forms.TextBox()
    Me.bpl.SuspendLayout()
    CType(Me.spc, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.spc.Panel1.SuspendLayout()
    Me.spc.Panel2.SuspendLayout()
    Me.spc.SuspendLayout()
    Me.SuspendLayout()
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdHTMLEditor)
    Me.bpl.Controls.Add(Me.cmdAddMore)
    Me.bpl.Controls.Add(Me.cmdOK)
    Me.bpl.Controls.Add(Me.cmdCancel)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bpl.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bpl.Location = New System.Drawing.Point(0, 184)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(637, 39)
    Me.bpl.TabIndex = 1
    '
    'cmdHTMLEditor
    '
    Me.cmdHTMLEditor.Location = New System.Drawing.Point(104, 6)
    Me.cmdHTMLEditor.Name = "cmdHTMLEditor"
    Me.cmdHTMLEditor.Size = New System.Drawing.Size(96, 27)
    Me.cmdHTMLEditor.TabIndex = 1
    Me.cmdHTMLEditor.Text = "&HTML Editor"
    Me.cmdHTMLEditor.UseVisualStyleBackColor = True
    Me.cmdHTMLEditor.Visible = False
    '
    'cmdAddMore
    '
    Me.cmdAddMore.Location = New System.Drawing.Point(215, 6)
    Me.cmdAddMore.Name = "cmdAddMore"
    Me.cmdAddMore.Size = New System.Drawing.Size(96, 27)
    Me.cmdAddMore.TabIndex = 2
    Me.cmdAddMore.Text = "&Add More"
    Me.cmdAddMore.UseVisualStyleBackColor = True
    '
    'cmdOK
    '
    Me.cmdOK.Location = New System.Drawing.Point(326, 6)
    Me.cmdOK.Name = "cmdOK"
    Me.cmdOK.Size = New System.Drawing.Size(96, 27)
    Me.cmdOK.TabIndex = 3
    Me.cmdOK.Text = "OK"
    Me.cmdOK.UseVisualStyleBackColor = True
    '
    'cmdCancel
    '
    Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdCancel.Location = New System.Drawing.Point(437, 6)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(96, 27)
    Me.cmdCancel.TabIndex = 4
    Me.cmdCancel.Text = "Cancel"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'spc
    '
    Me.spc.Dock = System.Windows.Forms.DockStyle.Fill
    Me.spc.IsSplitterFixed = True
    Me.spc.Location = New System.Drawing.Point(0, 0)
    Me.spc.Name = "spc"
    '
    'spc.Panel1
    '
    Me.spc.Panel1.Controls.Add(Me.epl)
    '
    'spc.Panel2
    '
    Me.spc.Panel2.Controls.Add(Me.txtNotes)
    Me.spc.Size = New System.Drawing.Size(637, 184)
    Me.spc.SplitterDistance = 494
    Me.spc.TabIndex = 4
    '
    'epl
    '
    Me.epl.AddressChanged = False
    Me.epl.BackColor = System.Drawing.Color.Transparent
    Me.epl.DataChanged = False
    Me.epl.Dock = System.Windows.Forms.DockStyle.Fill
    Me.epl.Location = New System.Drawing.Point(0, 0)
    Me.epl.Name = "epl"
    Me.epl.Recipients = Nothing
    Me.epl.Size = New System.Drawing.Size(494, 184)
    Me.epl.SuppressDrawing = False
    Me.epl.TabIndex = 0
    Me.epl.TabSelectedIndex = 0
    '
    'txtNotes
    '
    Me.txtNotes.BackColor = System.Drawing.SystemColors.Window
    Me.txtNotes.Dock = System.Windows.Forms.DockStyle.Fill
    Me.txtNotes.Location = New System.Drawing.Point(0, 0)
    Me.txtNotes.Multiline = True
    Me.txtNotes.Name = "txtNotes"
    Me.txtNotes.ReadOnly = True
    Me.txtNotes.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
    Me.txtNotes.Size = New System.Drawing.Size(139, 184)
    Me.txtNotes.TabIndex = 0
    Me.txtNotes.TabStop = False
    '
    'frmTableEntry
    '
    Me.AcceptButton = Me.cmdOK
    Me.AutoSize = True
    Me.CancelButton = Me.cmdCancel
    Me.ClientSize = New System.Drawing.Size(637, 223)
    Me.Controls.Add(Me.spc)
    Me.Controls.Add(Me.bpl)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmTableEntry"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.bpl.ResumeLayout(False)
    Me.spc.Panel1.ResumeLayout(False)
    Me.spc.Panel2.ResumeLayout(False)
    Me.spc.Panel2.PerformLayout()
    CType(Me.spc, System.ComponentModel.ISupportInitialize).EndInit()
    Me.spc.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents bpl As CDBNETCL.ButtonPanel
  Friend WithEvents cmdAddMore As System.Windows.Forms.Button
  Friend WithEvents cmdOK As System.Windows.Forms.Button
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents spc As System.Windows.Forms.SplitContainer
  Friend WithEvents epl As CDBNETCL.EditPanel
  Friend WithEvents txtNotes As System.Windows.Forms.TextBox
  Friend WithEvents cmdHTMLEditor As System.Windows.Forms.Button

End Class
